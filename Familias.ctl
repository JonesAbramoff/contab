VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Familias 
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10965
   ScaleHeight     =   6915
   ScaleWidth      =   10965
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   6120
      Index           =   3
      Left            =   180
      TabIndex        =   110
      Top             =   645
      Visible         =   0   'False
      Width           =   10590
      Begin VB.ListBox FilhosInfo 
         Columns         =   6
         Height          =   1410
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   117
         Top             =   4665
         Width           =   4155
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados dos Filhos"
         Height          =   4365
         Left            =   75
         TabIndex        =   111
         Top             =   120
         Width           =   10440
         Begin VB.TextBox FilhoEmail 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   855
            MaxLength       =   40
            TabIndex        =   205
            Text            =   "Text1"
            Top             =   390
            Width           =   2385
         End
         Begin VB.TextBox FilhoTel 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            MaxLength       =   40
            TabIndex        =   204
            Text            =   "Text1"
            Top             =   0
            Width           =   1470
         End
         Begin VB.CheckBox FilhoDataFalNoite 
            Caption         =   "Check1"
            Height          =   255
            Left            =   3585
            TabIndex        =   198
            Top             =   1830
            Width           =   735
         End
         Begin MSMask.MaskEdBox FilhoDataFal 
            Height          =   225
            Left            =   3645
            TabIndex        =   199
            Top             =   1410
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.CheckBox FilhoDataNascNoite 
            Caption         =   "Check1"
            Height          =   255
            Left            =   8445
            TabIndex        =   116
            Top             =   375
            Width           =   735
         End
         Begin MSMask.MaskEdBox FilhoDataNasc 
            Height          =   225
            Left            =   6195
            TabIndex        =   115
            Top             =   1125
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox FilhoNomeHebr 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   3960
            MaxLength       =   40
            TabIndex        =   114
            Text            =   "Text1"
            Top             =   450
            Width           =   2745
         End
         Begin VB.TextBox FilhoNome 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   360
            MaxLength       =   40
            TabIndex        =   113
            Text            =   "Text1"
            Top             =   375
            Width           =   2745
         End
         Begin MSFlexGridLib.MSFlexGrid GridFilhos 
            Height          =   3210
            Left            =   60
            TabIndex        =   112
            Top             =   225
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   5662
            _Version        =   393216
         End
      End
      Begin VB.Label NomeFilho 
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
         Left            =   4470
         TabIndex        =   197
         Top             =   4755
         Width           =   6015
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   6090
      Index           =   2
      Left            =   240
      TabIndex        =   64
      Top             =   690
      Visible         =   0   'False
      Width           =   10500
      Begin VB.ComboBox ConjSaudacao 
         Height          =   315
         ItemData        =   "Familias.ctx":0000
         Left            =   7980
         List            =   "Familias.ctx":0010
         TabIndex        =   195
         Top             =   1500
         Width           =   2415
      End
      Begin VB.ListBox ConjugeInfo 
         Columns         =   3
         Height          =   1410
         Left            =   7095
         Style           =   1  'Checkbox
         TabIndex        =   107
         Top             =   4140
         Width           =   3315
      End
      Begin VB.CheckBox ConjugeDtFalecNoite 
         Caption         =   "À Noite"
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
         Left            =   5895
         TabIndex        =   106
         Top             =   1125
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dados da Mãe"
         Height          =   1605
         Left            =   0
         TabIndex        =   82
         Top             =   4005
         Width           =   6990
         Begin VB.CheckBox ConjugeDtFalecMaeNoite 
            Caption         =   "À Noite"
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
            Left            =   5925
            TabIndex        =   102
            Top             =   1170
            Width           =   975
         End
         Begin VB.CheckBox ConjugeDtNascMaeNoite 
            Caption         =   "À Noite"
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
            Left            =   2310
            TabIndex        =   101
            Top             =   1170
            Width           =   975
         End
         Begin MSMask.MaskEdBox ConjugeMae 
            Height          =   315
            Left            =   1725
            TabIndex        =   83
            Top             =   240
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ConjugeMaeHebr 
            Height          =   315
            Left            =   1725
            TabIndex        =   84
            Top             =   690
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ConjugeDtNascMae 
            Height          =   315
            Left            =   870
            TabIndex        =   95
            Top             =   1110
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownConjugeDtNascMae 
            Height          =   300
            Left            =   1995
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   1125
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ConjugeDtFalecMae 
            Height          =   315
            Left            =   4470
            TabIndex        =   97
            Top             =   1110
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownConjugeDtFalecMae 
            Height          =   300
            Left            =   5595
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   1125
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelConjugeDtFalecMae 
            Alignment       =   1  'Right Justify
            Caption         =   "Falec.:"
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
            Left            =   3795
            TabIndex        =   100
            Top             =   1155
            Width           =   585
         End
         Begin VB.Label LabelConjugeDtNascMae 
            Alignment       =   1  'Right Justify
            Caption         =   "Nasc.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   99
            Top             =   1140
            Width           =   645
         End
         Begin VB.Label LabelConjugeMae 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   990
            TabIndex        =   86
            Top             =   270
            Width           =   615
         End
         Begin VB.Label LabelConjugeMaeHebr 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome Hebraico:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   105
            TabIndex        =   85
            Top             =   720
            Width           =   1500
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dados do Pai"
         Height          =   1620
         Left            =   0
         TabIndex        =   77
         Top             =   2325
         Width           =   6975
         Begin VB.CheckBox ConjugeDtFalecPaiNoite 
            Caption         =   "À Noite"
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
            Left            =   5925
            TabIndex        =   94
            Top             =   1185
            Width           =   975
         End
         Begin VB.CheckBox ConjugeDtNascPaiNoite 
            Caption         =   "À Noite"
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
            Left            =   2310
            TabIndex        =   90
            Top             =   1170
            Width           =   975
         End
         Begin MSMask.MaskEdBox ConjugePai 
            Height          =   315
            Left            =   1710
            TabIndex        =   78
            Top             =   225
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ConjugePaiHebr 
            Height          =   315
            Left            =   1710
            TabIndex        =   79
            Top             =   675
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ConjugeDtNascPai 
            Height          =   315
            Left            =   825
            TabIndex        =   87
            Top             =   1095
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownConjugeDtNascPai 
            Height          =   300
            Left            =   1995
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   1110
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ConjugeDtFalecPai 
            Height          =   315
            Left            =   4485
            TabIndex        =   91
            Top             =   1125
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownConjugeDtFalecPai 
            Height          =   300
            Left            =   5610
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   1140
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelConjugeDtFalecPai 
            Alignment       =   1  'Right Justify
            Caption         =   "Falec.:"
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
            Left            =   3720
            TabIndex        =   93
            Top             =   1170
            Width           =   675
         End
         Begin VB.Label LabelConjugeDtNascPai 
            Alignment       =   1  'Right Justify
            Caption         =   "Nasc.:"
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
            Left            =   135
            TabIndex        =   89
            Top             =   1125
            Width           =   645
         End
         Begin VB.Label LabelConjugePai 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   810
            TabIndex        =   81
            Top             =   255
            Width           =   780
         End
         Begin VB.Label LabelConjugePaiHebr 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome Hebraico:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   80
            Top             =   705
            Width           =   1500
         End
      End
      Begin VB.CheckBox ConjugeDtNascNoite 
         Caption         =   "À Noite"
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
         Left            =   2340
         TabIndex        =   76
         Top             =   1110
         Width           =   1065
      End
      Begin MSMask.MaskEdBox ConjugeNome 
         Height          =   315
         Left            =   1710
         TabIndex        =   65
         Top             =   150
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ConjugeNomeHebr 
         Height          =   315
         Left            =   1710
         TabIndex        =   66
         Top             =   600
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ConjugeDtNasc 
         Height          =   315
         Left            =   855
         TabIndex        =   67
         Top             =   1035
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownConjugeDtNasc 
         Height          =   300
         Left            =   2025
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1050
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox ConjugeProfissao 
         Height          =   315
         Left            =   1710
         TabIndex        =   69
         Top             =   1950
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ConjugeNomeFirma 
         Height          =   315
         Left            =   1710
         TabIndex        =   70
         Top             =   1485
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ConjugeDtFalec 
         Height          =   315
         Left            =   4455
         TabIndex        =   103
         Top             =   1065
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownConjugeDtFalec 
         Height          =   300
         Left            =   5595
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label8 
         Caption         =   "Saudação:"
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
         Left            =   7035
         TabIndex        =   196
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label LabelConjugeDtFalec 
         Alignment       =   1  'Right Justify
         Caption         =   "Falec.:"
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
         Left            =   3675
         TabIndex        =   105
         Top             =   1095
         Width           =   690
      End
      Begin VB.Label LabelConjugeNomeFirma 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome da Firma:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   75
         Top             =   1515
         Width           =   1395
      End
      Begin VB.Label LabelConjugeProfissao 
         Alignment       =   1  'Right Justify
         Caption         =   "Profissão:"
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
         Left            =   525
         TabIndex        =   74
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label LabelConjugeDtNasc 
         Alignment       =   1  'Right Justify
         Caption         =   "Nasc.:"
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
         TabIndex        =   73
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label LabelConjugeNomeHebr 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome Hebraico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   72
         Top             =   630
         Width           =   1380
      End
      Begin VB.Label LabelConjugeNome 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   900
         TabIndex        =   71
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6045
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   720
      Width           =   10575
      Begin VB.ComboBox TitSaudacao 
         Height          =   315
         ItemData        =   "Familias.ctx":002B
         Left            =   8085
         List            =   "Familias.ctx":002D
         TabIndex        =   193
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sócio"
         Height          =   1635
         Index           =   0
         Left            =   7125
         TabIndex        =   28
         Top             =   2760
         Width           =   3330
         Begin VB.ComboBox LocalCobranca 
            Height          =   315
            ItemData        =   "Familias.ctx":002F
            Left            =   1890
            List            =   "Familias.ctx":003C
            TabIndex        =   62
            Top             =   690
            Width           =   1290
         End
         Begin MSMask.MaskEdBox CodCliente 
            Height          =   315
            Left            =   1875
            TabIndex        =   60
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorContribuicao 
            Height          =   285
            Left            =   2070
            TabIndex        =   108
            Top             =   1155
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            Caption         =   "Valor de Contribuicao:"
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
            Left            =   105
            TabIndex        =   109
            Top             =   1185
            Width           =   1950
         End
         Begin VB.Label Label4 
            Caption         =   "Local de Cobrança:"
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
            Left            =   150
            TabIndex        =   63
            Top             =   750
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código de Cliente:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   61
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.ComboBox EstadoCivil 
         Height          =   315
         ItemData        =   "Familias.ctx":005A
         Left            =   8085
         List            =   "Familias.ctx":0070
         TabIndex        =   58
         Top             =   600
         Width           =   1440
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados da Mae"
         Height          =   1530
         Left            =   300
         TabIndex        =   45
         Top             =   4455
         Width           =   6690
         Begin VB.CheckBox TitularDtFalecMaeNoite 
            Caption         =   "À Noite"
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
            Left            =   5640
            TabIndex        =   57
            Top             =   1155
            Width           =   975
         End
         Begin VB.CheckBox TitularDtNascMaeNoite 
            Caption         =   "À Noite"
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
            Left            =   2160
            TabIndex        =   56
            Top             =   1170
            Width           =   975
         End
         Begin MSMask.MaskEdBox TitularMae 
            Height          =   315
            Left            =   1575
            TabIndex        =   46
            Top             =   225
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TitularMaeHebr 
            Height          =   315
            Left            =   1560
            TabIndex        =   47
            Top             =   675
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TitularDtNascMae 
            Height          =   315
            Left            =   795
            TabIndex        =   50
            Top             =   1110
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownTitularDtNascMae 
            Height          =   300
            Left            =   1845
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1125
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox TitularDtFalecMae 
            Height          =   315
            Left            =   4275
            TabIndex        =   53
            Top             =   1110
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownTitularDtFalecMae 
            Height          =   300
            Left            =   5340
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1125
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelTitularDtFalecMae 
            Alignment       =   1  'Right Justify
            Caption         =   "Falec.:"
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
            Left            =   3585
            TabIndex        =   55
            Top             =   1140
            Width           =   660
         End
         Begin VB.Label LabelTitularDtNascMae 
            Alignment       =   1  'Right Justify
            Caption         =   "Nasc.:"
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
            TabIndex        =   52
            Top             =   1140
            Width           =   645
         End
         Begin VB.Label LabelTitularMaeHebr 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome Hebraico:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Top             =   705
            Width           =   1395
         End
         Begin VB.Label LabelTitularMae 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   870
            TabIndex        =   48
            Top             =   315
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados do Pai"
         Height          =   1650
         Left            =   300
         TabIndex        =   32
         Top             =   2760
         Width           =   6690
         Begin VB.CheckBox TitularDtFalecPaiNoite 
            Caption         =   "À Noite"
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
            Left            =   5625
            TabIndex        =   44
            Top             =   1185
            Width           =   975
         End
         Begin VB.CheckBox TitularDtNascPaiNoite 
            Caption         =   "À Noite"
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
            Left            =   2175
            TabIndex        =   43
            Top             =   1215
            Width           =   975
         End
         Begin MSMask.MaskEdBox TitularPai 
            Height          =   315
            Left            =   1590
            TabIndex        =   33
            Top             =   210
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TitularPaiHebr 
            Height          =   315
            Left            =   1590
            TabIndex        =   34
            Top             =   660
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TitularDtNascPai 
            Height          =   315
            Left            =   750
            TabIndex        =   37
            Top             =   1155
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownTitularDtNascPai 
            Height          =   300
            Left            =   1860
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1155
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox TitularDtFalecPai 
            Height          =   315
            Left            =   4215
            TabIndex        =   40
            Top             =   1155
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownTitularDtFalecPai 
            Height          =   300
            Left            =   5325
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1155
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelTitularDtFalecPai 
            Alignment       =   1  'Right Justify
            Caption         =   "Falec.:"
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
            Left            =   3495
            TabIndex        =   42
            Top             =   1185
            Width           =   675
         End
         Begin VB.Label LabelTitularDtNascPai 
            Alignment       =   1  'Right Justify
            Caption         =   "Nasc.:"
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
            Left            =   90
            TabIndex        =   39
            Top             =   1185
            Width           =   615
         End
         Begin VB.Label LabelTitularPaiHebr 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome Hebraico:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   36
            Top             =   690
            Width           =   1455
         End
         Begin VB.Label LabelTitularPai 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   855
            TabIndex        =   35
            Top             =   285
            Width           =   675
         End
      End
      Begin VB.ComboBox CohenLeviIsrael 
         Height          =   315
         ItemData        =   "Familias.ctx":00A3
         Left            =   8085
         List            =   "Familias.ctx":00B3
         TabIndex        =   30
         Top             =   1500
         Width           =   1590
      End
      Begin VB.ListBox TitularInfo 
         Columns         =   3
         Height          =   1410
         Left            =   7110
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   4515
         Width           =   3315
      End
      Begin VB.CheckBox DataCasamentoNoite 
         Caption         =   "À Noite"
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
         Left            =   9540
         TabIndex        =   23
         Top             =   1125
         Width           =   1065
      End
      Begin VB.CheckBox TitularDtNascNoite 
         Caption         =   "À Noite"
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
         Left            =   2565
         TabIndex        =   19
         Top             =   1560
         Width           =   1065
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1920
         Picture         =   "Familias.ctx":00CE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Numeração Automática"
         Top             =   135
         Width           =   300
      End
      Begin MSMask.MaskEdBox CodFamilia 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Sobrenome 
         Height          =   315
         Left            =   3570
         TabIndex        =   10
         Top             =   150
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TitularNome 
         Height          =   315
         Left            =   1905
         TabIndex        =   11
         Top             =   615
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TitularNomeHebr 
         Height          =   315
         Left            =   1905
         TabIndex        =   12
         Top             =   1065
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TitularDtNasc 
         Height          =   315
         Left            =   1095
         TabIndex        =   16
         Top             =   1500
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownTitularDtNasc 
         Height          =   300
         Left            =   2250
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1500
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCasamento 
         Height          =   315
         Left            =   8085
         TabIndex        =   20
         Top             =   1050
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataCasamento 
         Height          =   300
         Left            =   9240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox TitularNomeFirma 
         Height          =   315
         Left            =   1920
         TabIndex        =   24
         Top             =   1950
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TitularProfissao 
         Height          =   315
         Left            =   1905
         TabIndex        =   25
         Top             =   2400
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Caption         =   "Saudação:"
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
         Left            =   7140
         TabIndex        =   194
         Top             =   1995
         Width           =   945
      End
      Begin VB.Label Label27 
         Caption         =   "Atualizado em:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8115
         TabIndex        =   146
         Top             =   195
         Width           =   1320
      End
      Begin VB.Label AtualizadoEm 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   9420
         TabIndex        =   145
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Estado Civil:"
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
         Left            =   6990
         TabIndex        =   59
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descendência:"
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
         Left            =   6795
         TabIndex        =   31
         Top             =   1545
         Width           =   1305
      End
      Begin VB.Label LabelTitularProfissao 
         Alignment       =   1  'Right Justify
         Caption         =   "Profissão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   285
         TabIndex        =   27
         Top             =   2430
         Width           =   1500
      End
      Begin VB.Label LabelTitularNomeFirma 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome da Firma:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   26
         Top             =   1980
         Width           =   1500
      End
      Begin VB.Label LabelDataCasamento 
         Alignment       =   1  'Right Justify
         Caption         =   "Casamento:"
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
         Left            =   6960
         TabIndex        =   22
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label LabelTitularDtNasc 
         Alignment       =   1  'Right Justify
         Caption         =   "Nasc.:"
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
         Left            =   315
         TabIndex        =   18
         Top             =   1530
         Width           =   660
      End
      Begin VB.Label LabelTitularNomeHebr 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome Hebraico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   285
         TabIndex        =   15
         Top             =   1095
         Width           =   1500
      End
      Begin VB.Label LabelTitularNome 
         Alignment       =   1  'Right Justify
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
         Left            =   1185
         TabIndex        =   14
         Top             =   630
         Width           =   600
      End
      Begin VB.Label LabelSobrenome 
         Alignment       =   1  'Right Justify
         Caption         =   "Sobrenome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2415
         TabIndex        =   13
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label LabelCodFamilia 
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   180
         Width           =   660
      End
   End
   Begin VB.CommandButton BotaoConsulta3 
      Caption         =   "Consulta3"
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
      Left            =   6600
      TabIndex        =   203
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton BotaoConsulta2 
      Caption         =   "Consulta2"
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
      Left            =   5257
      TabIndex        =   202
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton BotaoConsulta1 
      Caption         =   "Consulta1"
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
      Left            =   3915
      TabIndex        =   201
      Top             =   135
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   6090
      Index           =   4
      Left            =   165
      TabIndex        =   118
      Top             =   660
      Visible         =   0   'False
      Width           =   10545
      Begin VB.Frame Frame9 
         Caption         =   "Enderecos Comerciais"
         Height          =   3285
         Index           =   0
         Left            =   75
         TabIndex        =   120
         Top             =   2790
         Width           =   8880
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   2295
            Index           =   1
            Left            =   165
            TabIndex        =   147
            Top             =   765
            Width           =   8580
            Begin VB.ComboBox Pais 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   1
               Left            =   4020
               TabIndex        =   150
               Top             =   1020
               Width           =   1995
            End
            Begin VB.ComboBox Estado 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   149
               Top             =   1005
               Width           =   630
            End
            Begin VB.TextBox Endereco 
               Height          =   315
               Index           =   1
               Left            =   1260
               MaxLength       =   40
               TabIndex        =   148
               Top             =   120
               Width           =   6345
            End
            Begin MSMask.MaskEdBox Cidade 
               Height          =   315
               Index           =   1
               Left            =   4020
               TabIndex        =   151
               Top             =   570
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Bairro 
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   152
               Top             =   570
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CEP 
               Height          =   315
               Index           =   1
               Left            =   6660
               TabIndex        =   153
               Top             =   570
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
               Left            =   1260
               TabIndex        =   154
               Top             =   1440
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Telefone2 
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   155
               Top             =   1845
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Email 
               Height          =   315
               Index           =   1
               Left            =   4020
               TabIndex        =   156
               Top             =   1845
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   50
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Contato 
               Height          =   315
               Index           =   1
               Left            =   6660
               TabIndex        =   157
               Top             =   1845
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   50
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Fax 
               Height          =   315
               Index           =   1
               Left            =   4020
               TabIndex        =   158
               Top             =   1425
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin VB.Label PaisLabel 
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
               Index           =   1
               Left            =   3480
               TabIndex        =   169
               Top             =   1065
               Width           =   495
            End
            Begin VB.Label Label7 
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
               Index           =   0
               Left            =   5865
               TabIndex        =   168
               Top             =   1890
               Width           =   750
            End
            Begin VB.Label Label9 
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
               Index           =   0
               Left            =   6150
               TabIndex        =   167
               Top             =   645
               Width           =   465
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Celular:"
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
               Left            =   3285
               TabIndex        =   166
               Top             =   1470
               Width           =   660
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Email:"
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
               Left            =   3405
               TabIndex        =   165
               Top             =   1890
               Width           =   525
            End
            Begin VB.Label Label38 
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
               Index           =   0
               Left            =   195
               TabIndex        =   164
               Top             =   1890
               Width           =   1005
            End
            Begin VB.Label Label39 
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
               Index           =   0
               Left            =   195
               TabIndex        =   163
               Top             =   1470
               Width           =   1005
            End
            Begin VB.Label Label40 
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
               Index           =   0
               Left            =   615
               TabIndex        =   162
               Top             =   615
               Width           =   585
            End
            Begin VB.Label Label41 
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
               Index           =   0
               Left            =   525
               TabIndex        =   161
               Top             =   1050
               Width           =   675
            End
            Begin VB.Label LabelCidade 
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
               Index           =   1
               Left            =   3300
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   160
               Top             =   645
               Width           =   675
            End
            Begin VB.Label Label43 
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
               Index           =   0
               Left            =   285
               TabIndex        =   159
               Top             =   165
               Width           =   915
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   2295
            Index           =   2
            Left            =   165
            TabIndex        =   122
            Top             =   765
            Visible         =   0   'False
            Width           =   8595
            Begin VB.TextBox Endereco 
               Height          =   315
               Index           =   2
               Left            =   1260
               MaxLength       =   40
               TabIndex        =   125
               Top             =   120
               Width           =   6345
            End
            Begin VB.ComboBox Estado 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   2
               Left            =   1260
               TabIndex        =   124
               Top             =   1005
               Width           =   630
            End
            Begin VB.ComboBox Pais 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   2
               Left            =   4020
               TabIndex        =   123
               Top             =   1020
               Width           =   1995
            End
            Begin MSMask.MaskEdBox Cidade 
               Height          =   315
               Index           =   2
               Left            =   4020
               TabIndex        =   126
               Top             =   570
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Bairro 
               Height          =   315
               Index           =   2
               Left            =   1260
               TabIndex        =   127
               Top             =   570
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CEP 
               Height          =   315
               Index           =   2
               Left            =   6660
               TabIndex        =   128
               Top             =   570
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
               Index           =   2
               Left            =   1260
               TabIndex        =   129
               Top             =   1440
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Telefone2 
               Height          =   315
               Index           =   2
               Left            =   1260
               TabIndex        =   130
               Top             =   1845
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Email 
               Height          =   315
               Index           =   2
               Left            =   4020
               TabIndex        =   131
               Top             =   1845
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   50
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Contato 
               Height          =   315
               Index           =   2
               Left            =   6660
               TabIndex        =   132
               Top             =   1845
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   50
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Fax 
               Height          =   315
               Index           =   2
               Left            =   4020
               TabIndex        =   133
               Top             =   1425
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin VB.Label Label26 
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
               Left            =   285
               TabIndex        =   144
               Top             =   165
               Width           =   915
            End
            Begin VB.Label LabelCidade 
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
               Index           =   2
               Left            =   3300
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   143
               Top             =   645
               Width           =   675
            End
            Begin VB.Label Label25 
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
               Left            =   525
               TabIndex        =   142
               Top             =   1050
               Width           =   675
            End
            Begin VB.Label Label24 
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
               Left            =   615
               TabIndex        =   141
               Top             =   615
               Width           =   585
            End
            Begin VB.Label Label23 
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
               Left            =   195
               TabIndex        =   140
               Top             =   1470
               Width           =   1005
            End
            Begin VB.Label Label22 
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
               Left            =   195
               TabIndex        =   139
               Top             =   1890
               Width           =   1005
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Email:"
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
               TabIndex        =   138
               Top             =   1890
               Width           =   525
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Celular:"
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
               Left            =   3285
               TabIndex        =   137
               Top             =   1470
               Width           =   660
            End
            Begin VB.Label Label19 
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
               Left            =   6150
               TabIndex        =   136
               Top             =   645
               Width           =   465
            End
            Begin VB.Label Label18 
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
               Left            =   5865
               TabIndex        =   135
               Top             =   1890
               Width           =   750
            End
            Begin VB.Label PaisLabel 
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
               Index           =   2
               Left            =   3480
               TabIndex        =   134
               Top             =   1065
               Width           =   495
            End
         End
         Begin MSComctlLib.TabStrip TabStrip2 
            Height          =   2925
            Left            =   105
            TabIndex        =   121
            Top             =   255
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   5159
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Titular"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Conjuge"
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
      Begin VB.Frame Frame8 
         Caption         =   "Endereco Residencial"
         Height          =   2565
         Left            =   75
         TabIndex        =   119
         Top             =   60
         Width           =   8865
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   2295
            Index           =   1
            Left            =   210
            TabIndex        =   170
            Top             =   210
            Width           =   8580
            Begin VB.ComboBox Pais 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   0
               Left            =   4020
               TabIndex        =   173
               Top             =   1020
               Width           =   1995
            End
            Begin VB.ComboBox Estado 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   0
               Left            =   1260
               TabIndex        =   172
               Top             =   1005
               Width           =   630
            End
            Begin VB.TextBox Endereco 
               Height          =   315
               Index           =   0
               Left            =   1260
               MaxLength       =   40
               TabIndex        =   171
               Top             =   120
               Width           =   6345
            End
            Begin MSMask.MaskEdBox Cidade 
               Height          =   315
               Index           =   0
               Left            =   4020
               TabIndex        =   174
               Top             =   570
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Bairro 
               Height          =   315
               Index           =   0
               Left            =   1260
               TabIndex        =   175
               Top             =   570
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CEP 
               Height          =   315
               Index           =   0
               Left            =   6660
               TabIndex        =   176
               Top             =   570
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
               Left            =   1260
               TabIndex        =   177
               Top             =   1440
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Telefone2 
               Height          =   315
               Index           =   0
               Left            =   1260
               TabIndex        =   178
               Top             =   1845
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Email 
               Height          =   315
               Index           =   0
               Left            =   4005
               TabIndex        =   179
               Top             =   1845
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   50
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Contato 
               Height          =   315
               Index           =   0
               Left            =   6675
               TabIndex        =   180
               Top             =   1845
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   50
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Fax 
               Height          =   315
               Index           =   0
               Left            =   4020
               TabIndex        =   181
               Top             =   1425
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin VB.Label PaisLabel 
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
               Index           =   0
               Left            =   3480
               TabIndex        =   192
               Top             =   1065
               Width           =   495
            End
            Begin VB.Label Label7 
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
               Index           =   1
               Left            =   5865
               TabIndex        =   191
               Top             =   1890
               Width           =   750
            End
            Begin VB.Label Label9 
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
               Index           =   1
               Left            =   6150
               TabIndex        =   190
               Top             =   645
               Width           =   465
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Celular:"
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
               Left            =   3285
               TabIndex        =   189
               Top             =   1470
               Width           =   660
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Email:"
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
               Left            =   3420
               TabIndex        =   188
               Top             =   1890
               Width           =   525
            End
            Begin VB.Label Label38 
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
               Index           =   1
               Left            =   195
               TabIndex        =   187
               Top             =   1890
               Width           =   1005
            End
            Begin VB.Label Label39 
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
               Index           =   1
               Left            =   195
               TabIndex        =   186
               Top             =   1470
               Width           =   1005
            End
            Begin VB.Label Label40 
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
               Index           =   1
               Left            =   615
               TabIndex        =   185
               Top             =   615
               Width           =   585
            End
            Begin VB.Label Label41 
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
               Index           =   1
               Left            =   525
               TabIndex        =   184
               Top             =   1050
               Width           =   675
            End
            Begin VB.Label LabelCidade 
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
               Index           =   0
               Left            =   3300
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   183
               Top             =   645
               Width           =   675
            End
            Begin VB.Label Label43 
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
               Index           =   1
               Left            =   285
               TabIndex        =   182
               Top             =   165
               Width           =   915
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8250
      ScaleHeight     =   450
      ScaleWidth      =   2535
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2595
      Begin VB.CommandButton BotaoPessoas 
         Height          =   360
         Left            =   75
         Picture         =   "Familias.ctx":01B8
         Style           =   1  'Graphical
         TabIndex        =   200
         ToolTipText     =   "Consulta de Pessoas"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2040
         Picture         =   "Familias.ctx":0762
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1560
         Picture         =   "Familias.ctx":08E0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1065
         Picture         =   "Familias.ctx":0E12
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   570
         Picture         =   "Familias.ctx":0F9C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6510
      Left            =   90
      TabIndex        =   5
      Top             =   330
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   11483
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Titular"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Conjuge"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filhos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereços"
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
Attribute VB_Name = "Familias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim iFrameAtual1 As Integer
Dim iFrameAtual2 As Integer
Dim iEndereco As Integer

Dim gcolcolFilhosInfo As Collection

Dim objGridFilhos As AdmGrid
Dim iGrid_FilhoNome_Col As Integer
Dim iGrid_FilhoNomeHebr_Col As Integer
Dim iGrid_FilhoDataNasc_Col As Integer
Dim iGrid_FilhoDataNascNoite_Col As Integer
Dim iGrid_FilhoTelefone_Col As Integer
Dim iGrid_FilhoEmail_Col As Integer
Dim iGrid_FilhoDataFal_Col As Integer
Dim iGrid_FilhoDataFalNoite_Col As Integer

Private WithEvents objEventoCodFamilia As AdmEvento
Attribute objEventoCodFamilia.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1

Const TAB_TITULAR = 1
Const TAB_ENDERECO_TITULAR = 1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Famílias"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Familias"

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

Private Sub BotaoConsulta1_Click()
    Call Chama_Tela("MembrosFamiliaCompleto1Lista")
End Sub

Private Sub BotaoConsulta2_Click()
    Call Chama_Tela("MembrosFamiliaCompleto2Lista")
End Sub

Private Sub BotaoConsulta3_Click()
    Call Chama_Tela("MembrosFamiliaCompleto3Lista")
End Sub

Private Sub CohenLeviIsrael_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EstadoCivil_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LocalCobranca_Change()
    iAlterado = REGISTRO_ALTERADO
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodFamilia = Nothing
    Set objEventoCidade = Nothing
    
    Set objGridFilhos = Nothing
    
    Set gcolcolFilhosInfo = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159904)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodFamilia = New AdmEvento
    Set objEventoCidade = New AdmEvento

    Set objGridFilhos = New AdmGrid
    
    Set gcolcolFilhosInfo = New Collection
    
    'Inicializa o Grid de Filhos
    Call Inicializa_GridFilhos(objGridFilhos)
    
    'Limita tamanhos dinamicamente
    Call Inicializa_Tamanhos

    'Carrega a combo de Estados
    lErro = Carrega_Estados()
    If lErro <> SUCESSO Then gError 140960

    'Carrega a combo de países
    lErro = Carrega_Paises()
    If lErro <> SUCESSO Then gError 140961

    'Carrega os ListBox de Tipo
    lErro = Carrega_Infos()
    If lErro <> SUCESSO Then gError 140962
    
    Call Carrega_Saudacoes(TitSaudacao)
    
    Call Carrega_Saudacoes(ConjSaudacao)
    
    iFrameAtual1 = 1
    iFrameAtual2 = 1

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 140928 To 140930, 140945, 140960 To 140962

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159905)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objFamilias As ClassFamilias) As Long

Dim lErro As Long, iCodigo As Integer, objFamiliasAux As ClassFamilias
Dim objFamiliasAux1 As ClassFamilias

On Error GoTo Erro_Trata_Parametros

    If Not (objFamilias Is Nothing) Then

        lErro = Traz_Familias_Tela(objFamilias)
        If lErro <> SUCESSO Then gError 130432

    End If

'    For iCodigo = 1 To 5000
'
'        Set objFamiliasAux = New ClassFamilias
'        Set objFamiliasAux1 = New ClassFamilias
'
'        objFamiliasAux.lCodFamilia = iCodigo
'        objFamiliasAux1.lCodFamilia = iCodigo
'
'        'Lê o Familias que está sendo Passado
'        lErro = CF("Familias_Le", objFamiliasAux)
'        If lErro = SUCESSO Then
'
'            lErro = Traz_Familias_Tela(objFamiliasAux1)
'            If lErro = SUCESSO Then
'                DoEvents
'                Call BotaoGravar_Click
'                DoEvents
'                Call Limpa_Tela_Familias
'                DoEvents
'            End If
'
'        End If
'
'    Next

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 130432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159906)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objFamilias As ClassFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objFamilias.sTitularSaudacao = TitSaudacao.Text
    objFamilias.sConjugeSaudacao = ConjSaudacao.Text

    objFamilias.lCodFamilia = StrParaLong(CodFamilia.Text)
    objFamilias.sSobrenome = Sobrenome.Text
    objFamilias.sTitularNome = TitularNome.Text
    objFamilias.sTitularNomeHebr = TitularNomeHebr.Text
    objFamilias.sTitularNomeFirma = TitularNomeFirma.Text
    
    If LocalCobranca.ListIndex <> -1 Then
        objFamilias.iLocalCobranca = LocalCobranca.ItemData(LocalCobranca.ListIndex)
    End If
    
    If EstadoCivil.ListIndex <> -1 Then
        objFamilias.iEstadoCivil = EstadoCivil.ItemData(EstadoCivil.ListIndex)
    End If
    
    objFamilias.sTitularProfissao = TitularProfissao.Text
    objFamilias.dtTitularDtNasc = MaskedParaDate(TitularDtNasc)
    objFamilias.iTitularDtNascNoite = TitularDtNascNoite.Value
    objFamilias.dtDataCasamento = MaskedParaDate(DataCasamento)
    objFamilias.iDataCasamentoNoite = DataCasamentoNoite.Value
    
    If CohenLeviIsrael.ListIndex <> -1 Then
    
        Select Case CohenLeviIsrael.ItemData(CohenLeviIsrael.ListIndex)
        
            Case TRIPO_COHEN
                objFamilias.sCohenLeviIsrael = STRING_TRIPO_COHEN
        
            Case TRIPO_ISRAEL
                objFamilias.sCohenLeviIsrael = STRING_TRIPO_ISRAEL
        
            Case TRIPO_LEVI
                objFamilias.sCohenLeviIsrael = STRING_TRIPO_LEVI
        
        End Select
        
    End If
    
    objFamilias.sTitularPai = TitularPai.Text
    objFamilias.sTitularPaiHebr = TitularPaiHebr.Text
    objFamilias.sTitularMae = TitularMae.Text
    objFamilias.sTitularMaeHebr = TitularMaeHebr.Text
    objFamilias.dtTitularDtNascPai = MaskedParaDate(TitularDtNascPai)
    objFamilias.iTitularDtNascPaiNoite = TitularDtNascPaiNoite.Value
    
    objFamilias.dtTitularDtFalecPai = MaskedParaDate(TitularDtFalecPai)
    objFamilias.iTitularDtFalecPaiNoite = TitularDtFalecPaiNoite.Value
    objFamilias.dtTitularDtNascMae = MaskedParaDate(TitularDtNascMae)
    objFamilias.iTitularDtNascMaeNoite = TitularDtNascMaeNoite.Value
    objFamilias.dtTitularDtFalecMae = MaskedParaDate(TitularDtFalecMae)
    objFamilias.iTitularDtFalecMaeNoite = TitularDtFalecMaeNoite.Value
    
    objFamilias.sConjugeNome = ConjugeNome.Text
    objFamilias.sConjugeNomeHebr = ConjugeNomeHebr.Text
    objFamilias.dtConjugeDtNasc = MaskedParaDate(ConjugeDtNasc)
    objFamilias.iConjugeDtNascNoite = ConjugeDtNascNoite.Value
    objFamilias.sConjugeProfissao = ConjugeProfissao.Text
    objFamilias.sConjugeNomeFirma = ConjugeNomeFirma.Text
    
    objFamilias.sConjugePai = ConjugePai.Text
    objFamilias.sConjugePaiHebr = ConjugePaiHebr.Text
    objFamilias.sConjugeMae = ConjugeMae.Text
    objFamilias.sConjugeMaeHebr = ConjugeMaeHebr.Text
    objFamilias.dtConjugeDtNascPai = MaskedParaDate(ConjugeDtNascPai)
    objFamilias.iConjugeDtNascPaiNoite = ConjugeDtNascPaiNoite.Value
    objFamilias.dtConjugeDtFalecPai = MaskedParaDate(ConjugeDtFalecPai)
    objFamilias.iConjugeDtFalecPaiNoite = ConjugeDtFalecPaiNoite.Value
    objFamilias.dtConjugeDtNascMae = MaskedParaDate(ConjugeDtNascMae)
    objFamilias.iConjugeDtNascMaeNoite = ConjugeDtNascMaeNoite.Value
    objFamilias.dtConjugeDtFalecMae = MaskedParaDate(ConjugeDtFalecMae)
    objFamilias.iConjugeDtFalecMaeNoite = ConjugeDtFalecMaeNoite.Value
    objFamilias.dtConjugeDtFalec = MaskedParaDate(ConjugeDtFalec)
    objFamilias.iConjugeDtFalecNoite = ConjugeDtFalecNoite.Value
    
    objFamilias.dtAtualizadoEm = gdtDataHoje
    
    objFamilias.lCodCliente = StrParaLong(CodCliente.Text)
    objFamilias.dValorContribuicao = StrParaDbl(ValorContribuicao.Text)
    
    lErro = Move_GridFilhos_Memoria(objFamilias)
    If lErro <> SUCESSO Then gError 140916

    lErro = Move_Info_Memoria(objFamilias)
    If lErro <> SUCESSO Then gError 140917

    lErro = Move_Enderecos_Memoria(objFamilias)
    If lErro <> SUCESSO Then gError 140918

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 140916 To 140918

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159907)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objFamilias As New ClassFamilias

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Familias"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objFamilias)
    If lErro <> SUCESSO Then gError 130433

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodFamilia", objFamilias.lCodFamilia, 0, "CodFamilia"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 130433

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159908)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objFamilias As New ClassFamilias

On Error GoTo Erro_Tela_Preenche

    objFamilias.lCodFamilia = colCampoValor.Item("CodFamilia").vValor

    If objFamilias.lCodFamilia <> 0 Then
    
        lErro = Traz_Familias_Tela(objFamilias)
        If lErro <> SUCESSO Then gError 130434
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 130434

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159909)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFamilias As New ClassFamilias

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) = 0 Then gError 130435
    '#####################

    'Preenche o objFamilias
    lErro = Move_Tela_Memoria(objFamilias)
    If lErro <> SUCESSO Then gError 130436

    lErro = Trata_Alteracao(objFamilias, objFamilias.lCodFamilia)
    If lErro <> SUCESSO Then gError 130437

    'Grava o/a Familias no Banco de Dados
    lErro = CF("Familias_Grava", objFamilias)
    If lErro <> SUCESSO Then gError 130438

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130435
            Call Rotina_Erro(vbOKOnly, "ERRO_CODFAMILIA_FAMILIAS_NAO_PREENCHIDO", gErr)
            CodFamilia.SetFocus

        Case 130436, 130437, 130438

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159910)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Familias() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Limpa_Tela_Familias

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Limpa_ListaBox(TitularInfo)
    Call Limpa_ListaBox(ConjugeInfo)
    Call Limpa_ListaBox(FilhosInfo)

    'Escolhe Estado da FilialEmpresa
    lErro = CF("Estado_Seleciona", Estado)
    If lErro <> SUCESSO Then gError 140931

    'Seleciona Brasil nas Combos de Pais se existir
    For iIndice = 0 To 2
        For iIndice2 = 0 To Pais(iIndice).ListCount - 1

            If right(Pais(iIndice).List(iIndice2), 6) = "Brasil" Then
                Pais(iIndice).ListIndex = iIndice2
                Exit For
            End If

        Next
    Next
    
    EstadoCivil.ListIndex = -1
    CohenLeviIsrael.ListIndex = -1
    LocalCobranca.ListIndex = -1
    
    Call Grid_Limpa(objGridFilhos)
     
    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual1 = TAB_TITULAR
    'Torna Frame atual visível
    Frame1(iFrameAtual1).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual1).Selected = True
     
    'Torna Frame atual invisível
    Frame2(TabStrip2.SelectedItem.Index).Visible = False
    iFrameAtual2 = TAB_ENDERECO_TITULAR
    'Torna Frame atual visível
    Frame2(iFrameAtual2).Visible = True
    TabStrip2.Tabs.Item(iFrameAtual2).Selected = True
    
    NomeFilho.Caption = ""
    
    TitSaudacao.Text = ""
    ConjSaudacao.Text = ""
    
    Set gcolcolFilhosInfo = New Collection
     
    iAlterado = 0

    Limpa_Tela_Familias = SUCESSO

    Exit Function

Erro_Limpa_Tela_Familias:

    Limpa_Tela_Familias = gErr

    Select Case gErr
    
        Case 140931

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159911)

    End Select

    Exit Function

End Function

Function Traz_Familias_Tela(objFamilias As ClassFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Familias_Tela

    'Lê o Familias que está sendo Passado
    lErro = CF("Familias_Le", objFamilias)
    If lErro <> SUCESSO And lErro <> 130413 Then gError 130439

    If lErro = SUCESSO Then
        
        TitSaudacao.Text = objFamilias.sTitularSaudacao
        ConjSaudacao.Text = objFamilias.sConjugeSaudacao

        If objFamilias.lCodFamilia <> 0 Then CodFamilia.Text = CStr(objFamilias.lCodFamilia)
        Sobrenome.Text = objFamilias.sSobrenome
        TitularNome.Text = objFamilias.sTitularNome
        TitularNomeHebr.Text = objFamilias.sTitularNomeHebr
        TitularNomeFirma.Text = objFamilias.sTitularNomeFirma
        
        If objFamilias.iLocalCobranca <> 0 Then Call Combo_Seleciona_ItemData(LocalCobranca, objFamilias.iLocalCobranca)
        If objFamilias.iEstadoCivil <> 0 Then Call Combo_Seleciona_ItemData(EstadoCivil, objFamilias.iEstadoCivil)
        
        TitularProfissao.Text = objFamilias.sTitularProfissao

        Call DateParaMasked(TitularDtNasc, objFamilias.dtTitularDtNasc)
        TitularDtNascNoite.Value = objFamilias.iTitularDtNascNoite

        Call DateParaMasked(DataCasamento, objFamilias.dtDataCasamento)
        DataCasamentoNoite.Value = objFamilias.iDataCasamentoNoite
        
        Select Case objFamilias.sCohenLeviIsrael
        
            Case STRING_TRIPO_COHEN
                Call Combo_Seleciona_ItemData(CohenLeviIsrael, TRIPO_COHEN)
        
            Case STRING_TRIPO_ISRAEL
                Call Combo_Seleciona_ItemData(CohenLeviIsrael, TRIPO_ISRAEL)
        
            Case STRING_TRIPO_LEVI
                Call Combo_Seleciona_ItemData(CohenLeviIsrael, TRIPO_LEVI)
        
        End Select
        
        TitularPai.Text = objFamilias.sTitularPai
        TitularPaiHebr.Text = objFamilias.sTitularPaiHebr
        TitularMae.Text = objFamilias.sTitularMae
        TitularMaeHebr.Text = objFamilias.sTitularMaeHebr

        Call DateParaMasked(TitularDtNascPai, objFamilias.dtTitularDtNascPai)
        TitularDtNascPaiNoite.Value = objFamilias.iTitularDtNascPaiNoite
        Call DateParaMasked(TitularDtFalecPai, objFamilias.dtTitularDtFalecPai)
        TitularDtFalecPaiNoite.Value = objFamilias.iTitularDtFalecPaiNoite
        Call DateParaMasked(TitularDtNascMae, objFamilias.dtTitularDtNascMae)
        TitularDtNascMaeNoite.Value = objFamilias.iTitularDtNascMaeNoite
        Call DateParaMasked(TitularDtFalecMae, objFamilias.dtTitularDtFalecMae)
        TitularDtFalecMaeNoite.Value = objFamilias.iTitularDtFalecMaeNoite
        
        ConjugeNome.Text = objFamilias.sConjugeNome
        ConjugeNomeHebr.Text = objFamilias.sConjugeNomeHebr
        Call DateParaMasked(ConjugeDtNasc, objFamilias.dtConjugeDtNasc)
        ConjugeDtNascNoite.Value = objFamilias.iConjugeDtNascNoite
        ConjugeProfissao.Text = objFamilias.sConjugeProfissao
        ConjugeNomeFirma.Text = objFamilias.sConjugeNomeFirma
        
        ConjugePai.Text = objFamilias.sConjugePai
        ConjugePaiHebr.Text = objFamilias.sConjugePaiHebr
        ConjugeMae.Text = objFamilias.sConjugeMae
        ConjugeMaeHebr.Text = objFamilias.sConjugeMaeHebr
        Call DateParaMasked(ConjugeDtNascPai, objFamilias.dtConjugeDtNascPai)
        ConjugeDtNascPaiNoite.Value = objFamilias.iConjugeDtNascPaiNoite
        Call DateParaMasked(ConjugeDtFalecPai, objFamilias.dtConjugeDtFalecPai)
        ConjugeDtFalecPaiNoite.Value = objFamilias.iConjugeDtFalecPaiNoite
        Call DateParaMasked(ConjugeDtNascMae, objFamilias.dtConjugeDtNascMae)
        ConjugeDtNascMaeNoite.Value = objFamilias.iConjugeDtNascMaeNoite
        Call DateParaMasked(ConjugeDtFalecMae, objFamilias.dtConjugeDtFalecMae)
        ConjugeDtFalecMaeNoite.Value = objFamilias.iConjugeDtFalecMaeNoite
        Call DateParaMasked(ConjugeDtFalec, objFamilias.dtConjugeDtFalec)
        ConjugeDtFalecNoite.Value = objFamilias.iConjugeDtFalecNoite
        If objFamilias.dtAtualizadoEm <> DATA_NULA Then
            AtualizadoEm.Caption = Format(objFamilias.dtAtualizadoEm, "dd/mm/yyyy")
        Else
            AtualizadoEm.Caption = ""
        End If

        If objFamilias.lCodCliente <> 0 Then
            CodCliente.Text = CStr(objFamilias.lCodCliente)
        Else
            CodCliente.Text = ""
        End If
        
        If objFamilias.dValorContribuicao <> 0 Then
            ValorContribuicao.Text = Format(objFamilias.dValorContribuicao, ValorContribuicao.Format)
        Else
            ValorContribuicao.Text = ""
        End If
        
        lErro = Traz_Filhos_Tela(objFamilias)
        If lErro <> SUCESSO Then gError 140920
        
        lErro = Traz_Info_Tela(objFamilias)
        If lErro <> SUCESSO Then gError 140921
        
        lErro = Traz_Enderecos_Tela(objFamilias)
        If lErro <> SUCESSO Then gError 140922
        
    End If
    
    iAlterado = 0

    Traz_Familias_Tela = SUCESSO

    Exit Function

Erro_Traz_Familias_Tela:

    Traz_Familias_Tela = gErr

    Select Case gErr

        Case 130439, 140920 To 140922

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159912)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 130440

    'Limpa Tela
    Call Limpa_Tela_Familias

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 130440

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159913)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159914)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 130441

    Call Limpa_Tela_Familias

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 130441

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159915)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFamilias As New ClassFamilias
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) = 0 Then gError 130442
    '#####################

    objFamilias.lCodFamilia = StrParaLong(CodFamilia.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_FAMILIAS", objFamilias.lCodFamilia)

    If vbMsgRes = vbYes Then

        'Exclui a Família de consumo
        lErro = CF("Familias_Exclui", objFamilias)
        If lErro <> SUCESSO Then gError 130443

        'Limpa Tela
        Call Limpa_Tela_Familias

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130442
            Call Rotina_Erro(vbOKOnly, "ERRO_CODFAMILIA_FAMILIAS_NAO_PREENCHIDO", gErr)
            CodFamilia.SetFocus

        Case 130443

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159916)

    End Select

    Exit Sub

End Sub

Private Sub CodFamilia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodFamilia_Validate

    'Verifica se CodFamilia está preenchida
    If Len(Trim(CodFamilia.Text)) <> 0 Then

       'Critica a CodFamilia
       lErro = Long_Critica(CodFamilia.Text)
       If lErro <> SUCESSO Then gError 130444

    End If

    Exit Sub

Erro_CodFamilia_Validate:

    Cancel = True

    Select Case gErr

        Case 130444

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159917)

    End Select

    Exit Sub

End Sub

Private Sub CodFamilia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodFamilia, iAlterado)
    
End Sub

Private Sub CodFamilia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sobrenome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularNome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularNomeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularNomeFirma_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularProfissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtNasc_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNasc_DownClick

    TitularDtNasc.SetFocus

    If Len(TitularDtNasc.ClipText) > 0 Then

        sData = TitularDtNasc.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130449

        TitularDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNasc_DownClick:

    Select Case gErr

        Case 130449

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159918)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtNasc_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNasc_UpClick

    TitularDtNasc.SetFocus

    If Len(Trim(TitularDtNasc.ClipText)) > 0 Then

        sData = TitularDtNasc.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130450

        TitularDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNasc_UpClick:

    Select Case gErr

        Case 130450

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159919)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNasc_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNasc, iAlterado)
    
End Sub

Private Sub TitularDtNasc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNasc_Validate

    If Len(Trim(TitularDtNasc.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtNasc.Text)
        If lErro <> SUCESSO Then gError 130451

    End If

    Exit Sub

Erro_TitularDtNasc_Validate:

    Cancel = True

    Select Case gErr

        Case 130451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159920)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNasc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtNascNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataCasamento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCasamento_DownClick

    DataCasamento.SetFocus

    If Len(DataCasamento.ClipText) > 0 Then

        sData = DataCasamento.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130453

        DataCasamento.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCasamento_DownClick:

    Select Case gErr

        Case 130453

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159921)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCasamento_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCasamento_UpClick

    DataCasamento.SetFocus

    If Len(Trim(DataCasamento.ClipText)) > 0 Then

        sData = DataCasamento.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130454

        DataCasamento.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCasamento_UpClick:

    Select Case gErr

        Case 130454

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159922)

    End Select

    Exit Sub

End Sub

Private Sub DataCasamento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCasamento, iAlterado)
    
End Sub

Private Sub DataCasamento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCasamento_Validate

    If Len(Trim(DataCasamento.ClipText)) <> 0 Then

        lErro = Data_Critica(DataCasamento.Text)
        If lErro <> SUCESSO Then gError 130455

    End If

    Exit Sub

Erro_DataCasamento_Validate:

    Cancel = True

    Select Case gErr

        Case 130455

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159923)

    End Select

    Exit Sub

End Sub

Private Sub DataCasamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCasamentoNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularPaiHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularMaeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtNascPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascPai_DownClick

    TitularDtNascPai.SetFocus

    If Len(TitularDtNascPai.ClipText) > 0 Then

        sData = TitularDtNascPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130457

        TitularDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascPai_DownClick:

    Select Case gErr

        Case 130457

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159924)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtNascPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascPai_UpClick

    TitularDtNascPai.SetFocus

    If Len(Trim(TitularDtNascPai.ClipText)) > 0 Then

        sData = TitularDtNascPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130458

        TitularDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascPai_UpClick:

    Select Case gErr

        Case 130458

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159925)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNascPai, iAlterado)
    
End Sub

Private Sub TitularDtNascPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNascPai_Validate

    If Len(Trim(TitularDtNascPai.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtNascPai.Text)
        If lErro <> SUCESSO Then gError 130459

    End If

    Exit Sub

Erro_TitularDtNascPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130459

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159926)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtNascPaiNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtFalecPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecPai_DownClick

    TitularDtFalecPai.SetFocus

    If Len(TitularDtFalecPai.ClipText) > 0 Then

        sData = TitularDtFalecPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130461

        TitularDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecPai_DownClick:

    Select Case gErr

        Case 130461

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159927)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtFalecPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecPai_UpClick

    TitularDtFalecPai.SetFocus

    If Len(Trim(TitularDtFalecPai.ClipText)) > 0 Then

        sData = TitularDtFalecPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130462

        TitularDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecPai_UpClick:

    Select Case gErr

        Case 130462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159928)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtFalecPai, iAlterado)
    
End Sub

Private Sub TitularDtFalecPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtFalecPai_Validate

    If Len(Trim(TitularDtFalecPai.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtFalecPai.Text)
        If lErro <> SUCESSO Then gError 130463

    End If

    Exit Sub

Erro_TitularDtFalecPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130463

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159929)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtFalecPaiNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtNascMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascMae_DownClick

    TitularDtNascMae.SetFocus

    If Len(TitularDtNascMae.ClipText) > 0 Then

        sData = TitularDtNascMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130465

        TitularDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascMae_DownClick:

    Select Case gErr

        Case 130465

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159930)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtNascMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascMae_UpClick

    TitularDtNascMae.SetFocus

    If Len(Trim(TitularDtNascMae.ClipText)) > 0 Then

        sData = TitularDtNascMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130466

        TitularDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascMae_UpClick:

    Select Case gErr

        Case 130466

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159931)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNascMae, iAlterado)
    
End Sub

Private Sub TitularDtNascMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNascMae_Validate

    If Len(Trim(TitularDtNascMae.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtNascMae.Text)
        If lErro <> SUCESSO Then gError 130467

    End If

    Exit Sub

Erro_TitularDtNascMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130467

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159932)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub TitularDtNascMaeNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtFalecMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecMae_DownClick

    TitularDtFalecMae.SetFocus

    If Len(TitularDtFalecMae.ClipText) > 0 Then

        sData = TitularDtFalecMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130469

        TitularDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecMae_DownClick:

    Select Case gErr

        Case 130469

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159933)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtFalecMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecMae_UpClick

    TitularDtFalecMae.SetFocus

    If Len(Trim(TitularDtFalecMae.ClipText)) > 0 Then

        sData = TitularDtFalecMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130470

        TitularDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecMae_UpClick:

    Select Case gErr

        Case 130470

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159934)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtFalecMae, iAlterado)
    
End Sub

Private Sub TitularDtFalecMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtFalecMae_Validate

    If Len(Trim(TitularDtFalecMae.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtFalecMae.Text)
        If lErro <> SUCESSO Then gError 130471

    End If

    Exit Sub

Erro_TitularDtFalecMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130471

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159935)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtFalecMaeNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeNome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeNomeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtNasc_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNasc_DownClick

    ConjugeDtNasc.SetFocus

    If Len(ConjugeDtNasc.ClipText) > 0 Then

        sData = ConjugeDtNasc.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130473

        ConjugeDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNasc_DownClick:

    Select Case gErr

        Case 130473

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159936)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtNasc_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNasc_UpClick

    ConjugeDtNasc.SetFocus

    If Len(Trim(ConjugeDtNasc.ClipText)) > 0 Then

        sData = ConjugeDtNasc.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130474

        ConjugeDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNasc_UpClick:

    Select Case gErr

        Case 130474

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159937)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNasc_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNasc, iAlterado)
    
End Sub

Private Sub ConjugeDtNasc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNasc_Validate

    If Len(Trim(ConjugeDtNasc.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtNasc.Text)
        If lErro <> SUCESSO Then gError 130475

    End If

    Exit Sub

Erro_ConjugeDtNasc_Validate:

    Cancel = True

    Select Case gErr

        Case 130475

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159938)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNasc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtNascNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeProfissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeNomeFirma_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugePai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugePaiHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeMaeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtNascPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascPai_DownClick

    ConjugeDtNascPai.SetFocus

    If Len(ConjugeDtNascPai.ClipText) > 0 Then

        sData = ConjugeDtNascPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130478

        ConjugeDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascPai_DownClick:

    Select Case gErr

        Case 130478

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159939)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtNascPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascPai_UpClick

    ConjugeDtNascPai.SetFocus

    If Len(Trim(ConjugeDtNascPai.ClipText)) > 0 Then

        sData = ConjugeDtNascPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130479

        ConjugeDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascPai_UpClick:

    Select Case gErr

        Case 130479

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159940)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNascPai, iAlterado)
    
End Sub

Private Sub ConjugeDtNascPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNascPai_Validate

    If Len(Trim(ConjugeDtNascPai.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtNascPai.Text)
        If lErro <> SUCESSO Then gError 130480

    End If

    Exit Sub

Erro_ConjugeDtNascPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159941)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtNascPaiNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtFalecPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecPai_DownClick

    ConjugeDtFalecPai.SetFocus

    If Len(ConjugeDtFalecPai.ClipText) > 0 Then

        sData = ConjugeDtFalecPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130482

        ConjugeDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecPai_DownClick:

    Select Case gErr

        Case 130482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159942)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtFalecPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecPai_UpClick

    ConjugeDtFalecPai.SetFocus

    If Len(Trim(ConjugeDtFalecPai.ClipText)) > 0 Then

        sData = ConjugeDtFalecPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130483

        ConjugeDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecPai_UpClick:

    Select Case gErr

        Case 130483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159943)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalecPai, iAlterado)
    
End Sub

Private Sub ConjugeDtFalecPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalecPai_Validate

    If Len(Trim(ConjugeDtFalecPai.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtFalecPai.Text)
        If lErro <> SUCESSO Then gError 130484

    End If

    Exit Sub

Erro_ConjugeDtFalecPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159944)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtFalecPaiNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtNascMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascMae_DownClick

    ConjugeDtNascMae.SetFocus

    If Len(ConjugeDtNascMae.ClipText) > 0 Then

        sData = ConjugeDtNascMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130486

        ConjugeDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascMae_DownClick:

    Select Case gErr

        Case 130486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159945)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtNascMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascMae_UpClick

    ConjugeDtNascMae.SetFocus

    If Len(Trim(ConjugeDtNascMae.ClipText)) > 0 Then

        sData = ConjugeDtNascMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130487

        ConjugeDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascMae_UpClick:

    Select Case gErr

        Case 130487

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159946)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNascMae, iAlterado)
    
End Sub

Private Sub ConjugeDtNascMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNascMae_Validate

    If Len(Trim(ConjugeDtNascMae.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtNascMae.Text)
        If lErro <> SUCESSO Then gError 130488

    End If

    Exit Sub

Erro_ConjugeDtNascMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130488

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159947)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtNascMaeNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtFalecMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecMae_DownClick

    ConjugeDtFalecMae.SetFocus

    If Len(ConjugeDtFalecMae.ClipText) > 0 Then

        sData = ConjugeDtFalecMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130490

        ConjugeDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecMae_DownClick:

    Select Case gErr

        Case 130490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159948)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtFalecMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecMae_UpClick

    ConjugeDtFalecMae.SetFocus

    If Len(Trim(ConjugeDtFalecMae.ClipText)) > 0 Then

        sData = ConjugeDtFalecMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130491

        ConjugeDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecMae_UpClick:

    Select Case gErr

        Case 130491

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159949)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalecMae, iAlterado)
    
End Sub

Private Sub ConjugeDtFalecMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalecMae_Validate

    If Len(Trim(ConjugeDtFalecMae.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtFalecMae.Text)
        If lErro <> SUCESSO Then gError 130492

    End If

    Exit Sub

Erro_ConjugeDtFalecMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159950)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtFalecMaeNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtFalec_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalec_DownClick

    ConjugeDtFalec.SetFocus

    If Len(ConjugeDtFalec.ClipText) > 0 Then

        sData = ConjugeDtFalec.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130494

        ConjugeDtFalec.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalec_DownClick:

    Select Case gErr

        Case 130494

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159951)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtFalec_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalec_UpClick

    ConjugeDtFalec.SetFocus

    If Len(Trim(ConjugeDtFalec.ClipText)) > 0 Then

        sData = ConjugeDtFalec.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130495

        ConjugeDtFalec.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalec_UpClick:

    Select Case gErr

        Case 130495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159952)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalec_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalec, iAlterado)
    
End Sub

Private Sub ConjugeDtFalec_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalec_Validate

    If Len(Trim(ConjugeDtFalec.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtFalec.Text)
        If lErro <> SUCESSO Then gError 130496

    End If

    Exit Sub

Erro_ConjugeDtFalec_Validate:

    Cancel = True

    Select Case gErr

        Case 130496

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159953)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalec_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtFalecNoite_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodCliente_Validate

    'Verifica se CodCliente está preenchida
    If Len(Trim(CodCliente.Text)) <> 0 Then

       'Critica a CodCliente
       lErro = Long_Critica(CodCliente.Text)
       If lErro <> SUCESSO Then gError 130501

    End If

    Exit Sub

Erro_CodCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 130501

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159954)

    End Select

    Exit Sub

End Sub

Private Sub CodCliente_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodCliente, iAlterado)
    
End Sub

Private Sub CodCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorContribuicao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorContribuicao_Validate

    'Verifica se ValorContribuicao está preenchida
    If Len(Trim(ValorContribuicao.Text)) <> 0 Then

       'Critica a ValorContribuicao
       lErro = Valor_Positivo_Critica(ValorContribuicao.Text)
       If lErro <> SUCESSO Then gError 130502

    End If

    Exit Sub

Erro_ValorContribuicao_Validate:

    Cancel = True

    Select Case gErr

        Case 130502

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159955)

    End Select

    Exit Sub

End Sub

Private Sub ValorContribuicao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorContribuicao, iAlterado)
    
End Sub

Private Sub ValorContribuicao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodFamilia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFamilias As ClassFamilias

On Error GoTo Erro_objEventoCodFamilia_evSelecao

    Set objFamilias = obj1

    'Mostra os dados do Familias na tela
    lErro = Traz_Familias_Tela(objFamilias)
    If lErro <> SUCESSO Then gError 130503

    Me.Show

    Exit Sub

Erro_objEventoCodFamilia_evSelecao:

    Select Case gErr

        Case 130503


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159956)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodFamilia_Click()

Dim lErro As Long
Dim objFamilias As New ClassFamilias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodFamilia_Click

    'Verifica se o CodFamilia foi preenchido
    If Len(Trim(CodFamilia.Text)) <> 0 Then

        objFamilias.lCodFamilia = CodFamilia.Text

    End If

    Call Chama_Tela("FamiliasLista", colSelecao, objFamilias, objEventoCodFamilia)

    Exit Sub

Erro_LabelCodFamilia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159957)

    End Select

    Exit Sub

End Sub

'###########################################################
'Inserido por Wagner 21/11/2005
Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual1 Then

        If TabStrip_PodeTrocarTab(iFrameAtual1, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual1).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual1 = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Sub TabStrip2_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index <> iFrameAtual2 Then

        If TabStrip_PodeTrocarTab(iFrameAtual2, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame2(TabStrip2.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame2(iFrameAtual2).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual2 = TabStrip2.SelectedItem.Index
        
    End If

End Sub

Private Function Inicializa_GridFilhos(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Nome")
    objGrid.colColuna.Add ("Nome Hebraico")
    objGrid.colColuna.Add ("Dt Nascimento")
    objGrid.colColuna.Add ("Noite")
    objGrid.colColuna.Add ("Tel")
    objGrid.colColuna.Add ("Email")
    objGrid.colColuna.Add ("Dt Falecimento")
    objGrid.colColuna.Add ("Noite")

    'Controles que participam do Grid
    objGrid.colCampo.Add (FilhoNome.Name)
    objGrid.colCampo.Add (FilhoNomeHebr.Name)
    objGrid.colCampo.Add (FilhoDataNasc.Name)
    objGrid.colCampo.Add (FilhoDataNascNoite.Name)
    objGrid.colCampo.Add (FilhoTel.Name)
    objGrid.colCampo.Add (FilhoEmail.Name)
    objGrid.colCampo.Add (FilhoDataFal.Name)
    objGrid.colCampo.Add (FilhoDataFalNoite.Name)

    'Colunas do Grid
    iGrid_FilhoNome_Col = 1
    iGrid_FilhoNomeHebr_Col = 2
    iGrid_FilhoDataNasc_Col = 3
    iGrid_FilhoDataNascNoite_Col = 4
    iGrid_FilhoTelefone_Col = 5
    iGrid_FilhoEmail_Col = 6
    iGrid_FilhoDataFal_Col = 7
    iGrid_FilhoDataFalNoite_Col = 8

    objGrid.objGrid = GridFilhos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iLinhasVisiveis = 11

    'Largura da primeira coluna
    GridFilhos.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridFilhos = SUCESSO

End Function

Private Sub GridFilhos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFilhos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFilhos, iAlterado)
    End If

End Sub

Private Sub GridFilhos_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFilhos)

End Sub

Private Sub GridFilhos_EnterCell()

    Call Grid_Entrada_Celula(objGridFilhos, iAlterado)

End Sub

Private Sub GridFilhos_LeaveCell()
    
    Call Saida_Celula(objGridFilhos)

End Sub

Private Sub GridFilhos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFilhos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFilhos, iAlterado)
    End If

End Sub

Private Sub GridFilhos_RowColChange()

    Call Grid_RowColChange(objGridFilhos)
    
    If GridFilhos.Row <> 0 Then
        If gcolcolFilhosInfo.Count >= GridFilhos.Row Then
            Call Traz_Info_Tela2(gcolcolFilhosInfo.Item(GridFilhos.Row), 0, FilhosInfo)
            NomeFilho.Caption = GridFilhos.TextMatrix(GridFilhos.Row, iGrid_FilhoNome_Col)
        Else
            Call Limpa_ListaBox(FilhosInfo)
            NomeFilho.Caption = ""
        End If
    End If

End Sub

Private Sub GridFilhos_Scroll()

    Call Grid_Scroll(objGridFilhos)

End Sub

Private Sub GridFilhos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhaAnterior As Integer
Dim iLinhasExistentesAnterior As Integer

On Error GoTo Erro_GridFilhos_KeyDown

    'guarda as linhas do grid antes de apagar
    iLinhaAnterior = GridFilhos.Row
    iLinhasExistentesAnterior = objGridFilhos.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridFilhos)
        
    If objGridFilhos.iLinhasExistentes < iLinhasExistentesAnterior Then
    
        'apaga a Info
        gcolcolFilhosInfo.Remove iLinhaAnterior
            
    End If
    
    Exit Sub
    
Erro_GridFilhos_KeyDown:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159958)
    
    End Select

    Exit Sub
    
End Sub

Private Sub GridFilhos_LostFocus()

    Call Grid_Libera_Foco(objGridFilhos)

End Sub

Private Sub FilhoNome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoNome_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoNome_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoNome_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoNome
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilhoNomeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoNomeHebr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoNomeHebr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoNomeHebr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoNomeHebr
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilhoDataNasc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoDataNasc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoDataNasc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoDataNasc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoDataNasc
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilhoDataNascNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoDataNascNoite_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoDataNascNoite_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoDataNascNoite_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoDataNascNoite
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilhoDataFal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoDataFal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoDataFal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoDataFal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoDataFal
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilhoDataFalNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoDataFalNoite_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoDataFalNoite_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoDataFalNoite_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoDataFalNoite
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridFilhos.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
            
                Case iGrid_FilhoDataNasc_Col
                
                    lErro = Saida_Celula_FilhoDataNasc(objGridInt)
                    If lErro <> SUCESSO Then gError 140914
                    
                Case iGrid_FilhoDataFal_Col
                
                    lErro = Saida_Celula_FilhoDataFal(objGridInt)
                    If lErro <> SUCESSO Then gError 140966
                
                Case iGrid_FilhoNome_Col
                
                    lErro = Saida_Celula_FilhoNome(objGridInt)
                    If lErro <> SUCESSO Then gError 140915
                
                Case iGrid_FilhoNomeHebr_Col
                
                    lErro = Saida_Celula_FilhoNomeHebr(objGridInt)
                    If lErro <> SUCESSO Then gError 140916
                
                Case iGrid_FilhoTelefone_Col
                
                    lErro = Saida_Celula_FilhoTel(objGridInt)
                    If lErro <> SUCESSO Then gError 140916
                
                Case iGrid_FilhoEmail_Col
                
                    lErro = Saida_Celula_FilhoEmail(objGridInt)
                    If lErro <> SUCESSO Then gError 140916
                
                Case Else
        
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 140873

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 140873
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 140914 To 140916

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159959)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilhoNome(objGridInt As AdmGrid) As Long
'faz a critica da celula de nome do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_FilhoNome

    Set objGridInt.objControle = FilhoNome

    If Len(Trim(FilhoNome.Text)) > 0 Then
        If GridFilhos.Row - GridFilhos.FixedRows = objGridFilhos.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Call Adiciona_colFilhosInfo
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140910

    Saida_Celula_FilhoNome = SUCESSO

    Exit Function

Erro_Saida_Celula_FilhoNome:

    Saida_Celula_FilhoNome = gErr

    Select Case gErr

        Case 140910
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159960)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilhoNomeHebr(objGridInt As AdmGrid) As Long
'faz a critica da celula de nome do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_FilhoNomeHebr

    Set objGridInt.objControle = FilhoNomeHebr

    If Len(Trim(FilhoNomeHebr.Text)) > 0 Then
        If GridFilhos.Row - GridFilhos.FixedRows = objGridFilhos.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Call Adiciona_colFilhosInfo
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140911

    Saida_Celula_FilhoNomeHebr = SUCESSO

    Exit Function

Erro_Saida_Celula_FilhoNomeHebr:

    Saida_Celula_FilhoNomeHebr = gErr

    Select Case gErr

        Case 140911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159961)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilhoDataNasc(objGridInt As AdmGrid) As Long
'faz a critica da celula de nome do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_FilhoDataNasc

    Set objGridInt.objControle = FilhoDataNasc

    'verifica se a data está preenchida
    If Len(Trim(FilhoDataNasc.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(FilhoDataNasc.Text)
        If lErro <> SUCESSO Then gError 140912
    
        If GridFilhos.Row - GridFilhos.FixedRows = objGridFilhos.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Call Adiciona_colFilhosInfo
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140913

    Saida_Celula_FilhoDataNasc = SUCESSO

    Exit Function

Erro_Saida_Celula_FilhoDataNasc:

    Saida_Celula_FilhoDataNasc = gErr

    Select Case gErr

        Case 140912, 140913
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159962)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilhoDataFal(objGridInt As AdmGrid) As Long
'faz a critica da celula de nome do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_FilhoDataFal

    Set objGridInt.objControle = FilhoDataFal

    'verifica se a data está preenchida
    If Len(Trim(FilhoDataFal.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(FilhoDataFal.Text)
        If lErro <> SUCESSO Then gError 140964
    
        If GridFilhos.Row - GridFilhos.FixedRows = objGridFilhos.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Call Adiciona_colFilhosInfo
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140965

    Saida_Celula_FilhoDataFal = SUCESSO

    Exit Function

Erro_Saida_Celula_FilhoDataFal:

    Saida_Celula_FilhoDataFal = gErr

    Select Case gErr

        Case 140964, 140965
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159963)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Traz_Filhos_Tela(objFamilias As ClassFamilias) As Long
'Carrega o Grid de Filhos

Dim lErro As Long
Dim iIndice As Integer
Dim iIndex As Integer
Dim bAchou As Boolean
Dim objFilhosFamilias As ClassFilhosFamilias
Dim colFilhosInfo As Collection
Dim objFilhosInfo As ClassFamiliasInfo

On Error GoTo Erro_Traz_Filhos_Tela

    Call Grid_Limpa(objGridFilhos)
     
    Set gcolcolFilhosInfo = New Collection

    For Each objFilhosFamilias In objFamilias.colFilhos
    
        iIndice = iIndice + 1
    
        If objFilhosFamilias.dtDtNasc <> DATA_NULA Then GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataNasc_Col) = Format(objFilhosFamilias.dtDtNasc, "dd/mm/yyyy")
        GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataNascNoite_Col) = objFilhosFamilias.iDtNascNoite
        If objFilhosFamilias.dtDtFal <> DATA_NULA Then GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataFal_Col) = Format(objFilhosFamilias.dtDtFal, "dd/mm/yyyy")
        GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataFalNoite_Col) = objFilhosFamilias.iDtFalNoite
        GridFilhos.TextMatrix(iIndice, iGrid_FilhoNome_Col) = objFilhosFamilias.sNome
        GridFilhos.TextMatrix(iIndice, iGrid_FilhoNomeHebr_Col) = objFilhosFamilias.sNomeHebr
        GridFilhos.TextMatrix(iIndice, iGrid_FilhoTelefone_Col) = objFilhosFamilias.sTelefone
        GridFilhos.TextMatrix(iIndice, iGrid_FilhoEmail_Col) = objFilhosFamilias.sEmail
        
        Set colFilhosInfo = New Collection
        
        For Each objFilhosInfo In objFamilias.colFamiliaInfo
        
            If objFilhosInfo.iSeq = objFilhosFamilias.iSeqFilho Then
            
                objFilhosInfo.iSeq = 0
            
                colFilhosInfo.Add objFilhosInfo
                
            End If
        
        Next
        
        For iIndex = 0 To FilhosInfo.ListCount - 1
        
            bAchou = False
        
            For Each objFilhosInfo In colFilhosInfo
                
                If FilhosInfo.ItemData(iIndex) = objFilhosInfo.iCodInfo Then
                    bAchou = True
                    Exit For
                End If
                            
            Next
        
            'Se o item é novo adiciona no filho como não marcado
            If Not bAchou Then
             
                Set objFilhosInfo = New ClassFamiliasInfo
                
                objFilhosInfo.iCodInfo = FilhosInfo.ItemData(iIndex)
                objFilhosInfo.iValor = DESMARCADO
                
                colFilhosInfo.Add objFilhosInfo
             
            End If
        
        Next
        
        gcolcolFilhosInfo.Add colFilhosInfo
    
    Next
    
    objGridFilhos.iLinhasExistentes = objFamilias.colFilhos.Count
    
    Call Grid_Refresh_Checkbox(objGridFilhos)

    Traz_Filhos_Tela = SUCESSO

    Exit Function

Erro_Traz_Filhos_Tela:

    Traz_Filhos_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159964)

    End Select

    Exit Function

End Function

Function Traz_Info_Tela(objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFamiliaInfo As ClassFamiliasInfo

On Error GoTo Erro_Traz_Info_Tela

    'Traz as Infomações do Titular para Tela
    lErro = Traz_Info_Tela2(objFamilias.colFamiliaInfo, FAMILIAINFO_TITULAR, TitularInfo)
    If lErro <> SUCESSO Then gError 140957
    
    'Traz as Infomações da Conjuge para Tela
    lErro = Traz_Info_Tela2(objFamilias.colFamiliaInfo, FAMILIAINFO_CONJUGE, ConjugeInfo)
    If lErro <> SUCESSO Then gError 140958
    
'    'Traz as Infomações do primeiro filho para Tela
'    lErro = Traz_Info_Tela2(objFamilias.colFamiliaInfo, 1, FilhosInfo)
'    If lErro <> SUCESSO Then gError 140959
'
'    NomeFilho.Caption = GridFilhos.TextMatrix(1, iGrid_FilhoNome_Col)
    
    Traz_Info_Tela = SUCESSO

    Exit Function

Erro_Traz_Info_Tela:

    Traz_Info_Tela = gErr

    Select Case gErr
    
        Case 140957 To 140959

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159965)

    End Select

    Exit Function

End Function

Function Traz_Enderecos_Tela(objFamilias As ClassFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Enderecos_Tela

    'Traz o Endereço Residencial do Titular
    lErro = Traz_Enderecos_Tela2(objFamilias.objEnderecoRes, 0)
    If lErro <> SUCESSO Then gError 140963

    'Traz o Endereço Comercial do Titular
    lErro = Traz_Enderecos_Tela2(objFamilias.objEnderecoCom, 1)
    If lErro <> SUCESSO Then gError 140964

    'Traz o Endereço Comercial da Conjuge
    lErro = Traz_Enderecos_Tela2(objFamilias.objEnderecoComConj, 2)
    If lErro <> SUCESSO Then gError 140965
        
    Traz_Enderecos_Tela = SUCESSO

    Exit Function

Erro_Traz_Enderecos_Tela:

    Traz_Enderecos_Tela = gErr

    Select Case gErr
    
        Case 140963 To 140965

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159966)

    End Select

    Exit Function

End Function

Function Move_GridFilhos_Memoria(ByVal objFamilias As ClassFamilias) As Long
'move itens do Grid

Dim lErro As Long
Dim iIndice As Integer
Dim objFilhosFamilia As ClassFilhosFamilias

On Error GoTo Erro_Move_GridFilhos_Memoria

    For iIndice = 1 To objGridFilhos.iLinhasExistentes

        Set objFilhosFamilia = New ClassFilhosFamilias

        objFilhosFamilia.dtDtNasc = StrParaDate(GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataNasc_Col))
        objFilhosFamilia.iDtNascNoite = StrParaInt(GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataNascNoite_Col))
        objFilhosFamilia.dtDtFal = StrParaDate(GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataFal_Col))
        objFilhosFamilia.iDtFalNoite = StrParaInt(GridFilhos.TextMatrix(iIndice, iGrid_FilhoDataFalNoite_Col))
        objFilhosFamilia.sNome = GridFilhos.TextMatrix(iIndice, iGrid_FilhoNome_Col)
        objFilhosFamilia.sNomeHebr = GridFilhos.TextMatrix(iIndice, iGrid_FilhoNomeHebr_Col)
        objFilhosFamilia.sTelefone = GridFilhos.TextMatrix(iIndice, iGrid_FilhoTelefone_Col)
        objFilhosFamilia.sEmail = GridFilhos.TextMatrix(iIndice, iGrid_FilhoEmail_Col)
        objFilhosFamilia.iSeqFilho = iIndice
        
        objFamilias.colFilhos.Add objFilhosFamilia

    Next

    Move_GridFilhos_Memoria = SUCESSO

    Exit Function

Erro_Move_GridFilhos_Memoria:

    Move_GridFilhos_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159967)

    End Select

    Exit Function

End Function

Function Move_Info_Memoria(ByVal objFamilias As ClassFamilias) As Long
'move itens dos ListBox

Dim lErro As Long
Dim iIndice As Integer
Dim colFilhosInfo As Collection
Dim objFamiliasInfo As ClassFamiliasInfo
Dim objFilhosFamilia As ClassFilhosFamilias

On Error GoTo Erro_Move_Info_Memoria

    '--INFO TITULAR
    For iIndice = 0 To TitularInfo.ListCount - 1
    
        Set objFamiliasInfo = New ClassFamiliasInfo
    
        objFamiliasInfo.iSeq = FAMILIAINFO_TITULAR
        objFamiliasInfo.iCodInfo = TitularInfo.ItemData(iIndice)
    
        If TitularInfo.Selected(iIndice) = True Then
            objFamiliasInfo.iValor = MARCADO
        Else
            objFamiliasInfo.iValor = DESMARCADO
        End If
    
        objFamilias.colFamiliaInfo.Add objFamiliasInfo
    
        Select Case TitularInfo.List(iIndice)
        
            Case "AN"
                objFamilias.iAN = objFamiliasInfo.iValor
        
            Case "CD"
                objFamilias.iCD = objFamiliasInfo.iValor
        
            Case "CH"
                objFamilias.iCH = objFamiliasInfo.iValor
        
            Case "CJ"
                objFamilias.iCJ = objFamiliasInfo.iValor
        
            Case "H"
                objFamilias.iH = objFamiliasInfo.iValor
        
            Case "H1"
                objFamilias.iH1 = objFamiliasInfo.iValor
        
            Case "H2"
                objFamilias.iH2 = objFamiliasInfo.iValor
        
            Case "LE"
                objFamilias.iLE = objFamiliasInfo.iValor
        
            Case "LR"
                objFamilias.iLR = objFamiliasInfo.iValor
        
            Case "PA"
                objFamilias.iPA = objFamiliasInfo.iValor
        
            Case "RE"
                objFamilias.iRE = objFamiliasInfo.iValor
        
            Case "SH"
                objFamilias.iSH = objFamiliasInfo.iValor
        
            Case "SI"
                objFamilias.iSI = objFamiliasInfo.iValor
        
            Case "TH"
                objFamilias.iTH = objFamiliasInfo.iValor
        
            Case "VF"
                objFamilias.iVF = objFamiliasInfo.iValor
        
        End Select
        
    Next

    '--INFO CONJUGE
    For iIndice = 0 To ConjugeInfo.ListCount - 1
    
        Set objFamiliasInfo = New ClassFamiliasInfo
    
        objFamiliasInfo.iSeq = FAMILIAINFO_CONJUGE
        objFamiliasInfo.iCodInfo = ConjugeInfo.ItemData(iIndice)
    
        If ConjugeInfo.Selected(iIndice) = True Then
            objFamiliasInfo.iValor = MARCADO
        Else
            objFamiliasInfo.iValor = DESMARCADO
        End If
    
        objFamilias.colFamiliaInfo.Add objFamiliasInfo
    
        Select Case ConjugeInfo.List(iIndice)
        
            Case "AN"
                objFamilias.iANConj = objFamiliasInfo.iValor
        
            Case "CD"
                objFamilias.iCDConj = objFamiliasInfo.iValor
        
            Case "CH"
                objFamilias.iCHConj = objFamiliasInfo.iValor
        
            Case "CJ"
                objFamilias.iCJConj = objFamiliasInfo.iValor
        
            Case "H"
                objFamilias.iHConj = objFamiliasInfo.iValor
        
            Case "H1"
                objFamilias.iH1Conj = objFamiliasInfo.iValor
        
            Case "H2"
                objFamilias.iH2Conj = objFamiliasInfo.iValor
        
            Case "LE"
                objFamilias.iLEConj = objFamiliasInfo.iValor
        
            Case "LR"
                objFamilias.iLRConj = objFamiliasInfo.iValor
        
            Case "PA"
                objFamilias.iPAConj = objFamiliasInfo.iValor
        
            Case "RE"
                objFamilias.iREConj = objFamiliasInfo.iValor
        
            Case "SH"
                objFamilias.iSHConj = objFamiliasInfo.iValor
        
            Case "SI"
                objFamilias.iSIConj = objFamiliasInfo.iValor
        
            Case "TH"
                objFamilias.iTHConj = objFamiliasInfo.iValor
        
            Case "VF"
                objFamilias.iVFConj = objFamiliasInfo.iValor
        
        End Select
        
    Next
    
    iIndice = 0
    '--INFO FILHOS
    'Já está na memória
    For Each colFilhosInfo In gcolcolFilhosInfo
    
        iIndice = iIndice + 1
        
        Set objFilhosFamilia = objFamilias.colFilhos(iIndice)

        For Each objFamiliasInfo In colFilhosInfo
    
            objFamiliasInfo.iSeq = iIndice
    
            objFamilias.colFamiliaInfo.Add objFamiliasInfo
    
            Select Case objFamiliasInfo.iCodInfo
            
                Case 7
                    objFilhosFamilia.iAN = objFamiliasInfo.iValor
                    
                Case 16
                    objFilhosFamilia.iCD = objFamiliasInfo.iValor
                    
                Case 18
                    objFilhosFamilia.iCH = objFamiliasInfo.iValor
                    
                Case 9
                    objFilhosFamilia.iCJ = objFamiliasInfo.iValor
                    
                Case 19
                    objFilhosFamilia.iH = objFamiliasInfo.iValor
                    
                Case 1
                    objFilhosFamilia.iH1 = objFamiliasInfo.iValor
                    
                Case 2
                    objFilhosFamilia.iH2 = objFamiliasInfo.iValor
                    
                Case 3
                    objFilhosFamilia.iLE = objFamiliasInfo.iValor
                    
                Case 20
                    objFilhosFamilia.iLR = objFamiliasInfo.iValor
                    
                Case 21
                    objFilhosFamilia.iPA = objFamiliasInfo.iValor
                    
                Case 22
                    objFilhosFamilia.iRE = objFamiliasInfo.iValor
                    
                Case 5
                    objFilhosFamilia.iSH = objFamiliasInfo.iValor
                    
                Case 10
                    objFilhosFamilia.iSI = objFamiliasInfo.iValor
                    
                Case 6
                    objFilhosFamilia.iTH = objFamiliasInfo.iValor
                    
                Case 23
                    objFilhosFamilia.iVF = objFamiliasInfo.iValor
                    
            End Select
            
        Next
            
    Next

    Move_Info_Memoria = SUCESSO

    Exit Function

Erro_Move_Info_Memoria:

    Move_Info_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159968)

    End Select

    Exit Function

End Function

Function Move_Enderecos_Memoria(ByVal objFamilias As ClassFamilias) As Long
'move itens do Grid

Dim lErro As Long
Dim iIndice As Integer
Dim colEnderecos As New Collection

On Error GoTo Erro_Move_Enderecos_Memoria

    lErro = Le_Dados_Enderecos(colEnderecos)
    If lErro <> SUCESSO Then gError 140926
    
    Set objFamilias.objEnderecoRes = colEnderecos.Item(1)
    Set objFamilias.objEnderecoCom = colEnderecos.Item(2)
    Set objFamilias.objEnderecoComConj = colEnderecos.Item(3)
    
    Move_Enderecos_Memoria = SUCESSO

    Exit Function

Erro_Move_Enderecos_Memoria:

    Move_Enderecos_Memoria = gErr

    Select Case gErr
    
        Case 140926

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159969)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Enderecos(colEndereco As Collection) As Long
'Lê os dados relativos ao endereco e coloca em colEndereco

Dim objEndereco As ClassEndereco
Dim iIndice As Integer
Dim iEstadoPreenchido As Integer

On Error GoTo Erro_Le_Dados_Enderecos

    'Verifica se tem algum estado Preenchido
    For iIndice = 2 To 0 Step -1
        
        If Len(Trim(Estado(iIndice).Text)) > 0 Then
            iEstadoPreenchido = iIndice
        End If
    
    Next
    
    'Para os 3 endereços,
    For iIndice = 0 To 2
    
        Set objEndereco = New ClassEndereco
    
        'Preenche objEndereco com os dados do Endereço
        objEndereco.sEndereco = Trim(Endereco(iIndice).Text)
        objEndereco.sBairro = Trim(Bairro(iIndice).Text)
        objEndereco.sCidade = Trim(Cidade(iIndice).Text)
        objEndereco.sCEP = Trim(CEP(iIndice).Text)
    
        'Se o Endereco não estiver Preenchido --> Seta o Estado que esta Preenchido em Algum dos Frames
        If Len(Trim(Endereco(iIndice).Text)) > 0 Then
            objEndereco.iCodigoPais = Codigo_Extrai(Pais(iIndice).Text)
            objEndereco.sSiglaEstado = Trim(Estado(iIndice).Text)
            If objEndereco.iCodigoPais = PAIS_BRASIL And Estado(iIndice).ListIndex = -1 Then gError 140923
        Else
            objEndereco.iCodigoPais = Codigo_Extrai(Pais(iEstadoPreenchido).Text)
            If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
            objEndereco.sSiglaEstado = Trim(Estado(iEstadoPreenchido).Text)
        End If
    
        objEndereco.sTelefone1 = Trim(Telefone1(iIndice).Text)
        objEndereco.sTelefone2 = Trim(Telefone2(iIndice).Text)
        objEndereco.sFax = Trim(Fax(iIndice).Text)
        objEndereco.sEmail = Trim(Email(iIndice).Text)
        objEndereco.sContato = Trim(Contato(iIndice).Text)
    
        'Adiciona objEndereco na coleção
        colEndereco.Add objEndereco
    
    Next

    Le_Dados_Enderecos = SUCESSO

    Exit Function
    
Erro_Le_Dados_Enderecos:

    Le_Dados_Enderecos = gErr

    Select Case gErr
        
        Case 140923
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, Estado(iIndice).Text)

        Case 140924
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_SELECIONADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159970)

    End Select

    Exit Function

End Function

Public Sub Estado_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Estado(Index).Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Estado
    If Estado(Index).Text = Estado(Index).List(Estado(Index).ListIndex) Then Exit Sub

    'Verifica se existe o item no Estado, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(Estado(Index))
    If lErro <> SUCESSO And lErro <> 58583 Then gError 140928

    Exit Sub

Erro_Estado_Validate:

    Cancel = True

    Select Case Err

        Case 140928

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159971)
    
    End Select

    Exit Sub

End Sub

Public Sub Estado_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estado_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(Pais(Index).Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox Pais
    If Pais(Index).Text = Pais(Index).List(Pais(Index).ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Pais(Index), iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 140929

    'Nao existe o item com o CODIGO na List da ComboBox
    If lErro = 6730 Then gError 140930

    'Nao existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 140931

    Exit Sub

Erro_Pais_Validate:

    Cancel = True

    Select Case gErr

        Case 140931
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(Pais(Index).Text))

        Case 140929  'Tratado na rotina chamada

        Case 140930
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159972)

    End Select

    Exit Sub

End Sub

Public Sub Pais_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Email_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Endereco_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Fax_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Bairro_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cidade_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CEP_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Telefone1_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Telefone2_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Contato_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LabelCidade_Click(Index As Integer)

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    iEndereco = Index

    objCidade.sDescricao = Cidade(Index).Text
    
    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)

End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1

    If objCidade Is Nothing Then
        Cidade(iEndereco).Text = ""
    Else
        Cidade(iEndereco).Text = CStr(objCidade.sDescricao)
    End If

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159973)

    End Select

    Exit Sub

End Sub

Public Sub Cidade_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cidade_Validate

    If Len(Trim(Cidade(Index).Text)) = 0 Then Exit Sub
    
    objCidade.sDescricao = Cidade(Index).Text
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 140933
    
    If lErro <> SUCESSO Then gError 140934
    
    Exit Sub
     
Erro_Cidade_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case 140933
        
        Case 140934
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")

            If vbMsgRes = vbYes Then
    
                 Call Chama_Tela("CidadeCadastro", objCidade)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159974)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para uma Familia
    lErro = CF("Familias_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 140951
    
    CodFamilia.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 140951
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159975)
    
    End Select

    Exit Sub

End Sub

Private Sub FilhosInfo_ItemCheck(Item As Integer)

Dim lErro As Long
Dim objFilhosInfo As ClassFamiliasInfo
   
On Error GoTo Erro_FilhosInfo_Click
    
    'Se não tiver linha selecionada => Erro
    If GridFilhos.Row <> 0 Then
    
        If GridFilhos.Row <= gcolcolFilhosInfo.Count Then
        
            For Each objFilhosInfo In gcolcolFilhosInfo.Item(GridFilhos.Row)
            
                If objFilhosInfo.iCodInfo = FilhosInfo.ItemData(Item) Then
                        
                    If FilhosInfo.Selected(Item) = True Then
                        objFilhosInfo.iValor = MARCADO
                    Else
                        objFilhosInfo.iValor = DESMARCADO
                    End If
                    
                    Exit For
                
                End If
            
            Next
            
        End If
    
    End If

    Exit Sub

Erro_FilhosInfo_Click:

    Select Case gErr

        Case 140956
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159976)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Adiciona_colFilhosInfo()

Dim iIndice As Integer
Dim colFilhosInfo As New Collection
Dim objFilhosInfo As ClassFamiliasInfo

    For iIndice = 0 To FilhosInfo.ListCount - 1
    
        Set objFilhosInfo = New ClassFamiliasInfo
        
        objFilhosInfo.iCodInfo = FilhosInfo.ItemData(iIndice)
        objFilhosInfo.iValor = DESMARCADO

        colFilhosInfo.Add objFilhosInfo
        
    Next

    gcolcolFilhosInfo.Add colFilhosInfo

End Sub

Private Function Traz_Info_Tela2(ByVal colFamiliaInfo As Collection, ByVal iSeq As Integer, ByVal objListBox As ListBox) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFamiliaInfo As ClassFamiliasInfo

On Error GoTo Erro_Traz_Info_Tela2

    For Each objFamiliaInfo In colFamiliaInfo
    
        If objFamiliaInfo.iSeq = iSeq Then
            
            For iIndice = 0 To objListBox.ListCount - 1
        
                If objFamiliaInfo.iCodInfo = objListBox.ItemData(iIndice) Then
        
                    If objFamiliaInfo.iValor = MARCADO Then
                        objListBox.Selected(iIndice) = True
                    Else
                        objListBox.Selected(iIndice) = False
                    End If
                    
                    Exit For
                End If
        
            Next
        
        End If
        
    Next
        
    Traz_Info_Tela2 = SUCESSO

    Exit Function

Erro_Traz_Info_Tela2:

    Traz_Info_Tela2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159977)

    End Select

    Exit Function

End Function

Private Sub Inicializa_Tamanhos()

    'Implementado pois agora é possível ter constantes cutomizadas em função de tamanhos de campos do BD. AdmLib.ClassConsCust
    Endereco(0).MaxLength = STRING_ENDERECO
    Endereco(1).MaxLength = STRING_ENDERECO
    Endereco(2).MaxLength = STRING_ENDERECO
    Bairro(0).MaxLength = STRING_BAIRRO
    Bairro(1).MaxLength = STRING_BAIRRO
    Bairro(2).MaxLength = STRING_BAIRRO
    Cidade(0).MaxLength = STRING_CIDADE
    Cidade(1).MaxLength = STRING_CIDADE
    Cidade(2).MaxLength = STRING_CIDADE

End Sub

Private Function Carrega_Estados() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection
Dim vCodigo As Variant

On Error GoTo Erro_Carrega_Estados

    'Lê cada codigo da tabela Estados e coloca em colCodigo
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 140928

    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    For iIndice = 0 To 2
        For Each vCodigo In colCodigo
            Estado(iIndice).AddItem vCodigo
        Next
    Next

    'Escolhe Estado da FilialEmpresa
    lErro = CF("Estado_Seleciona", Estado)
    If lErro <> SUCESSO Then gError 140929

    Carrega_Estados = SUCESSO

    Exit Function

Erro_Carrega_Estados:

    Carrega_Estados = gErr

    Select Case gErr
    
        Case 140928, 140929

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159978)

    End Select

    Exit Function
    
End Function

Private Function Carrega_Paises() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Paises

    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoNome)
    If lErro <> SUCESSO Then gError 140930

    'Percorre as 3 Combos de País
    For iIndice = 0 To 2
        'Preenche cada ComboBox País com os objetos da colecao colCodigoNome
        For Each objCodigoNome In colCodigoNome
            Pais(iIndice).AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            Pais(iIndice).ItemData(Pais(iIndice).NewIndex) = objCodigoNome.iCodigo
        Next

        'Seleciona Brasil se existir
        For iIndice2 = 0 To Pais(iIndice).ListCount - 1
            If right(Pais(iIndice).List(iIndice2), 6) = "Brasil" Then
                Pais(iIndice).ListIndex = iIndice2
                Exit For
            End If
        Next
    Next

    Carrega_Paises = SUCESSO

    Exit Function

Erro_Carrega_Paises:

    Carrega_Paises = gErr

    Select Case gErr

        Case 140930

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159979)

    End Select

    Exit Function
    
End Function

Private Function Carrega_Infos() As Long

Dim lErro As Long
Dim colFamiliasTipoInfo As New Collection
Dim objFamiliasTipoInfo As ClassFamiliasTipoInfo

On Error GoTo Erro_Carrega_Infos

    'Le os tipos de Informação
    lErro = CF("FamiliasTipoInfo_Le", colFamiliasTipoInfo)
    If lErro <> SUCESSO Then gError 140945
    
    For Each objFamiliasTipoInfo In colFamiliasTipoInfo
    
        'Primeiro bit informa se é valido para o Titular
        If objFamiliasTipoInfo.iValidoPara And 1 Then
            TitularInfo.AddItem objFamiliasTipoInfo.sSigla
            TitularInfo.ItemData(TitularInfo.NewIndex) = objFamiliasTipoInfo.iCodInfo
        End If
        
        'Segundo bit informa se é valido para o Conjuge
        If objFamiliasTipoInfo.iValidoPara And 2 Then
            ConjugeInfo.AddItem objFamiliasTipoInfo.sSigla
            ConjugeInfo.ItemData(ConjugeInfo.NewIndex) = objFamiliasTipoInfo.iCodInfo
        End If
        
        'Terceiro bit informa se é valido para os Filhos
        If objFamiliasTipoInfo.iValidoPara And 4 Then
            FilhosInfo.AddItem objFamiliasTipoInfo.sSigla
            FilhosInfo.ItemData(FilhosInfo.NewIndex) = objFamiliasTipoInfo.iCodInfo
        End If
    
    Next

    Carrega_Infos = SUCESSO

    Exit Function

Erro_Carrega_Infos:

    Carrega_Infos = gErr

    Select Case gErr

        Case 140945

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159980)

    End Select

    Exit Function
    
End Function

Function Traz_Enderecos_Tela2(ByVal objEndereco As ClassEndereco, ByVal iIndex As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Enderecos_Tela2
    
    Endereco(iIndex).Text = objEndereco.sEndereco
    Bairro(iIndex).Text = objEndereco.sBairro
    Cidade(iIndex).Text = objEndereco.sCidade
    CEP(iIndex).Text = objEndereco.sCEP
    Estado(iIndex).Text = objEndereco.sSiglaEstado
    Call Estado_Validate(iIndex, bSGECancelDummy)

    If objEndereco.iCodigoPais = 0 Then
        Pais(iIndex).Text = ""
    Else
        Pais(iIndex).Text = objEndereco.iCodigoPais
        Call Pais_Validate(iIndex, bSGECancelDummy)
    End If

    Telefone1(iIndex).Text = objEndereco.sTelefone1
    Telefone2(iIndex).Text = objEndereco.sTelefone2
    Fax(iIndex).Text = objEndereco.sFax
    Email(iIndex).Text = objEndereco.sEmail
    Contato(iIndex).Text = objEndereco.sContato
    
    Traz_Enderecos_Tela2 = SUCESSO

    Exit Function

Erro_Traz_Enderecos_Tela2:

    Traz_Enderecos_Tela2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159981)

    End Select

    Exit Function

End Function

Private Sub Limpa_ListaBox(ByVal objListBox As ListBox)

Dim iIndice As Integer

    For iIndice = 0 To objListBox.ListCount - 1
        objListBox.Selected(iIndice) = False
    Next
    
End Sub

Private Sub BotaoPessoas_Click()

Dim lErro As Long
Dim objFamilias As New ClassFamilias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodFamilia_Click

    'Verifica se o CodFamilia foi preenchido
    If Len(Trim(CodFamilia.Text)) <> 0 Then

        objFamilias.lCodFamilia = CodFamilia.Text

    End If

    Call Chama_Tela("MembrosFamiliaLista", colSelecao, objFamilias, objEventoCodFamilia)

    Exit Sub

Erro_LabelCodFamilia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159982)

    End Select

    Exit Sub
    
End Sub

Private Sub Carrega_Saudacoes(ByVal objCombo As ComboBox)

    objCombo.Clear
    
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_SR
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_SRA
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_SRTA
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_RABINO
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_PROFESSOR
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_DOUTOR
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_COMENDADOR
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_MERITISSIMO_JUIZ
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_VOSSA_EXCELENCIA
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_VOSSA_MAGINIFICENCIA
    objCombo.AddItem STRING_PRONOME_TRATAMENTO_VOSSA_SENHORIA

End Sub
'###########################################################

Private Sub FilhoTel_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoTel_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoTel_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoTel_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoTel
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilhoEmail_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilhoEmail_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFilhos)

End Sub

Private Sub FilhoEmail_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFilhos)

End Sub

Private Sub FilhoEmail_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFilhos.objControle = FilhoEmail
    lErro = Grid_Campo_Libera_Foco(objGridFilhos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_FilhoTel(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FilhoTel

    Set objGridInt.objControle = FilhoTel

    If Len(Trim(FilhoTel.Text)) > 0 Then
        If GridFilhos.Row - GridFilhos.FixedRows = objGridFilhos.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Call Adiciona_colFilhosInfo
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140911

    Saida_Celula_FilhoTel = SUCESSO

    Exit Function

Erro_Saida_Celula_FilhoTel:

    Saida_Celula_FilhoTel = gErr

    Select Case gErr

        Case 140911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159961)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilhoEmail(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FilhoEmail

    Set objGridInt.objControle = FilhoEmail

    If Len(Trim(FilhoEmail.Text)) > 0 Then
        If GridFilhos.Row - GridFilhos.FixedRows = objGridFilhos.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Call Adiciona_colFilhosInfo
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140911

    Saida_Celula_FilhoEmail = SUCESSO

    Exit Function

Erro_Saida_Celula_FilhoEmail:

    Saida_Celula_FilhoEmail = gErr

    Select Case gErr

        Case 140911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159961)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

