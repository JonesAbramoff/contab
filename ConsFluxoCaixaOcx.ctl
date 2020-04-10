VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ConsFluxoCaixaOcx 
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   LockControls    =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   8640
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   750
      Width           =   8205
      Begin VB.Label Label10 
         Height          =   270
         Left            =   1290
         Top             =   990
         Width           =   4785
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "descricao"
      End
      Begin VB.Label Label8 
         Height          =   240
         Left            =   2775
         Top             =   2655
         Width           =   885
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "dd/mm/aa"
      End
      Begin VB.Label Label7 
         Height          =   270
         Left            =   225
         Top             =   2640
         Width           =   2550
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Dados Reais Atualizados até:"
      End
      Begin VB.Label Label3 
         Height          =   270
         Left            =   210
         Top             =   1545
         Width           =   1860
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mostrar previsão até:"
      End
      Begin VB.Label Label1 
         Height          =   270
         Left            =   270
         Top             =   990
         Width           =   990
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Descrição:"
      End
      Begin VB.Label Label2 
         Height          =   270
         Left            =   255
         Top             =   450
         Width           =   1185
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Identificação:"
      End
      Begin VB.Label Label4 
         Height          =   270
         Left            =   255
         Top             =   2085
         Width           =   1005
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Data Base:"
      End
      Begin VB.Label Label5 
         Height          =   240
         Left            =   1260
         Top             =   2100
         Width           =   750
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "database"
      End
      Begin VB.Label Label9 
         Height          =   225
         Left            =   1530
         Top             =   480
         Width           =   2790
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "identificacao"
      End
      Begin VB.Label Label11 
         Height          =   270
         Left            =   2055
         Top             =   1545
         Width           =   1140
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "dd/mm/aa"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   10
      Left            =   210
      TabIndex        =   118
      Top             =   780
      Width           =   8100
      Begin VB.CommandButton Command19 
         Caption         =   "Cancelar"
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
         Left            =   6735
         TabIndex        =   123
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton Command20 
         Caption         =   "OK"
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
         Left            =   5250
         TabIndex        =   122
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Imprimir..."
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
         Left            =   3780
         TabIndex        =   121
         Top             =   210
         Width           =   1320
      End
      Begin VB.Frame Frame9 
         Height          =   540
         Left            =   210
         TabIndex        =   132
         Top             =   630
         Width           =   4005
         Begin VB.OptionButton Option8 
            Caption         =   "Por Caixa / Conta"
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
            Left            =   225
            TabIndex        =   124
            Top             =   210
            Width           =   2010
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Cheque"
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
            Left            =   2625
            TabIndex        =   125
            Top             =   210
            Width           =   1170
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaCheque 
         Height          =   2160
         Left            =   210
         TabIndex        =   130
         Top             =   1395
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   3810
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox ContaCheque 
         Height          =   225
         Left            =   945
         TabIndex        =   126
         Top             =   1200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorCheque 
         Height          =   225
         Left            =   4410
         TabIndex        =   129
         Top             =   1200
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
      Begin MSMask.MaskEdBox Sequencial 
         Height          =   225
         Left            =   1785
         TabIndex        =   127
         Top             =   1200
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   300
         Left            =   1575
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   300
         Left            =   465
         TabIndex        =   119
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown12 
         Height          =   300
         Left            =   3330
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   300
         Left            =   2190
         TabIndex        =   120
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   225
         Left            =   3150
         TabIndex        =   128
         Top             =   1260
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
      Begin VB.Label Label35 
         Height          =   225
         Left            =   1995
         Top             =   3675
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label37 
         Height          =   255
         Left            =   2715
         Top             =   3675
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   105
         Top             =   255
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "De:"
      End
      Begin VB.Label Label62 
         Height          =   255
         Left            =   1890
         Top             =   255
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "até"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   9
      Left            =   210
      TabIndex        =   105
      Top             =   750
      Width           =   8100
      Begin VB.ComboBox Ordena 
         Height          =   315
         Left            =   1575
         TabIndex        =   117
         Text            =   "Ordenadas"
         Top             =   3615
         Width           =   4575
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Cancelar"
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
         Left            =   6720
         TabIndex        =   110
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton Command17 
         Caption         =   "OK"
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
         Left            =   5355
         TabIndex        =   109
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Imprimir..."
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
         Left            =   3990
         TabIndex        =   108
         Top             =   210
         Width           =   1320
      End
      Begin VB.Frame Frame6 
         Caption         =   "Ped. Compras por"
         Height          =   540
         Left            =   105
         TabIndex        =   135
         Top             =   630
         Width           =   3795
         Begin VB.OptionButton Option6 
            Caption         =   "Tipo de Fornecedor"
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
            Left            =   210
            TabIndex        =   111
            Top             =   210
            Width           =   2010
         End
         Begin VB.OptionButton Option7 
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
            Height          =   225
            Left            =   2310
            TabIndex        =   112
            Top             =   210
            Width           =   1380
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaCompras 
         Height          =   1890
         Left            =   105
         TabIndex        =   116
         Top             =   1260
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   3334
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox FornecedorCompra 
         Height          =   225
         Left            =   2625
         TabIndex        =   113
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorCompra 
         Height          =   225
         Left            =   5145
         TabIndex        =   115
         Top             =   1050
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox FilialCompra 
         Height          =   225
         Left            =   3885
         TabIndex        =   114
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown10 
         Height          =   300
         Left            =   1575
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   300
         Left            =   480
         TabIndex        =   106
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown11 
         Height          =   300
         Left            =   3315
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   300
         Left            =   2205
         TabIndex        =   107
         Top             =   195
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label55 
         Height          =   300
         Left            =   210
         Top             =   3615
         Width           =   1380
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ordenadas por:"
      End
      Begin VB.Label Label28 
         Height          =   225
         Left            =   2100
         Top             =   3255
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label34 
         Height          =   255
         Left            =   2820
         Top             =   3255
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label60 
         Height          =   255
         Left            =   105
         Top             =   255
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "De:"
      End
      Begin VB.Label Label61 
         Height          =   255
         Left            =   1890
         Top             =   255
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "até"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   8
      Left            =   210
      TabIndex        =   92
      Top             =   750
      Width           =   8205
      Begin VB.ComboBox Orden 
         Height          =   315
         Left            =   1695
         TabIndex        =   104
         Text            =   "Ordenadas"
         Top             =   3630
         Width           =   4575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Cancelar"
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
         Left            =   6750
         TabIndex        =   97
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton Command14 
         Caption         =   "OK"
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
         Left            =   5355
         TabIndex        =   96
         Top             =   225
         Width           =   1320
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Imprimir..."
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
         Left            =   3915
         TabIndex        =   95
         Top             =   210
         Width           =   1320
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ped. Vendas por"
         Height          =   540
         Left            =   105
         TabIndex        =   138
         Top             =   630
         Width           =   3375
         Begin VB.OptionButton Option2 
            Caption         =   "Tipo de Cliente"
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
            Left            =   210
            TabIndex        =   98
            Top             =   210
            Width           =   1695
         End
         Begin VB.OptionButton Option5 
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
            Height          =   225
            Left            =   1995
            TabIndex        =   99
            Top             =   210
            Width           =   1170
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaVendas 
         Height          =   2025
         Left            =   0
         TabIndex        =   103
         Top             =   1260
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   3572
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox ClienteVenda 
         Height          =   225
         Left            =   1785
         TabIndex        =   100
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorVenda 
         Height          =   225
         Left            =   4200
         TabIndex        =   102
         Top             =   1050
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox FilialVenda 
         Height          =   225
         Left            =   2940
         TabIndex        =   101
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown8 
         Height          =   300
         Left            =   1590
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   300
         Left            =   465
         TabIndex        =   93
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown9 
         Height          =   300
         Left            =   3315
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   300
         Left            =   2205
         TabIndex        =   94
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label56 
         Height          =   300
         Left            =   255
         Top             =   3630
         Width           =   1380
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ordenadas por:"
      End
      Begin VB.Label Label29 
         Height          =   225
         Left            =   1995
         Top             =   3315
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label32 
         Height          =   255
         Left            =   2715
         Top             =   3315
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label53 
         Height          =   255
         Left            =   105
         Top             =   255
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "De:"
      End
      Begin VB.Label Label54 
         Height          =   255
         Left            =   1890
         Top             =   255
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "até"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   7
      Left            =   225
      TabIndex        =   72
      Top             =   750
      Width           =   8190
      Begin VB.Frame Frame7 
         Caption         =   "Pagto de Comissões:"
         Height          =   1485
         Left            =   105
         TabIndex        =   141
         Top             =   2520
         Width           =   7935
         Begin MSFlexGridLib.MSFlexGrid GridFCaixaComissoes 
            Height          =   765
            Left            =   105
            TabIndex        =   91
            Top             =   315
            Width           =   7320
            _ExtentX        =   12912
            _ExtentY        =   1349
            _Version        =   393216
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   225
            Left            =   1365
            TabIndex        =   88
            Top             =   210
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorReal 
            Height          =   225
            Left            =   4410
            TabIndex        =   90
            Top             =   210
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox ValorProjetado 
            Height          =   225
            Left            =   2520
            TabIndex        =   89
            Top             =   210
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin VB.Label Label36 
            Height          =   255
            Left            =   2715
            Top             =   1155
            Width           =   1155
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label63 
            Height          =   255
            Left            =   3915
            Top             =   1155
            Width           =   1155
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label64 
            Height          =   225
            Left            =   1995
            Top             =   1155
            Width           =   705
            BackColor       =   12632256
            ForeColor       =   0
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Totais:"
         End
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Imprimir..."
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
         Left            =   3810
         TabIndex        =   75
         Top             =   105
         Width           =   1320
      End
      Begin VB.CommandButton Command23 
         Caption         =   "OK"
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
         Left            =   5250
         TabIndex        =   76
         Top             =   105
         Width           =   1320
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Cancelar"
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
         Left            =   6645
         TabIndex        =   77
         Top             =   105
         Width           =   1320
      End
      Begin VB.Frame Frame11 
         Caption         =   "Base"
         Height          =   2010
         Left            =   150
         TabIndex        =   143
         Top             =   420
         Width           =   7935
         Begin VB.Frame Frame10 
            Height          =   540
            Left            =   210
            TabIndex        =   131
            Top             =   210
            Width           =   5055
            Begin VB.OptionButton Option13 
               Caption         =   "Tipo de Vendedor"
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
               Left            =   210
               TabIndex        =   78
               Top             =   210
               Width           =   2010
            End
            Begin VB.OptionButton Option12 
               Caption         =   "Vendedor"
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
               Left            =   2310
               TabIndex        =   79
               Top             =   210
               Width           =   1170
            End
            Begin VB.OptionButton Option14 
               Caption         =   "Título"
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
               TabIndex        =   80
               Top             =   210
               Width           =   1170
            End
         End
         Begin MSFlexGridLib.MSFlexGrid GridFCaixaBase 
            Height          =   765
            Left            =   105
            TabIndex        =   87
            Top             =   840
            Width           =   7320
            _ExtentX        =   12912
            _ExtentY        =   1349
            _Version        =   393216
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   210
            TabIndex        =   81
            Top             =   735
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox NumDoc 
            Height          =   225
            Left            =   2520
            TabIndex        =   83
            Top             =   735
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox EB 
            Height          =   225
            Left            =   1365
            TabIndex        =   82
            Top             =   735
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ParcelaBase 
            Height          =   225
            Left            =   3885
            TabIndex        =   84
            Top             =   735
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   6195
            TabIndex        =   86
            Top             =   735
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   5040
            TabIndex        =   85
            Top             =   735
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin VB.Label Label65 
            Height          =   255
            Left            =   2715
            Top             =   1680
            Width           =   1155
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label66 
            Height          =   255
            Left            =   3915
            Top             =   1680
            Width           =   1155
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label67 
            Height          =   225
            Left            =   1995
            Top             =   1680
            Width           =   705
            BackColor       =   12632256
            ForeColor       =   0
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Totais:"
         End
      End
      Begin MSComCtl2.UpDown UpDown13 
         Height          =   300
         Left            =   1470
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   105
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox12 
         Height          =   300
         Left            =   360
         TabIndex        =   73
         Top             =   105
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown14 
         Height          =   300
         Left            =   3210
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   105
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   300
         Left            =   2100
         TabIndex        =   74
         Top             =   105
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label40 
         Height          =   255
         Left            =   1785
         Top             =   150
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "até"
      End
      Begin VB.Label Label68 
         Height          =   255
         Left            =   0
         Top             =   150
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "De:"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   6
      Left            =   255
      TabIndex        =   59
      Top             =   765
      Width           =   8175
      Begin VB.Frame Frame4 
         Caption         =   "Resgates por"
         Height          =   540
         Left            =   0
         TabIndex        =   146
         Top             =   525
         Width           =   3690
         Begin VB.OptionButton Option3 
            Caption         =   "Aplicação"
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
            Left            =   2415
            TabIndex        =   66
            Top             =   210
            Width           =   1170
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Tipo de Aplicação"
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
            Left            =   210
            TabIndex        =   65
            Top             =   210
            Width           =   2220
         End
      End
      Begin VB.ComboBox Ord 
         Height          =   315
         Left            =   1785
         TabIndex        =   71
         Text            =   "Ordenadas"
         Top             =   3630
         Width           =   4575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cancelar"
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
         Left            =   6750
         TabIndex        =   64
         Top             =   105
         Width           =   1320
      End
      Begin VB.CommandButton Command11 
         Caption         =   "OK"
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
         Left            =   5355
         TabIndex        =   63
         Top             =   120
         Width           =   1320
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Imprimir..."
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
         Left            =   3915
         TabIndex        =   62
         Top             =   105
         Width           =   1320
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaAplic 
         Height          =   1770
         Left            =   0
         TabIndex        =   70
         Top             =   1470
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   3122
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox Aplicacao 
         Height          =   225
         Left            =   435
         TabIndex        =   67
         Top             =   1365
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorPrevistoResgate 
         Height          =   225
         Left            =   4515
         TabIndex        =   69
         Top             =   1365
         Width           =   1470
         _ExtentX        =   2593
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
      Begin MSMask.MaskEdBox SaldoAplicado 
         Height          =   225
         Left            =   1890
         TabIndex        =   68
         Top             =   1380
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   1575
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   105
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   300
         Left            =   465
         TabIndex        =   60
         Top             =   105
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown7 
         Height          =   300
         Left            =   3315
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   105
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   300
         Left            =   2205
         TabIndex        =   61
         Top             =   105
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label31 
         Height          =   210
         Left            =   0
         Top             =   1155
         Width           =   1785
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Resgates Previstos:"
      End
      Begin VB.Label Label57 
         Height          =   300
         Left            =   375
         Top             =   3660
         Width           =   1380
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ordenadas por:"
      End
      Begin VB.Label Label25 
         Height          =   225
         Left            =   2520
         Top             =   3315
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label26 
         Height          =   255
         Left            =   4440
         Top             =   3315
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label27 
         Height          =   255
         Left            =   3240
         Top             =   3315
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label24 
         Height          =   255
         Left            =   105
         Top             =   150
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "De:"
      End
      Begin VB.Label Label52 
         Height          =   255
         Left            =   1890
         Top             =   150
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "até"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   5
      Left            =   195
      TabIndex        =   42
      Top             =   765
      Width           =   8205
      Begin VB.ComboBox OrdenadasPor 
         Height          =   315
         Left            =   1785
         TabIndex        =   58
         Text            =   "Ordenadas"
         Top             =   3675
         Width           =   4575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Cancelar"
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
         Left            =   6645
         TabIndex        =   47
         Top             =   105
         Width           =   1320
      End
      Begin VB.CommandButton Command8 
         Caption         =   "OK"
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
         Left            =   5250
         TabIndex        =   46
         Top             =   105
         Width           =   1320
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir..."
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
         Left            =   3810
         TabIndex        =   45
         Top             =   105
         Width           =   1320
      End
      Begin VB.Frame Frame3 
         Caption         =   "Recebimentos por"
         Height          =   540
         Left            =   105
         TabIndex        =   149
         Top             =   510
         Width           =   5265
         Begin VB.OptionButton Option1 
            Caption         =   "Título"
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
            Left            =   3780
            TabIndex        =   50
            Top             =   210
            Width           =   1170
         End
         Begin VB.OptionButton Client 
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
            Height          =   225
            Left            =   2310
            TabIndex        =   49
            Top             =   210
            Width           =   1380
         End
         Begin VB.OptionButton TipoCli 
            Caption         =   "Tipo de Cliente"
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
            Left            =   210
            TabIndex        =   48
            Top             =   210
            Width           =   2010
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaRec 
         Height          =   1980
         Left            =   105
         TabIndex        =   57
         Top             =   1155
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   3493
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox ParcelaRec 
         Height          =   225
         Left            =   4515
         TabIndex        =   55
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   225
         Left            =   315
         TabIndex        =   51
         Top             =   1050
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox NumTituloRec 
         Height          =   225
         Left            =   3255
         TabIndex        =   54
         Top             =   1050
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
      Begin MSMask.MaskEdBox TipoTituloRec 
         Height          =   225
         Left            =   1995
         TabIndex        =   53
         Top             =   1050
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
      Begin MSMask.MaskEdBox FilialRec 
         Height          =   225
         Left            =   1260
         TabIndex        =   52
         Top             =   1050
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox ValorRec 
         Height          =   225
         Left            =   5670
         TabIndex        =   56
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1575
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   105
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   300
         Left            =   465
         TabIndex        =   43
         Top             =   105
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown6 
         Height          =   300
         Left            =   3315
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   105
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   300
         Left            =   2205
         TabIndex        =   44
         Top             =   105
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label58 
         Height          =   300
         Left            =   330
         Top             =   3675
         Width           =   1485
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ordenadas por:"
      End
      Begin VB.Label Label18 
         Height          =   225
         Left            =   1680
         Top             =   3255
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   3600
         Top             =   3255
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   2400
         Top             =   3255
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label20 
         Height          =   255
         Left            =   105
         Top             =   150
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "De:"
      End
      Begin VB.Label Label33 
         Height          =   255
         Left            =   1890
         Top             =   150
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "até"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   4
      Left            =   210
      TabIndex        =   25
      Top             =   750
      Width           =   8205
      Begin VB.Frame Frame2 
         Caption         =   "Pagamentos por"
         Height          =   540
         Left            =   105
         TabIndex        =   152
         Top             =   630
         Width           =   5265
         Begin VB.OptionButton TipoForn 
            Caption         =   "Tipo de Fornecedor"
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
            Left            =   210
            TabIndex        =   31
            Top             =   210
            Width           =   2220
         End
         Begin VB.OptionButton Forn 
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
            Height          =   225
            Left            =   2430
            TabIndex        =   32
            Top             =   210
            Width           =   2010
         End
         Begin VB.OptionButton Titulo 
            Caption         =   "Título"
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
            Left            =   3990
            TabIndex        =   33
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.ComboBox Ordenadas 
         Height          =   315
         Left            =   1515
         TabIndex        =   41
         Text            =   "Ordenadas "
         Top             =   3660
         Width           =   4575
      End
      Begin VB.CommandButton TabAtual 
         Caption         =   "Imprimir..."
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
         Left            =   3810
         TabIndex        =   28
         Top             =   120
         Width           =   1320
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
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
         Left            =   5250
         TabIndex        =   29
         Top             =   120
         Width           =   1320
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
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
         Left            =   6660
         TabIndex        =   30
         Top             =   120
         Width           =   1320
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   300
         Left            =   1575
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   165
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaPagto 
         Height          =   1950
         Left            =   105
         TabIndex        =   40
         Top             =   1260
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   3440
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox DataPagto 
         Height          =   300
         Left            =   465
         TabIndex        =   26
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   225
         Left            =   4620
         TabIndex        =   38
         Top             =   1155
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   225
         Left            =   210
         TabIndex        =   34
         Top             =   1155
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox NumTitulo 
         Height          =   225
         Left            =   3360
         TabIndex        =   37
         Top             =   1155
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
      Begin MSMask.MaskEdBox TipoTitulo 
         Height          =   225
         Left            =   2100
         TabIndex        =   36
         Top             =   1155
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
      Begin MSMask.MaskEdBox Filial 
         Height          =   225
         Left            =   1260
         TabIndex        =   35
         Top             =   1155
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox Valor 
         Height          =   225
         Left            =   5670
         TabIndex        =   39
         Top             =   1155
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown5 
         Height          =   300
         Left            =   3315
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   300
         Left            =   2205
         TabIndex        =   27
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   105
         Top             =   210
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "De:"
      End
      Begin VB.Label Label59 
         Height          =   300
         Left            =   105
         Top             =   3690
         Width           =   1380
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ordenadas por:"
      End
      Begin VB.Label Label14 
         Height          =   225
         Left            =   1785
         Top             =   3255
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label16 
         Height          =   255
         Left            =   3705
         Top             =   3255
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label17 
         Height          =   255
         Left            =   2505
         Top             =   3255
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label30 
         Height          =   255
         Left            =   1890
         Top             =   210
         Width           =   360
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "até"
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   3
      Left            =   210
      TabIndex        =   17
      Top             =   750
      Width           =   8100
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
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
         Left            =   6615
         TabIndex        =   20
         Top             =   150
         Width           =   1320
      End
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
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
         Left            =   5145
         TabIndex        =   19
         Top             =   150
         Width           =   1320
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Imprimir..."
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
         Left            =   3675
         TabIndex        =   18
         Top             =   150
         Width           =   1320
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaIdent 
         Height          =   2565
         Left            =   105
         TabIndex        =   24
         Top             =   1005
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   4524
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox Conta 
         Height          =   225
         Left            =   1485
         TabIndex        =   21
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox SaldoReal 
         Height          =   225
         Left            =   3990
         TabIndex        =   23
         Top             =   840
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
      Begin MSMask.MaskEdBox SaldoSistema 
         Height          =   225
         Left            =   2310
         TabIndex        =   22
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label12 
         Height          =   195
         Left            =   105
         Top             =   840
         Width           =   1185
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Saldo Inicial:"
      End
      Begin VB.Label Label15 
         Height          =   225
         Left            =   1995
         Top             =   3675
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label22 
         Height          =   255
         Left            =   3915
         Top             =   3675
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label23 
         Height          =   255
         Left            =   2715
         Top             =   3675
         Width           =   1155
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   2
      Left            =   210
      TabIndex        =   1
      Top             =   750
      Width           =   8160
      Begin VB.CommandButton Command1 
         Caption         =   "Visualizar Gráficamente..."
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
         Left            =   5625
         TabIndex        =   5
         Top             =   165
         Width           =   2310
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Imprimir..."
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
         Left            =   4200
         TabIndex        =   4
         Top             =   165
         Width           =   1320
      End
      Begin VB.Frame Frame8 
         Height          =   540
         Left            =   315
         TabIndex        =   142
         Top             =   105
         Width           =   3165
         Begin VB.OptionButton Option9 
            Caption         =   "Revisão"
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
            Left            =   1680
            TabIndex        =   3
            Top             =   210
            Width           =   1170
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Projeção"
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
            Left            =   210
            TabIndex        =   2
            Top             =   210
            Width           =   1275
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridFCaixaSintetico 
         Height          =   2310
         Left            =   105
         TabIndex        =   16
         Top             =   1260
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   4075
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox DataSint 
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   1155
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox RecSistema 
         Height          =   225
         Left            =   1035
         TabIndex        =   7
         Top             =   1155
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox RecAjustado 
         Height          =   225
         Left            =   1890
         TabIndex        =   8
         Top             =   1155
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox RecPercentual 
         Height          =   225
         Left            =   2850
         TabIndex        =   9
         Top             =   1155
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PagSistema 
         Height          =   225
         Left            =   3255
         TabIndex        =   10
         Top             =   1155
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PagAjustado 
         Height          =   225
         Left            =   4200
         TabIndex        =   11
         Top             =   1155
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PagPercentual 
         Height          =   225
         Left            =   4935
         TabIndex        =   12
         Top             =   1155
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox SaldoAjustado 
         Height          =   225
         Left            =   6300
         TabIndex        =   14
         Top             =   1155
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox SaldoSist 
         Height          =   225
         Left            =   5355
         TabIndex        =   13
         Top             =   1155
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox SaldoPercentual 
         Height          =   225
         Left            =   7245
         TabIndex        =   15
         Top             =   1155
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin VB.Label Label38 
         Height          =   225
         Left            =   105
         Top             =   3675
         Width           =   705
         BackColor       =   12632256
         ForeColor       =   0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Totais:"
      End
      Begin VB.Label Label42 
         Height          =   330
         Left            =   1890
         Top             =   735
         Width           =   1290
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "Recebimentos"
      End
      Begin VB.Label Label43 
         Height          =   330
         Left            =   3885
         Top             =   735
         Width           =   1125
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "Pagamentos"
      End
      Begin VB.Label Label44 
         Height          =   330
         Left            =   6405
         Top             =   735
         Width           =   600
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "Saldo"
      End
      Begin VB.Label Label49 
         Height          =   255
         Left            =   7140
         Top             =   3675
         Width           =   525
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label50 
         Height          =   255
         Left            =   5460
         Top             =   3675
         Width           =   735
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label51 
         Height          =   255
         Left            =   6300
         Top             =   3675
         Width           =   735
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label39 
         Height          =   255
         Left            =   4725
         Top             =   3675
         Width           =   525
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label41 
         Height          =   255
         Left            =   3045
         Top             =   3675
         Width           =   735
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label45 
         Height          =   255
         Left            =   3885
         Top             =   3675
         Width           =   735
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label46 
         Height          =   255
         Left            =   2415
         Top             =   3675
         Width           =   525
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label47 
         Height          =   255
         Left            =   735
         Top             =   3675
         Width           =   735
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label48 
         Height          =   255
         Left            =   1575
         Top             =   3675
         Width           =   735
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4800
      Left            =   135
      TabIndex        =   155
      Top             =   120
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   8467
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sintético"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Saldos Iniciais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagamentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recebimentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicações"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ped.Vendas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ped.Compras"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cheque-Pré"
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
Attribute VB_Name = "ConsFluxoCaixaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim iGrid_Conta_Col As Integer
Dim iGrid_SaldoSistema_Col As Integer
Dim iGrid_SaldoReal_Col As Integer
Dim iGrid_DataSint_Col As Integer
Dim iGrid_RecSistema_Col As Integer
Dim iGrid_RecAjustado_Col As Integer
Dim iGrid_RecPercentual_Col As Integer
Dim iGrid_PagSistema_Col As Integer
Dim iGrid_PagAjustado_Col As Integer
Dim iGrid_PagPercentual_Col As Integer
Dim iGrid_SaldoSist_Col As Integer
Dim iGrid_SaldoAjustado_Col As Integer
Dim iGrid_SaldoPercentual_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_TipoTitulo_Col As Integer
Dim iGrid_NumTitulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_FilialRec_Col As Integer
Dim iGrid_TipoTituloRec_Col As Integer
Dim iGrid_NumTituloRec_Col As Integer
Dim iGrid_ParcelaRec_Col As Integer
Dim iGrid_ValorRec_Col As Integer
Dim iGrid_Aplicacao_Col As Integer
Dim iGrid_SaldoAplicado_Col As Integer
Dim iGrid_ValorPrevistoResgate_Col As Integer
Dim iGrid_ValorVenda_Col As Integer
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_EB_Col As Integer
Dim iGrid_NumDoc_Col As Integer
Dim iGrid_ParcelaBase_Col As Integer
Dim iGrid_ValorBase_Col As Integer
Dim iGrid_ValorComissao_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_ValorProjetado_Col As Integer
Dim iGrid_ValorReal_Col As Integer
Dim iGrid_ClienteVenda_Col As Integer
Dim iGrid_FilialVenda_Col As Integer
Dim iGrid_ValorCompra_Col As Integer
Dim iGrid_FornecedorCompra_Col As Integer
Dim iGrid_FilialCompra_Col As Integer
Dim iGrid_ContaCheque_Col As Integer
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_ValorCheque_Col As Integer
Dim objGridFCaixaIdent As AdmGrid
Dim objGridFCaixaSintetico As AdmGrid
Dim objGridFCaixaPagto As AdmGrid
Dim objGridFCaixaRec As AdmGrid
Dim objGridFCaixaAplic As AdmGrid
Dim objGridFCaixaBase As AdmGrid
Dim objGridFCaixaComissoes As AdmGrid
Dim objGridFCaixaVendas As AdmGrid
Dim objGridFCaixaCompras As AdmGrid
Dim objGridFCaixaCheque As AdmGrid
Dim iFrameAtual As Integer

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Sintetico = 2
Private Const TAB_SaldosIniciais = 3
Private Const TAB_Pagamentos = 4
Private Const TAB_Recebimentos = 5
Private Const TAB_Aplicacoes = 6
Private Const TAB_Comissoes = 7
Private Const TAB_PedVendas = 8
Private Const TAB_PedCompras = 9
Private Const TAB_ChequePre = 10

Private Function Inicializa_Mascaras() As Long
'inicializa as mascaras de conta e centro de custo

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 5694
    
    Conta.Mask = sMascaraConta
          
    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 5694
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154850)
        
    End Select

    Exit Function
    
End Function


Private Sub Conta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaIdent)
     
End Sub

Private Sub Conta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaIdent)
    
End Sub

Private Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaIdent.objControle = Conta
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaIdent)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGridFCaixaIdent = Nothing
    Set objGridFCaixaSintetico = Nothing
    Set objGridFCaixaPagto = Nothing
    Set objGridFCaixaRec = Nothing
    Set objGridFCaixaAplic = Nothing
    Set objGridFCaixaBase = Nothing
    Set objGridFCaixaComissoes = Nothing
    Set objGridFCaixaVendas = Nothing
    Set objGridFCaixaCompras = Nothing
    Set objGridFCaixaCheque = Nothing

End Sub

Private Sub DataPagto_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataPagto, iAlterado)

End Sub

Private Sub MaskEdBox1_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox1, iAlterado)

End Sub

Private Sub MaskEdBox10_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox10, iAlterado)

End Sub

Private Sub MaskEdBox11_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox11, iAlterado)

End Sub

Private Sub MaskEdBox12_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox12, iAlterado)

End Sub

Private Sub MaskEdBox13_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox13, iAlterado)

End Sub

Private Sub MaskEdBox2_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox2, iAlterado)

End Sub

Private Sub MaskEdBox3_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox3, iAlterado)

End Sub

Private Sub MaskEdBox4_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskEdBox4, iAlterado)

End Sub

Private Sub MaskEdBox5_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox5, iAlterado)

End Sub

Private Sub MaskEdBox6_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox6, iAlterado)

End Sub

Private Sub MaskEdBox7_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox7, iAlterado)

End Sub

Private Sub MaskEdBox8_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox8, iAlterado)

End Sub

Private Sub MaskEdBox9_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskEdBox9, iAlterado)

End Sub

Private Sub SaldoSistema_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaIdent)

End Sub

Private Sub SaldoSistema_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaIdent)
    
End Sub

Private Sub SaldoSistema_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaIdent.objControle = SaldoSistema
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaIdent)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SaldoReal_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaIdent)

End Sub

Private Sub SaldoReal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaIdent)
    
End Sub

Private Sub SaldoReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaIdent.objControle = SaldoReal
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaIdent)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaIdent_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaIdent, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaIdent, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaIdent_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaIdent)

End Sub

Private Sub GridFCaixaIdent_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaIdent, iAlterado)
    
End Sub

Private Sub GridFCaixaIdent_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaIdent)
    
End Sub

Private Sub GridFCaixaIdent_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaIdent)
    
End Sub

Private Sub GridFCaixaIdent_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaIdent, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaIdent, iAlterado)
    End If

End Sub

Private Sub GridFCaixaIdent_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaIdent)

End Sub

Private Sub GridFCaixaIdent_RowColChange()

    Call Grid_RowColChange(objGridFCaixaIdent)
       
End Sub

Private Sub GridFCaixaIdent_Scroll()

    Call Grid_Scroll(objGridFCaixaIdent)
    
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objGridFCaixaIdent = New AdmGrid
    Set objGridFCaixaSintetico = New AdmGrid
    Set objGridFCaixaPagto = New AdmGrid
    Set objGridFCaixaRec = New AdmGrid
    Set objGridFCaixaAplic = New AdmGrid
    Set objGridFCaixaBase = New AdmGrid
    Set objGridFCaixaComissoes = New AdmGrid
    Set objGridFCaixaVendas = New AdmGrid
    Set objGridFCaixaCompras = New AdmGrid
    Set objGridFCaixaCheque = New AdmGrid
    
    lErro = Inicializa_Grid_FCaixaIdent(objGridFCaixaIdent)
    If lErro <> SUCESSO Then Error 14250

    lErro = Inicializa_Grid_FCaixaSintetico(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Error 14250
    
    lErro = Inicializa_Grid_FCaixaPagto(objGridFCaixaPagto)
    If lErro <> SUCESSO Then Error 14250

    lErro = Inicializa_Grid_FCaixaRec(objGridFCaixaRec)
    If lErro <> SUCESSO Then Error 14250
    
    lErro = Inicializa_Grid_FCaixaAplic(objGridFCaixaAplic)
    If lErro <> SUCESSO Then Error 14250
    
    lErro = Inicializa_Grid_FCaixaBase(objGridFCaixaBase)
    If lErro <> SUCESSO Then Error 14250
    
    lErro = Inicializa_Grid_FCaixaComissoes(objGridFCaixaComissoes)
    If lErro <> SUCESSO Then Error 14250
    
    lErro = Inicializa_Grid_FCaixaVendas(objGridFCaixaVendas)
    If lErro <> SUCESSO Then Error 14250
    
    lErro = Inicializa_Grid_FCaixaCompras(objGridFCaixaCompras)
    If lErro <> SUCESSO Then Error 14250
    
    lErro = Inicializa_Grid_FCaixaCheque(objGridFCaixaCheque)
    If lErro <> SUCESSO Then Error 14250
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 14250
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154851)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_FCaixaIdent(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaIdent

    'tela em questão
    Set objGridFCaixaIdent.objForm = Me
        
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    objGridInt.colColuna.Add ("Saldo Sistema")
    objGridInt.colColuna.Add ("Saldo Real")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    objGridInt.colCampo.Add (SaldoSistema.Name)
    objGridInt.colCampo.Add (SaldoReal.Name)
    
    iGrid_Conta_Col = 1
    iGrid_SaldoSistema_Col = 2
    iGrid_SaldoReal_Col = 3
   
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 14251

    objGridInt.objGrid = GridFCaixaIdent
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
        
    GridFCaixaIdent.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaIdent = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaIdent:

    Inicializa_Grid_FCaixaIdent = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154852)
        
    End Select

    Exit Function
        
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
        
    End Select

    Exit Function

End Function

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        Select Case iFrameAtual
            
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_FLUXO_CAIXA_ID
            
            Case TAB_Sintetico
                Parent.HelpContextID = IDH_FLUXO_CAIXA_SINTETICO
            
            Case TAB_SaldosIniciais
                Parent.HelpContextID = IDH_FLUXO_CAIXA_SALDADOSINICIAIS
            
            Case TAB_Pagamentos
                Parent.HelpContextID = IDH_FLUXO_CAIXA_PAGAMENTOS
            
            Case TAB_Recebimentos
                Parent.HelpContextID = IDH_FLUXO_CAIXA_RECEBIMENTOS
            
            Case TAB_Aplicacoes
                Parent.HelpContextID = IDH_FLUXO_CAIXA_APLICACOES
            
            Case TAB_Comissoes
                Parent.HelpContextID = IDH_FLUXO_CAIXA_COMISSOES
            
            Case TAB_PedVendas
                Parent.HelpContextID = IDH_FLUXO_CAIXA_PEDVENDAS
            
            Case TAB_PedCompras
                Parent.HelpContextID = IDH_FLUXO_CAIXA_PEDCOMPRAS
            
            Case TAB_ChequePre
                Parent.HelpContextID = IDH_FLUXO_CAIXA_CHEQUEPRE
        
        End Select
    
    End If

End Sub

Private Sub DataSint_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub DataSint_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub DataSint_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = DataSint
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub RecSistema_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub RecSistema_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub RecSistema_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = RecSistema
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub RecAjustado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub RecAjustado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub RecAjustado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = RecAjustado
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub RecPercentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub RecPercentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub RecPercentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = RecPercentual
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PagSistema_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub PagSistema_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub PagSistema_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = PagSistema
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PagAjustado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub PagAjustado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub PagAjustado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = PagAjustado
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PagPercentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub PagPercentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub PagPercentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = PagPercentual
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SaldoSist_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub SaldoSist_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub SaldoSist_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = SaldoSist
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SaldoAjustado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub SaldoAjustado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub SaldoAjustado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = SaldoAjustado
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SaldoPercentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaSintetico)
     
End Sub

Private Sub SaldoPercentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaSintetico)
    
End Sub

Private Sub SaldoPercentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaSintetico.objControle = SaldoPercentual
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaSintetico)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaSintetico_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaSintetico, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaSintetico, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaSintetico_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaSintetico)

End Sub

Private Sub GridFCaixaSintetico_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaSintetico, iAlterado)
    
End Sub

Private Sub GridFCaixaSintetico_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaSintetico)
    
End Sub

Private Sub GridFCaixaSintetico_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaSintetico)
    
End Sub

Private Sub GridFCaixaSintetico_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaSintetico, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaSintetico, iAlterado)
    End If

End Sub

Private Sub GridFCaixaSintetico_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaSintetico)

End Sub

Private Sub GridFCaixaSintetico_RowColChange()

    Call Grid_RowColChange(objGridFCaixaSintetico)
       
End Sub

Private Sub GridFCaixaSintetico_Scroll()

    Call Grid_Scroll(objGridFCaixaSintetico)
    
End Sub

Private Function Inicializa_Grid_FCaixaSintetico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaSintetico
    
    'tela em questão
    Set objGridFCaixaSintetico.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Sistema")
    objGridInt.colColuna.Add ("Ajust.")
    objGridInt.colColuna.Add ("%")
    objGridInt.colColuna.Add ("Sistema")
    objGridInt.colColuna.Add ("Ajust.")
    objGridInt.colColuna.Add ("%")
    objGridInt.colColuna.Add ("Sistema")
    objGridInt.colColuna.Add ("Ajust.")
    objGridInt.colColuna.Add ("%")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (DataSint.Name)
    objGridInt.colCampo.Add (RecSistema.Name)
    objGridInt.colCampo.Add (RecAjustado.Name)
    objGridInt.colCampo.Add (RecPercentual.Name)
    objGridInt.colCampo.Add (PagSistema.Name)
    objGridInt.colCampo.Add (PagAjustado.Name)
    objGridInt.colCampo.Add (PagPercentual.Name)
    objGridInt.colCampo.Add (SaldoSist.Name)
    objGridInt.colCampo.Add (SaldoAjustado.Name)
    objGridInt.colCampo.Add (SaldoPercentual.Name)
        
    iGrid_DataSint_Col = 1
    iGrid_RecSistema_Col = 2
    iGrid_RecAjustado_Col = 3
    iGrid_RecPercentual_Col = 4
    iGrid_PagSistema_Col = 5
    iGrid_PagAjustado_Col = 6
    iGrid_PagPercentual_Col = 7
    iGrid_SaldoSist_Col = 8
    iGrid_SaldoAjustado_Col = 9
    iGrid_SaldoPercentual_Col = 10
    
    objGridInt.objGrid = GridFCaixaSintetico
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
        
    GridFCaixaSintetico.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaSintetico = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaSintetico:

    Inicializa_Grid_FCaixaSintetico = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154853)
        
    End Select

    Exit Function
        
End Function

Private Sub Fornecedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaPagto)
     
End Sub

Private Sub Fornecedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaPagto)
    
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaPagto.objControle = Fornecedor
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaPagto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Filial_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaPagto)
      
End Sub

Private Sub Filial_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaPagto)

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaPagto.objControle = Filial
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaPagto)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub TipoTitulo_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaPagto)

End Sub

Private Sub TipoTitulo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaPagto)
    
End Sub

Private Sub TipoTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaPagto.objControle = TipoTitulo
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaPagto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub NumTitulo_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaPagto)

End Sub

Private Sub NumTitulo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaPagto)
    
End Sub

Private Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaPagto.objControle = NumTitulo
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaPagto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Parcela_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaPagto)

End Sub

Private Sub Parcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaPagto)
    
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaPagto.objControle = Parcela
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaPagto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Valor_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaPagto)

End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaPagto)
    
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaPagto.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaPagto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaPagto_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaPagto, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaPagto, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaPagto_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaPagto)

End Sub

Private Sub GridFCaixaPagto_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaPagto, iAlterado)
    
End Sub

Private Sub GridFCaixaPagto_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaPagto)
    
End Sub

Private Sub GridFCaixaPagto_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaPagto)
    
End Sub

Private Sub GridFCaixaPagto_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaPagto, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaPagto, iAlterado)
    End If

End Sub

Private Sub GridFCaixaPagto_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaPagto)

End Sub

Private Sub GridFCaixaPagto_RowColChange()

    Call Grid_RowColChange(objGridFCaixaPagto)
       
End Sub

Private Sub GridFCaixaPagto_Scroll()

    Call Grid_Scroll(objGridFCaixaPagto)
    
End Sub

Private Function Inicializa_Grid_FCaixaPagto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaPagto
    
    'tela em questão
    Set objGridFCaixaPagto.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Tipo do Título")
    objGridInt.colColuna.Add ("Nº do Título")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (TipoTitulo.Name)
    objGridInt.colCampo.Add (NumTitulo.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Valor.Name)
        
    iGrid_Fornecedor_Col = 1
    iGrid_Filial_Col = 2
    iGrid_TipoTitulo_Col = 3
    iGrid_NumTitulo_Col = 4
    iGrid_Parcela_Col = 5
    iGrid_Valor_Col = 6
        
    objGridInt.objGrid = GridFCaixaPagto
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridFCaixaPagto.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaPagto = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaPagto:

    Inicializa_Grid_FCaixaPagto = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154854)
        
    End Select

    Exit Function
        
End Function

Private Sub Cliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaRec)
     
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaRec)
    
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaRec.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaRec)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FilialRec_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaRec)
      
End Sub

Private Sub FilialRec_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaRec)

End Sub

Private Sub FilialRec_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaRec.objControle = FilialRec
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaRec)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub TipoTituloRec_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaRec)

End Sub

Private Sub TipoTituloRec_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaRec)
    
End Sub

Private Sub TipoTituloRec_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaRec.objControle = TipoTituloRec
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaRec)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub NumTituloRec_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaRec)

End Sub

Private Sub NumTituloRec_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaRec)
    
End Sub

Private Sub NumTituloRec_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaRec.objControle = NumTituloRec
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaRec)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ParcelaRec_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaRec)

End Sub

Private Sub ParcelaRec_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaRec)
    
End Sub

Private Sub ParcelaRec_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaRec.objControle = ParcelaRec
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaRec)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorRec_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaRec)

End Sub

Private Sub ValorRec_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaRec)
    
End Sub

Private Sub ValorRec_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaRec.objControle = ValorRec
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaRec)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaRec_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaRec, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaRec, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaRec_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaRec)

End Sub

Private Sub GridFCaixaRec_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaRec, iAlterado)
    
End Sub

Private Sub GridFCaixaRec_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaRec)
    
End Sub

Private Sub GridFCaixaRec_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaRec)
    
End Sub

Private Sub GridFCaixaRec_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaRec, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaRec, iAlterado)
    End If

End Sub

Private Sub GridFCaixaRec_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaRec)

End Sub

Private Sub GridFCaixaRec_RowColChange()

    Call Grid_RowColChange(objGridFCaixaRec)
       
End Sub

Private Sub GridFCaixaRec_Scroll()

    Call Grid_Scroll(objGridFCaixaRec)
    
End Sub

Private Function Inicializa_Grid_FCaixaRec(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaRec
    
    'tela em questão
    Set objGridFCaixaRec.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Tipo do Título")
    objGridInt.colColuna.Add ("Nº do Título")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (FilialRec.Name)
    objGridInt.colCampo.Add (TipoTituloRec.Name)
    objGridInt.colCampo.Add (NumTituloRec.Name)
    objGridInt.colCampo.Add (ParcelaRec.Name)
    objGridInt.colCampo.Add (ValorRec.Name)
        
    iGrid_Cliente_Col = 1
    iGrid_FilialRec_Col = 2
    iGrid_TipoTituloRec_Col = 3
    iGrid_NumTituloRec_Col = 4
    iGrid_ParcelaRec_Col = 5
    iGrid_ValorRec_Col = 6
        
    objGridInt.objGrid = GridFCaixaRec
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridFCaixaRec.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaRec = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaRec:

    Inicializa_Grid_FCaixaRec = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154855)
        
    End Select

    Exit Function
        
End Function

Private Sub Aplicacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaAplic)
     
End Sub

Private Sub Aplicacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaAplic)
    
End Sub

Private Sub Aplicacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaAplic.objControle = Aplicacao
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaAplic)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SaldoAplicado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaAplic)
      
End Sub

Private Sub SaldoAplicado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaAplic)

End Sub

Private Sub SaldoAplicado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaAplic.objControle = SaldoAplicado
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaAplic)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub ValorPrevistoResgate_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaAplic)

End Sub

Private Sub ValorPrevistoResgate_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaAplic)
    
End Sub

Private Sub ValorPrevistoResgate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaAplic.objControle = ValorPrevistoResgate
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaAplic)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaAplic_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaAplic, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaAplic, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaAplic_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaAplic)

End Sub

Private Sub GridFCaixaAplic_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaAplic, iAlterado)
    
End Sub

Private Sub GridFCaixaAplic_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaAplic)
    
End Sub

Private Sub GridFCaixaAplic_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaAplic)
    
End Sub

Private Sub GridFCaixaAplic_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaAplic, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaAplic, iAlterado)
    End If

End Sub

Private Sub GridFCaixaAplic_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaAplic)

End Sub

Private Sub GridFCaixaAplic_RowColChange()

    Call Grid_RowColChange(objGridFCaixaAplic)
       
End Sub

Private Sub GridFCaixaAplic_Scroll()

    Call Grid_Scroll(objGridFCaixaAplic)
    
End Sub

Private Function Inicializa_Grid_FCaixaAplic(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaAplic
    
    'tela em questão
    Set objGridFCaixaAplic.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Aplicação")
    objGridInt.colColuna.Add ("Saldo Aplicado")
    objGridInt.colColuna.Add ("Valor Prev.Resgate")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (Aplicacao.Name)
    objGridInt.colCampo.Add (SaldoAplicado.Name)
    objGridInt.colCampo.Add (ValorPrevistoResgate.Name)
        
    iGrid_Aplicacao_Col = 1
    iGrid_SaldoAplicado_Col = 2
    iGrid_ValorPrevistoResgate_Col = 3
        
    objGridInt.objGrid = GridFCaixaAplic
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridFCaixaAplic.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaAplic = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaAplic:

    Inicializa_Grid_FCaixaAplic = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154856)
        
    End Select

    Exit Function
        
End Function

Private Sub Vendedor_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaBase)

End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaBase)
    
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaBase.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaBase)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub EB_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaBase)

End Sub

Private Sub EB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaBase)
    
End Sub

Private Sub EB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaBase.objControle = EB
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaBase)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub NumDoc_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaBase)

End Sub

Private Sub NumDoc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaBase)
    
End Sub

Private Sub NumDoc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaBase.objControle = NumDoc
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaBase)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ParcelaBase_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaBase)

End Sub

Private Sub ParcelaBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaBase)
    
End Sub

Private Sub ParcelaBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaBase.objControle = ParcelaBase
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaBase)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorBase_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaBase)

End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaBase)
    
End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaBase.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaBase)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorComissao_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaBase)

End Sub

Private Sub ValorComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaBase)
    
End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaBase.objControle = ValorComissao
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaBase)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaBase_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaBase, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaBase, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaBase_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaBase)

End Sub

Private Sub GridFCaixaBase_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaBase, iAlterado)
    
End Sub

Private Sub GridFCaixaBase_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaBase)
    
End Sub

Private Sub GridFCaixaBase_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaBase)
    
End Sub

Private Sub GridFCaixaBase_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaBase, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaBase, iAlterado)
    End If

End Sub

Private Sub GridFCaixaBase_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaBase)

End Sub

Private Sub GridFCaixaBase_RowColChange()

    Call Grid_RowColChange(objGridFCaixaBase)
       
End Sub

Private Sub GridFCaixaBase_Scroll()

    Call Grid_Scroll(objGridFCaixaBase)
    
End Sub

Private Function Inicializa_Grid_FCaixaBase(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaBase
    
    'tela em questão
    Set objGridFCaixaBase.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("E/B")
    objGridInt.colColuna.Add ("Nº Documento")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Valor Comissão")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Vendedor.Name)
    objGridInt.colCampo.Add (EB.Name)
    objGridInt.colCampo.Add (NumDoc.Name)
    objGridInt.colCampo.Add (ParcelaBase.Name)
    objGridInt.colCampo.Add (ValorBase.Name)
    objGridInt.colCampo.Add (ValorComissao.Name)
    
    iGrid_Vendedor_Col = 1
    iGrid_EB_Col = 2
    iGrid_NumDoc_Col = 3
    iGrid_ParcelaBase_Col = 4
    iGrid_ValorBase_Col = 5
    iGrid_ValorComissao_Col = 6
    
    objGridInt.objGrid = GridFCaixaBase
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 2
        
    GridFCaixaBase.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaBase = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaBase:

    Inicializa_Grid_FCaixaBase = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154857)
        
    End Select

    Exit Function
        
End Function

Private Sub Data_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaComissoes)

End Sub

Private Sub Data_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaComissoes)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaComissoes.objControle = Data
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorProjetado_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaComissoes)

End Sub

Private Sub ValorProjetado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaComissoes)
    
End Sub

Private Sub ValorProjetado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaComissoes.objControle = ValorProjetado
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorReal_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaComissoes)

End Sub

Private Sub ValorReal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaComissoes)
    
End Sub

Private Sub ValorReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaComissoes.objControle = ValorReal
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaComissoes_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaComissoes, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaComissoes, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaComissoes_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaComissoes)

End Sub

Private Sub GridFCaixaComissoes_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaComissoes, iAlterado)
    
End Sub

Private Sub GridFCaixaComissoes_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaComissoes)
    
End Sub

Private Sub GridFCaixaComissoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaComissoes)
    
End Sub

Private Sub GridFCaixaComissoes_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaComissoes, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaComissoes, iAlterado)
    End If

End Sub

Private Sub GridFCaixaComissoes_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaComissoes)

End Sub

Private Sub GridFCaixaComissoes_RowColChange()

    Call Grid_RowColChange(objGridFCaixaComissoes)
       
End Sub

Private Sub GridFCaixaComissoes_Scroll()

    Call Grid_Scroll(objGridFCaixaComissoes)
    
End Sub

Private Function Inicializa_Grid_FCaixaComissoes(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaComissoes
    
    'tela em questão
    Set objGridFCaixaComissoes.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor Projetado")
    objGridInt.colColuna.Add ("Valor Real")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (ValorProjetado.Name)
    objGridInt.colCampo.Add (ValorReal.Name)
    
    iGrid_Data_Col = 1
    iGrid_ValorProjetado_Col = 2
    iGrid_ValorReal_Col = 3
    
    objGridInt.objGrid = GridFCaixaComissoes
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 2
        
    GridFCaixaComissoes.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaComissoes = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaComissoes:

    Inicializa_Grid_FCaixaComissoes = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154858)
        
    End Select

    Exit Function
        
End Function

Private Sub ValorVenda_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaVendas)
      
End Sub

Private Sub ValorVenda_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaVendas)

End Sub

Private Sub ValorVenda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaVendas.objControle = ValorVenda
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaVendas)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub ClienteVenda_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaVendas)

End Sub

Private Sub ClienteVenda_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaVendas)
    
End Sub

Private Sub ClienteVenda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaVendas.objControle = ClienteVenda
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaVendas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FilialVenda_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaVendas)

End Sub

Private Sub FilialVenda_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaVendas)
    
End Sub

Private Sub FilialVenda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaVendas.objControle = FilialVenda
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaVendas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaVendas_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaVendas, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaVendas, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaVendas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaVendas)

End Sub

Private Sub GridFCaixaVendas_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaVendas, iAlterado)
    
End Sub

Private Sub GridFCaixaVendas_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaVendas)
    
End Sub

Private Sub GridFCaixaVendas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaVendas)
    
End Sub

Private Sub GridFCaixaVendas_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaVendas, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaVendas, iAlterado)
    End If

End Sub

Private Sub GridFCaixaVendas_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaVendas)

End Sub

Private Sub GridFCaixaVendas_RowColChange()

    Call Grid_RowColChange(objGridFCaixaVendas)
       
End Sub

Private Sub GridFCaixaVendas_Scroll()

    Call Grid_Scroll(objGridFCaixaVendas)
    
End Sub

Private Function Inicializa_Grid_FCaixaVendas(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaVendas
    
    'tela em questão
    Set objGridFCaixaVendas.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Valor")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (ClienteVenda.Name)
    objGridInt.colCampo.Add (FilialVenda.Name)
    objGridInt.colCampo.Add (ValorVenda.Name)
        
    iGrid_ClienteVenda_Col = 1
    iGrid_FilialVenda_Col = 2
    iGrid_ValorVenda_Col = 3
    
    objGridInt.objGrid = GridFCaixaVendas
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridFCaixaVendas.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaVendas = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaVendas:

    Inicializa_Grid_FCaixaVendas = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154859)
        
    End Select

    Exit Function
        
End Function

Private Sub ValorCompra_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaCompras)
      
End Sub

Private Sub ValorCompra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaCompras)

End Sub

Private Sub ValorCompra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaCompras.objControle = ValorCompra
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaCompras)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub FornecedorCompra_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaCompras)

End Sub

Private Sub FornecedorCompra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaCompras)
    
End Sub

Private Sub FornecedorCompra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaCompras.objControle = FornecedorCompra
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaCompras)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FilialCompra_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaCompras)

End Sub

Private Sub FilialCompra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaCompras)
    
End Sub

Private Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaCompras.objControle = FilialCompra
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaCompras)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaCompras_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaCompras, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaCompras, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaCompras_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaCompras)

End Sub

Private Sub GridFCaixaCompras_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaCompras, iAlterado)
    
End Sub

Private Sub GridFCaixaCompras_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaCompras)
    
End Sub

Private Sub GridFCaixaCompras_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaCompras)
    
End Sub

Private Sub GridFCaixaCompras_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaCompras, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaCompras, iAlterado)
    End If

End Sub

Private Sub GridFCaixaCompras_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaCompras)

End Sub

Private Sub GridFCaixaCompras_RowColChange()

    Call Grid_RowColChange(objGridFCaixaCompras)
       
End Sub

Private Sub GridFCaixaCompras_Scroll()

    Call Grid_Scroll(objGridFCaixaCompras)
    
End Sub

Private Function Inicializa_Grid_FCaixaCompras(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaCompras
    
    'tela em questão
    Set objGridFCaixaCompras.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Valor")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (FornecedorCompra.Name)
    objGridInt.colCampo.Add (FilialCompra.Name)
    objGridInt.colCampo.Add (ValorCompra.Name)
        
    iGrid_FornecedorCompra_Col = 1
    iGrid_FilialCompra_Col = 2
    iGrid_ValorCompra_Col = 3
        
    objGridInt.objGrid = GridFCaixaCompras
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridFCaixaCompras.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaCompras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaCompras:

    Inicializa_Grid_FCaixaCompras = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154860)
        
    End Select

    Exit Function
        
End Function

Private Sub ContaCheque_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaCheque)
     
End Sub

Private Sub ContaCheque_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaCheque)
    
End Sub

Private Sub ContaCheque_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaCheque.objControle = ContaCheque
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaCheque)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Sequencial_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFCaixaCheque)
      
End Sub

Private Sub Sequencial_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaCheque)

End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaCheque.objControle = Sequencial
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaCheque)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub Numero_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaCheque)

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaCheque)
    
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaCheque.objControle = Numero
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaCheque)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorCheque_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridFCaixaCheque)

End Sub

Private Sub ValorCheque_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFCaixaCheque)
    
End Sub

Private Sub ValorCheque_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFCaixaCheque.objControle = ValorCheque
    lErro = Grid_Campo_Libera_Foco(objGridFCaixaCheque)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixaCheque_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridFCaixaCheque, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaCheque, iAlterado)
    End If
    
End Sub

Private Sub GridFCaixaCheque_GotFocus()
    
    Call Grid_Recebe_Foco(objGridFCaixaCheque)

End Sub

Private Sub GridFCaixaCheque_EnterCell()
    
    Call Grid_Entrada_Celula(objGridFCaixaCheque, iAlterado)
    
End Sub

Private Sub GridFCaixaCheque_LeaveCell()
    
    Call Saida_Celula(objGridFCaixaCheque)
    
End Sub

Private Sub GridFCaixaCheque_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFCaixaCheque)
    
End Sub

Private Sub GridFCaixaCheque_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFCaixaCheque, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFCaixaCheque, iAlterado)
    End If

End Sub

Private Sub GridFCaixaCheque_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridFCaixaCheque)

End Sub

Private Sub GridFCaixaCheque_RowColChange()

    Call Grid_RowColChange(objGridFCaixaCheque)
       
End Sub

Private Sub GridFCaixaCheque_Scroll()

    Call Grid_Scroll(objGridFCaixaCheque)
    
End Sub

Private Function Inicializa_Grid_FCaixaCheque(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_FCaixaCheque
    
    'tela em questão
    Set objGridFCaixaCheque.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    objGridInt.colColuna.Add ("Sequencial")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Valor")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (ContaCheque.Name)
    objGridInt.colCampo.Add (Sequencial.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (ValorCheque.Name)
    
    iGrid_ContaCheque_Col = 1
    iGrid_Sequencial_Col = 2
    iGrid_Numero_Col = 3
    iGrid_ValorCheque_Col = 4
        
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 14251

    objGridInt.objGrid = GridFCaixaCheque
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
        
    GridFCaixaCheque.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FCaixaCheque = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_FCaixaCheque:

    Inicializa_Grid_FCaixaCheque = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154861)
        
    End Select

    Exit Function
        
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_ID
    Set Form_Load_Ocx = Me
    Caption = "Previsão de Fluxo de Caixa"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConsFluxoCaixa"
    
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




Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
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

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label62_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label62, Source, X, Y)
End Sub

Private Sub Label62_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label62, Button, Shift, X, Y)
End Sub

Private Sub Label55_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label55, Source, X, Y)
End Sub

Private Sub Label55_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label55, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label60_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label60, Source, X, Y)
End Sub

Private Sub Label60_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label60, Button, Shift, X, Y)
End Sub

Private Sub Label61_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label61, Source, X, Y)
End Sub

Private Sub Label61_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label61, Button, Shift, X, Y)
End Sub

Private Sub Label56_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label56, Source, X, Y)
End Sub

Private Sub Label56_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label56, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Label53_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label53, Source, X, Y)
End Sub

Private Sub Label53_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label53, Button, Shift, X, Y)
End Sub

Private Sub Label54_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label54, Source, X, Y)
End Sub

Private Sub Label54_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label54, Button, Shift, X, Y)
End Sub

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub Label63_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label63, Source, X, Y)
End Sub

Private Sub Label63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label63, Button, Shift, X, Y)
End Sub

Private Sub Label64_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label64, Source, X, Y)
End Sub

Private Sub Label64_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label64, Button, Shift, X, Y)
End Sub

Private Sub Label65_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label65, Source, X, Y)
End Sub

Private Sub Label65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label65, Button, Shift, X, Y)
End Sub

Private Sub Label66_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label66, Source, X, Y)
End Sub

Private Sub Label66_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label66, Button, Shift, X, Y)
End Sub

Private Sub Label67_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label67, Source, X, Y)
End Sub

Private Sub Label67_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label67, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Private Sub Label68_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label68, Source, X, Y)
End Sub

Private Sub Label68_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label68, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label57_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label57, Source, X, Y)
End Sub

Private Sub Label57_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label57, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub Label52_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label52, Source, X, Y)
End Sub

Private Sub Label52_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label52, Button, Shift, X, Y)
End Sub

Private Sub Label58_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label58, Source, X, Y)
End Sub

Private Sub Label58_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label58, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
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

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label59_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label59, Source, X, Y)
End Sub

Private Sub Label59_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label59, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label38_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label38, Source, X, Y)
End Sub

Private Sub Label38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label38, Button, Shift, X, Y)
End Sub

Private Sub Label42_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label42, Source, X, Y)
End Sub

Private Sub Label42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label42, Button, Shift, X, Y)
End Sub

Private Sub Label43_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label43, Source, X, Y)
End Sub

Private Sub Label43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label43, Button, Shift, X, Y)
End Sub

Private Sub Label44_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label44, Source, X, Y)
End Sub

Private Sub Label44_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label44, Button, Shift, X, Y)
End Sub

Private Sub Label49_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label49, Source, X, Y)
End Sub

Private Sub Label49_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label49, Button, Shift, X, Y)
End Sub

Private Sub Label50_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label50, Source, X, Y)
End Sub

Private Sub Label50_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label50, Button, Shift, X, Y)
End Sub

Private Sub Label51_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label51, Source, X, Y)
End Sub

Private Sub Label51_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label51, Button, Shift, X, Y)
End Sub

Private Sub Label39_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label39, Source, X, Y)
End Sub

Private Sub Label39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label39, Button, Shift, X, Y)
End Sub

Private Sub Label41_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label41, Source, X, Y)
End Sub

Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label41, Button, Shift, X, Y)
End Sub

Private Sub Label45_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label45, Source, X, Y)
End Sub

Private Sub Label45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label45, Button, Shift, X, Y)
End Sub

Private Sub Label46_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label46, Source, X, Y)
End Sub

Private Sub Label46_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label46, Button, Shift, X, Y)
End Sub

Private Sub Label47_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label47, Source, X, Y)
End Sub

Private Sub Label47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label47, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel,Opcao)
End Sub

