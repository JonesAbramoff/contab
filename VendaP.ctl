VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl VendaP 
   ClientHeight    =   10785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14700
   KeyPreview      =   -1  'True
   MaskColor       =   &H00C0C0C0&
   ScaleHeight     =   10785
   ScaleWidth      =   14700
   Begin VB.Timer TimerIntegracao 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   -105
      Top             =   1365
   End
   Begin VB.PictureBox Logo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   10320
      ScaleHeight     =   930
      ScaleWidth      =   4185
      TabIndex        =   112
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox ListCF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   1395
      ItemData        =   "VendaP.ctx":0000
      Left            =   7470
      List            =   "VendaP.ctx":0002
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   3870
      Width           =   7020
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   -120
      Top             =   2655
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm LeitoraCodBarras 
      Left            =   -210
      Top             =   2955
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   945
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   10875
      Begin VB.Label Exibe 
         BackStyle       =   0  'Transparent
         Caption         =   "ABC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   300
         TabIndex        =   85
         Top             =   30
         Width           =   10425
      End
      Begin VB.Label Exibe1 
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   270
         TabIndex        =   84
         Top             =   465
         Width           =   10305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   315
      TabIndex        =   81
      Top             =   915
      Width           =   6735
      Begin VB.CommandButton BotaoLimpar 
         Height          =   300
         Left            =   5850
         Picture         =   "VendaP.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   2370
         Width           =   300
      End
      Begin VB.CheckBox ImprimirNFCE 
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
         Height          =   195
         Left            =   5295
         TabIndex        =   109
         Top             =   2010
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.TextBox Email 
         Height          =   300
         Left            =   1440
         MaxLength       =   60
         TabIndex        =   107
         Top             =   1950
         Width           =   3720
      End
      Begin VB.CommandButton BotaoNFD2 
         Caption         =   "NF Manual"
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
         Left            =   5265
         TabIndex        =   101
         Top             =   780
         Width           =   1320
      End
      Begin VB.CheckBox IEIsento 
         Caption         =   "Isento"
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
         Left            =   4020
         TabIndex        =   6
         Top             =   2025
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox IENaoContrib 
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
         Height          =   285
         Left            =   4905
         TabIndex        =   7
         Top             =   1995
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CommandButton BotaoEndEntrega 
         Caption         =   "Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5265
         TabIndex        =   89
         Top             =   1245
         Width           =   1335
      End
      Begin VB.OptionButton OptionDAV 
         Caption         =   "&DAV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2820
         TabIndex        =   10
         Top             =   2400
         Width           =   795
      End
      Begin VB.TextBox NomeCliente 
         Height          =   315
         Left            =   1440
         MaxLength       =   60
         TabIndex        =   4
         Top             =   1290
         Width           =   3720
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   300
         Left            =   5850
         Picture         =   "VendaP.ctx":0536
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Numeração Automática"
         Top             =   2370
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.OptionButton OptionPreVenda 
         Caption         =   "&Pré Venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1275
         TabIndex        =   9
         Top             =   2400
         Width           =   1485
      End
      Begin VB.OptionButton OptionCF 
         Caption         =   "&Cupom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   2385
         Width           =   1110
      End
      Begin MSMask.MaskEdBox CodVendedor 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   210
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Orcamento 
         Height          =   300
         Left            =   4755
         TabIndex        =   11
         Top             =   2370
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Top             =   855
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   315
         Left            =   1770
         TabIndex        =   5
         Top             =   1950
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin VB.Label LabelEmail 
         Alignment       =   1  'Right Justify
         Caption         =   "E-mail:"
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
         Left            =   720
         TabIndex        =   108
         Top             =   1995
         Width           =   645
      End
      Begin VB.Label LabelIE 
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
         Left            =   75
         TabIndex        =   100
         Top             =   2010
         Width           =   1650
      End
      Begin VB.Label LabelCPF 
         AutoSize        =   -1  'True
         Caption         =   "Cnpj/Cpf:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   88
         Top             =   885
         Width           =   960
      End
      Begin VB.Label LabelNomeCliente 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   675
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   86
         Top             =   1290
         Width           =   660
      End
      Begin VB.Label LabelVendedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2160
         TabIndex        =   82
         Top             =   210
         Width           =   3675
      End
      Begin VB.Label LabelCodVendedor 
         AutoSize        =   -1  'True
         Caption         =   "&Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LabelOrcamento 
         AutoSize        =   -1  'True
         Caption         =   "&Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3795
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   2385
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   80
      Top             =   5040
      Width           =   6825
      Begin VB.CommandButton BotaoPagamento 
         Caption         =   "(F11)    Pagamento"
         Height          =   375
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   1920
      End
      Begin VB.CommandButton BotaoCancelaCupom 
         Caption         =   "(Esc)   Cancela Cupom"
         Height          =   360
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1035
         Width           =   1920
      End
      Begin VB.CommandButton BotaoCancelaItemAtual 
         Caption         =   "(F4)   Cancela Item Atual"
         Height          =   360
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   615
         Width           =   1920
      End
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "(F9)   Produtos/Preço"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   1920
      End
      Begin VB.CommandButton BotaoCancelaItem 
         Caption         =   "(F6)    Cancela Item"
         Height          =   360
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   615
         Width           =   1920
      End
      Begin VB.CommandButton BotaoPreco 
         Caption         =   "(F5)    Preço"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1920
      End
      Begin VB.CommandButton BotaoAbrirGaveta 
         Caption         =   "(F10)   Abrir Gaveta"
         Height          =   375
         Left            =   4455
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         Width           =   1920
      End
      Begin VB.CommandButton BotaoSuspender 
         Caption         =   "(F7)  Suspender"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   1035
         Width           =   1920
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   375
         Left            =   4440
         Picture         =   "VendaP.ctx":0620
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   1035
         Width           =   1920
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11400
      Top             =   600
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   4155
      Left            =   120
      TabIndex        =   26
      Top             =   6540
      Width           =   14415
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   12
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   13
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   8
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   9
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   10
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   11
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1900
      End
      Begin VB.CommandButton Menu 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   4
         Left            =   8280
         TabIndex        =   106
         Top             =   255
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.CommandButton Menu 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   3
         Left            =   6243
         TabIndex        =   105
         Top             =   255
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.CommandButton Menu 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   2
         Left            =   4207
         TabIndex        =   104
         Top             =   255
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.CommandButton Menu 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   2171
         TabIndex        =   103
         Top             =   255
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.CommandButton Menu 
         Appearance      =   0  'Flat
         Caption         =   "Principal"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   102
         ToolTipText     =   "Principal"
         Top             =   255
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   28
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   3360
         UseMaskColor    =   -1  'True
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   29
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   30
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   31
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   32
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   33
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   25
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   24
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   23
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   22
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   21
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   27
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   26
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   17
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   16
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   15
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   14
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   20
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   19
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   18
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   6
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   5
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   4
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   34
         Left            =   12360
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   3
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   2
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   1
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   1900
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   62
         Left            =   3435
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   61
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   60
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   59
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   58
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   57
         Left            =   7215
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   56
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   55
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   54
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   -750
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   53
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   52
         Left            =   3435
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   51
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   50
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   49
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   48
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   47
         Left            =   7215
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   46
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   45
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   -1335
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   44
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   43
         Left            =   7215
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   42
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   41
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   40
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   39
         Left            =   3435
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   38
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   37
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   -1920
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   36
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   -1920
         UseMaskColor    =   -1  'True
         Width           =   950
      End
   End
   Begin MSMask.MaskEdBox ProdutoNomeRed 
      Height          =   480
      Left            =   -19275
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1635
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   585
      Left            =   2400
      TabIndex        =   1
      Top             =   3795
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1032
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"VendaP.ctx":071B
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2220
      Left            =   7455
      TabIndex        =   97
      Top             =   3900
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   3916
      _Version        =   393216
      Rows            =   1000
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   12648447
      GridColor       =   12648447
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5520
      Left            =   7110
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   1005
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   9737
      _Version        =   393216
      Rows            =   20
      Cols            =   1
      FixedRows       =   19
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   12648447
      GridColor       =   12648447
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   5580
      Left            =   7440
      ScaleHeight     =   5520
      ScaleWidth      =   4320
      TabIndex        =   98
      Top             =   1320
      Visible         =   0   'False
      Width           =   4380
   End
   Begin MSMask.MaskEdBox DescontoValor 
      Height          =   585
      Left            =   5925
      TabIndex        =   110
      Top             =   3780
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1032
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desconto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4110
      TabIndex        =   111
      Top             =   3825
      Width           =   1770
   End
   Begin VB.Image Figura 
      BorderStyle     =   1  'Fixed Single
      Height          =   765
      Left            =   12840
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   6780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Quantidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   3855
      Width           =   2145
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   765
      TabIndex        =   79
      Top             =   4500
      Width           =   1560
   End
   Begin VB.Label PrecoUnitario 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7230
      TabIndex        =   78
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label PrecoItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9510
      TabIndex        =   77
      Top             =   2760
      Width           =   1920
   End
   Begin VB.Label Subtotal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2400
      TabIndex        =   76
      Top             =   4455
      Width           =   4665
   End
End
Attribute VB_Name = "VendaP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Global
Dim giItem As Integer
Public gobjVenda As ClassVenda
Dim iAlterado As Integer
Dim gsMinutoAnt As String
Dim gcolTecladoProdutoItens As Collection
Dim gsNomeOperador As String
Dim giNaoTrataCF As Integer
Dim giLarguraOrig As Integer
Dim giAlturaOrig As Integer
Dim giLarguraListCF As Integer
Dim giAlturaListCF As Integer
Dim giAlturaFigura As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim giMaxCaracteres As Integer
Dim giLarguraGrid As Integer
Dim giAlturaGrid As Integer
Dim giUltimaLinhaGrid As Integer
Dim giLinhasVisiveisGrid As Integer
Const GRID_COL_ITEM = 0
'Const GRID_COL_CODIGO = 1
Const GRID_COL_DESCRICAO = 1
Const GRID_COL_QUANTIDADE = 2
Const GRID_COL_X = 3
Const GRID_COL_VALOR_UNITARIO = 4
'Const GRID_COL_ST = 6
Const GRID_COL_VALOR_TOTAL = 5

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoOrcamento As AdmEvento
Attribute objEventoOrcamento.VB_VarHelpID = -1

Dim sIEAnt As String

'Origens das chamadas de cancelamento de nfce
Private Const BOTAO_CANCELA_CUPOM_CFE = 1
Private Const BOTAO_LIMPAR_CUPOM_CFE = 2
Private Const BOTAO_FECHAR_CUPOM_CFE = 3 'queryunload
Private Const BOTAO_PAGTO_CANCELA_CUPOM_CFE = 4
Private Const BOTAO_DAV_CANCELA_CUPOM_CFE = 5
Private Const BOTAO_PREVENDA_CANCELA_CUPOM_CFE = 6

Private gbCarregandoTela As Boolean

Private Sub ResizePicture(pBox As PictureBox, pPic As Picture)
'http://www.vbforums.com/showthread.php?428541-Run-Time-Image-Resizing&p=2625056#post2625056

Dim lWidth      As Single, lHeight    As Single
Dim lnewWidth   As Single, lnewHeight As Single, DifLeft As Single
 
    'Clear the Picture in the PictureBox
    pBox.Picture = Nothing
    
    'Clear the Image  in the Picturebox
    pBox.Cls
    
    'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
    lWidth = pBox.ScaleX(pPic.Width, vbHimetric, pBox.ScaleMode)
    lHeight = pBox.ScaleY(pPic.Height, vbHimetric, pBox.ScaleMode)
    
    lnewHeight = pBox.ScaleHeight                   'new Height = PB Height
    lnewWidth = lWidth * (lnewHeight / lHeight)  'Risize Width keeping proportions
    
    DifLeft = pBox.Width - lnewWidth
    pBox.Width = lnewWidth
    pBox.left = pBox.left + DifLeft
    
    pBox.PaintPicture pPic, 0, 0, lnewWidth, lnewHeight
                            
    'Update the Picture with the new image if you need it
    Set pBox.Picture = pBox.Image
    
End Sub

Public Sub Form_Load()
        
Dim lErro As Long
Dim objTeclado As ClassTecladoProduto
Dim objOperador As New ClassOperador
Dim bAchou As Boolean
Dim objTela As Object
Dim sPerfil As String
Dim sRetorno As String
Dim lTamanho As Long
Dim iTelaVendaMaximizada As Integer
Dim lPic As Picture
 
On Error GoTo Erro_Form_Load
   
    gbCarregandoTela = True
    
    If Dir(App.Path & "\LogoLoja.jpg") <> "" Then
        
        
'        Logo.Picture = LoadPicture(App.Path & "\LogoLoja.jpg")
        
        Set lPic = LoadPicture(App.Path & "\LogoLoja.jpg")
        Call ResizePicture(Logo, lPic)
 
        Logo.Visible = True
        
    End If
    
    If giCodModeloECF = IMPRESSORA_SAT_2_5_15 Then
        OptionCF.Caption = "SAT"
    End If
    
    If giCodModeloECF = IMPRESSORA_NFCE Then
        LabelEmail.Visible = True
        Email.Visible = True
        OptionCF.Caption = "NFCe"
    Else
        LabelEmail.Visible = False
        Email.Visible = False
    End If
    
    If giCodModeloECF <> IMPRESSORA_NFE Then
    
        LabelIE.Visible = False
        InscricaoEstadual.Visible = False
        IEIsento.Visible = False
        IENaoContrib.Visible = False
        
        Call Afrac_UF_ObtemPerfil(sPerfil)
        If sPerfil = "Y" Or sPerfil = "Z" Or AFRAC_ImpressoraCFe(giCodModeloECF) Then BotaoNFD2.Visible = False
        
    Else
    
        BotaoNFD2.Visible = False
        BotaoEndEntrega.Caption = "Endereço"
        OptionCF.Caption = "NF"
    
    End If
    
   
    giLarguraOrig = Me.Width
    giAlturaOrig = Me.Height
    
    giLarguraListCF = ListCF.Width
    giAlturaListCF = ListCF.Height
    
    giLarguraGrid = Grid.Width
    giAlturaGrid = Grid.Height
    
    giAlturaFigura = Figura.Height

    gobjLojaECF.dtTime = CDate(Now)
        
'    Apresentacao.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
    If giCodModeloECF = IMPRESSORA_NFCE And gobjNFeInfo.iEmContingencia <> 0 Then
        Exibe.Caption = "PRÓXIMO CLIENTE...(EM CONTINGENCIA)"
    Else
        Exibe.Caption = "PRÓXIMO CLIENTE..."
    End If
    Exibe1.Caption = ""
    
    'fim nova tela

    giItem = 0
    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
    
    If giStatusCaixa = STATUS_CAIXA_FECHADO And giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then gError 133828
    
    'Se a Sessão Estiver Fechada então gera Erro
    If giStatusSessao = SESSAO_ENCERRADA Then gError 99868

    'Se Sessão estiver Suspensa
    If giStatusSessao = SESSAO_SUSPENSA Then gError 99869
     
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210886
     
     
     
'    Call Inicializa_ColTecladoProdutoItens
    
    bAchou = False
    
    For Each objTeclado In gcolTeclados
        'Verifica qual o teclado padrão para o Teclado utilizado
        If objTeclado.iPadrao = TECLADO_PADRAO And objTeclado.iTeclado = giTeclado Then
            'Joga o teclado na tela
            Call Traz_ColTecladoProdutoItens(objTeclado.colTecladoProdutoItem)
            Set gcolTecladoProdutoItens = objTeclado.colTecladoProdutoItem
            bAchou = True
        End If
    Next
    
    'Se não existe teclado --> erro.
    If Not (bAchou) Then gError 99745
    
    'Inicia o Cupom
'    Call Inicia_Cupom_Tela
        
    Quantidade.Text = 1
        
    PrecoUnitario.Caption = Format(0, "standard")
    Subtotal.Caption = Format(0, "standard")
    PrecoItem.Caption = Format(0, "standard")
    
    Orcamento.Text = ""
'    Orcamento.Enabled = False
'    LabelOrcamento.Enabled = False
    BotaoProxNum.Enabled = False
        
    If giPreVenda = 0 Then
        OptionPreVenda.Enabled = False
    End If
    
    If giDAV = 0 Then
        OptionDAV.Enabled = False
    End If
        
    If giPreVenda = 1 And giUsaImpressoraFiscal = 0 Then
        OptionPreVenda.Value = True
    ElseIf giDAV = 1 Then
        OptionDAV.Value = True
    End If
        
    If OptionDAV.Enabled = False And OptionPreVenda.Enabled = False Then
    
        LabelOrcamento.Enabled = False
        Orcamento.Enabled = False
        BotaoProxNum.Enabled = False
        
    End If
        
    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then
        
        If giCodModeloECF = IMPRESSORA_NFE Then
            OptionCF.Enabled = True
        Else
            OptionCF.Enabled = False
        End If
        If giDAV = 1 Then
            BotaoCancelaCupom.Caption = "(Esc)   Imprime Orçamento"
        ElseIf giPreVenda = 1 Then
            BotaoCancelaCupom.Caption = "(Esc)   Limpa Tela"
        End If

        BotaoAbrirGaveta.Caption = "(F10)   Grava Orçamento"
    Else
        OptionCF.Value = True
    End If
    
    'Seleiona o nome do operador
    For Each objOperador In gcolOperadores
        If objOperador.iCodigo = giCodOperador Then gsNomeOperador = objOperador.sNome
    Next
    
    'informa o seu nome
    lErro = AFRAC_InformarOperador(gsNomeOperador)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informar Operador")
        If lErro <> SUCESSO Then gError 99921
    End If
    Quantidade.SelStart = 0
    Quantidade.SelLength = Len(Quantidade.Text)
    
    If giDinheiroAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        BotaoAbrirGaveta.Enabled = False
    End If
    
    Set objTela = Me
    
    lErro = CF_ECF("Inicializa_LeitoraCodBarras", objTela)
    If lErro <> SUCESSO Then gError 117684
    
    Call Timer1_Timer
     
    Set objEventoProduto = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoOrcamento = New AdmEvento
     
    If giCodModeloECF = IMPRESSORA_NFCE Then
    
        If gobjVenda.objNFeInfo.iNFCeEnviarEmail <> 0 Then
        
            ImprimirNFCE.Visible = True
            
            If gobjVenda.objNFeInfo.iNFCeImprimir <> 0 Then
                ImprimirNFCE.Value = vbChecked
                gobjVenda.objCupomFiscal.bNFCEImprimir = True
            Else
                ImprimirNFCE.Value = vbUnchecked
                gobjVenda.objCupomFiscal.bNFCEImprimir = False
            End If
            
        Else
            
            ImprimirNFCE.Visible = False
            gobjVenda.objCupomFiscal.bNFCEImprimir = True
    
        End If
    
    End If
    
    If giIntegracaoTipo <> INTEGRACAO_TIPO_NENHUMA Then
        TimerIntegracao.Enabled = True
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr
        
        Case 99745
            Call Rotina_ErroECF(vbOKOnly, ERRO_TECLADOPADRAO_NAO_EXISTE, gErr)
            
        Case 99921, 117684, 210886
        
        Case 99868
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_ABERTA_INEXISTENTE, gErr, giCodCaixa)

        Case 99869
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_SUSPENSA, gErr, giCodCaixa)
        
        Case 133828
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_FECHADO, gErr, giCodCaixa)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175696)

    End Select

    Exit Sub

End Sub

Private Function ObtemLarguraColuna(ByVal sTexto As String) As Long

Dim tSizeAux As typeSize

        Call GetTextExtentPoint32(Picture1.hdc, sTexto, Len(sTexto), tSizeAux)
        ObtemLarguraColuna = tSizeAux.cx * Screen.TwipsPerPixelX

End Function

Public Sub Inicia_Cupom_Tela()
'Joga o endereço, Nome da Empresa, CNPJ, IE no Cupom

Dim iIndice As Integer
Dim lErro As Long
Dim lLargura As Long
Dim tSize As typeSize
Dim tSize1 As typeSize

    If Grid.Width < 8000 Then

'    ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, 36, "*", gsNomeEmpresa)
'    ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, 36, " ", gsEndereco)
'    ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, 36, " ", gsEnderecoComplemento)
'    ListCF.AddItem "C.G.C.:" & Formata_Campo(ALINHAMENTO_DIREITA, 36, " ", gsCNPJ)
'    ListCF.AddItem "I.E.:" & Formata_Campo(ALINHAMENTO_DIREITA, 36, " ", gsInscricaoEstadual)
'    ListCF.AddItem TRACO_CAB2
'    ListCF.AddItem ""
'    ListCF.AddItem "  ECF :" & FormataCpoNum(giCodECF, 4) & "    LJ :" & FormataCpoNum(giFilialEmpresa, 4) & "    OP :" & FormataCpoNum(giCodOperador, 4)
'    ListCF.AddItem ""
'    ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, 36, " ", "CUPOM FISCAL")
'    ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 36, " ", "Item   Descrição")
'    ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 36, " ", "     Qtd x Unitário      Valor(" & gobjLojaECF.sSimboloMoeda & ")")
'    ListCF.AddItem TRACO_CAB2
    
    Else
    
        Grid.Row = Grid.FixedRows
        Grid.Col = Grid.FixedCols
        Grid.RowSel = Grid.Rows - 1
        Grid.ColSel = Grid.Cols - 1
        Grid.FillStyle = flexFillRepeat
        Grid.Text = ""
        Grid.FillStyle = flexFillSingle
        
        Grid.Row = 0
        Grid.Col = 0
        
        Grid.ScrollBars = flexScrollBarVertical
    
        Grid.TopRow = 2
        
        Grid.TextMatrix(0, GRID_COL_ITEM) = "Item"
        'Grid.TextMatrix(0, GRID_COL_CODIGO) = "Codigo"
        Grid.TextMatrix(0, GRID_COL_DESCRICAO) = "Codigo - Descrição"
        Grid.TextMatrix(0, GRID_COL_QUANTIDADE) = "Qtd. Un."
        Grid.TextMatrix(0, GRID_COL_X) = "x"
        Grid.TextMatrix(0, GRID_COL_VALOR_UNITARIO) = "Vl Un R$"
        'Grid.TextMatrix(0, GRID_COL_ST) = "ST"
        Grid.TextMatrix(0, GRID_COL_VALOR_TOTAL) = "Vl Tot R$"
        
        Call GetTextExtentPoint32(Picture1.hdc, "00000", 5, tSize)
        
        Call GetTextExtentPoint32(Picture1.hdc, "x", 1, tSize1)
        
        'lLargura = (Grid.Width - (tSize1.cx * Screen.TwipsPerPixelX)) / GRID_COL_VALOR_TOTAL
        
        giMaxCaracteres = ControleMaxTamVisivel(Picture1, "-")
        
        For iIndice = GRID_COL_ITEM To GRID_COL_VALOR_TOTAL
        
            Grid.TextMatrix(1, iIndice) = Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres + 100, "-*", "-")
'            If iIndice <> 4 Then Grid.ColWidth(iIndice) = (Grid.Width - ((GetSystemMetrics(SM_CXVSCROLL) + tSize1.cx) * Screen.TwipsPerPixelX)) / GRID_COL_VALOR_TOTAL
            
            If iIndice <> GRID_COL_DESCRICAO Then
                Grid.ColAlignment(iIndice) = flexAlignCenterCenter
            Else
                Grid.ColAlignment(iIndice) = flexAlignLeftCenter
            End If
            
        Next
        
        Grid.ColWidth(GRID_COL_ITEM) = ObtemLarguraColuna("0000")
        Grid.ColWidth(GRID_COL_QUANTIDADE) = ObtemLarguraColuna("12345 und")
        Grid.ColWidth(GRID_COL_X) = ObtemLarguraColuna("X")
        Grid.ColWidth(GRID_COL_VALOR_UNITARIO) = ObtemLarguraColuna("999.999,99")
        Grid.ColWidth(GRID_COL_VALOR_TOTAL) = ObtemLarguraColuna("999.999,99")
        
        'fica com o espaço que sobrar
        Grid.ColWidth(GRID_COL_DESCRICAO) = Grid.Width - (Grid.ColWidth(GRID_COL_ITEM) + Grid.ColWidth(GRID_COL_QUANTIDADE) + Grid.ColWidth(GRID_COL_X) + Grid.ColWidth(GRID_COL_VALOR_UNITARIO) + Grid.ColWidth(GRID_COL_VALOR_TOTAL) + (GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX))
        
        Grid.ColWidth(GRID_COL_VALOR_TOTAL + 1) = 0
        
        giLinhasVisiveisGrid = Grid.Height / Grid.RowHeight(2)
        
        giUltimaLinhaGrid = 1
    
        Grid.Rows = giLinhasVisiveisGrid + 1
        
    End If
    
    Grid1.ColWidth(0) = Grid1.Width
    Grid1.ColAlignment(0) = flexAlignCenterCenter
        
'        Grid.Row = 0
    
        
'    For iIndice = 0 To 10
'        Grid.Col = iIndice
''        Grid.CellWidth = Grid.Width
'        Grid.CellAlignment = flexAlignCenterCenter
'    Next
    
    Grid1.TextMatrix(0, 0) = Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, "*", gsNomeEmpresa)
    Grid1.TextMatrix(1, 0) = gsEndereco
    Grid1.TextMatrix(2, 0) = gsEnderecoComplemento
    Grid1.Row = 4
    Grid1.CellAlignment = flexAlignLeftCenter
    Grid1.TextMatrix(4, 0) = "CNPJ/CPF: " & gsCNPJ
    Grid1.Row = 5
    Grid1.CellAlignment = flexAlignLeftCenter
    Grid1.TextMatrix(5, 0) = "I.E.: " & gsInscricaoEstadual
    Grid1.Row = 6
    Grid1.CellAlignment = flexAlignCenterCenter
    giMaxCaracteres = ControleMaxTamVisivel(Picture1, "-")
    Grid1.TextMatrix(6, 0) = Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, "-*", "-")
    giMaxCaracteres = ControleMaxTamVisivel(Picture1, " ")
    Grid1.TextMatrix(7, 0) = "  ECF :" & FormataCpoNum(giCodECF, 4) & "           LJ :" & FormataCpoNum(giFilialEmpresa, 4) & "          OP :" & FormataCpoNum(giCodOperador, 4)
    Grid1.TextMatrix(9, 0) = "CUPOM FISCAL"
    
    If Grid1.Width < 8000 Then
        
        Grid1.Row = 11
        Grid1.CellAlignment = flexAlignLeftCenter
        Grid1.TextMatrix(11, 0) = "Item    Codigo         Descrição"
        Grid1.Row = 12
        Grid1.CellAlignment = flexAlignLeftCenter
        Grid1.TextMatrix(12, 0) = "        Qtd  Un. x Unitário      ST            Valor(" & gobjLojaECF.sSimboloMoeda & ")"
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "Item    Codigo         Descrição")
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "        Qtd  Un. x Unitário      ST            Valor(" & gobjLojaECF.sSimboloMoeda & ")")
    
    End If
    
End Sub

Private Sub BotaoEndEntrega_Click()
    Call Chama_TelaECF_Modal("EnderecoEntrega", gobjVenda)
    Email.Text = gobjVenda.objCupomFiscal.sEndEntEmail
End Sub

Private Sub BotaoLimpar_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim lNum As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer

On Error GoTo Erro_Botaolimpar_Click
    
    If gobjVenda.objCupomFiscal.bEditavel = False Then gError 201583
    
    If gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        Call CancelaCupomEletronico(BOTAO_LIMPAR_CUPOM_CFE)
        Exit Sub
    
    End If
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210016
    
    lNum = Retorna_Count_ItensCupom
    
    'Se tiver uma venda acontecendo
    If gobjVenda.objCupomFiscal.lNumero <> 0 And gobjVenda.iTipo = OPTION_CF Then
        'Envia aviso perguntando se deseja cancelar a venda
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_VENDA)

        If vbMsgRes = vbYes Then
            'Se for Necessário a autorização do Gerente para abertura do Caixa
            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
        
                'Chama a Tela de Senha
                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
        
                'Sai de Função se a Tela de Login não Retornar ok
                If giRetornoTela <> vbOK Then gError 102506
                
                'Se Operador for Gerente
                iCodGerente = objOperador.iCodigo
        
            End If

            'Cancelar o Cupom de Venda
            lErro = AFRAC_CancelarCupom(Me, gobjVenda)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
            If lErro <> SUCESSO Then gError 99617
            
            Call Move_Dados_Memoria_1
            
            'Realiza as operações necessárias para gravar
            lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204538
            
            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204539
            
        Else
            Exit Sub
        End If
        
    End If
    
    'Se tiver um orçamento na tela
    If lNum > 0 And (OptionPreVenda.Value = True Or OptionDAV.Value = True) Then
        
        'Envia aviso que o orçamento será cancelado
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_ORCAMENTO)
        
        If vbMsgRes = vbYes Then
            'Se for Necessário a autorização do Gerente para abertura do Caixa
            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
        
                'Chama a Tela de Senha
                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
        
                'Sai de Função se a Tela de Login não Retornar ok
                If giRetornoTela <> vbOK Then gError 102506
                
                'Se Operador for Gerente
                iCodGerente = objOperador.iCodigo
        
            End If
            
        Else
           Exit Sub
        End If
        'If vbMsgRes = vbNo Then Exit Sub
            
    End If
    
    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
    
    Call Limpa_Tela_Venda
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr
                
        Case 201583
            Call Rotina_ErroECF(vbOKOnly, ERRO_EM_PROCESSAMENTO_SEFAZ, gErr)
        
        Case 99617, 102506, 204538, 204539, 210016
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175697)

    End Select
    
    Exit Sub

End Sub

'Private Sub CodCliente_GotFocus()
'    Call MaskEdBox_TrataGotFocus(CodCliente, iAlterado)
'End Sub

Private Sub CodVendedor_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodVendedor, iAlterado)
End Sub



Private Sub DescontoValor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescontoValor_GotFocus()
    
    'Posiciona o cursor na frente do campo
    Call MaskEdBox_TrataGotFocus(DescontoValor)
    
End Sub


Private Sub ImprimirNFCE_Click()
    
    If ImprimirNFCE.Visible Then
        gobjVenda.objCupomFiscal.bNFCEImprimir = IIf(ImprimirNFCE.Value = vbChecked, True, False)
    End If

End Sub

Private Sub LabelCPF_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    Call Chama_TelaECF_Modal("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelNomeCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
        
    objCliente.sNomeReduzido = NomeCliente.Text
        
'    'Chama Tela ClienteLista
'    Call Chama_TelaECF_Modal("ClienteLista", objCliente)
'
'    If giRetornoTela = vbOK Then
'
'        Select Case Len(Trim(objCliente.sCgc))
'
'            Case STRING_CPF 'CPF
'
'                'Formata e coloca na Tela
'                CGC.Format = "000\.000\.000-00; ; ; "
'
'            Case STRING_CGC 'CGC
'
'                'Formata e Coloca na Tela
'                CGC.Format = "00\.000\.000\/0000-00; ; ; "
'
'        End Select
'
'        NomeCliente.Text = objCliente.sNomeReduzido
'        gobjVenda.objCupomFiscal.sCPFCGC = objCliente.sCgc
'        gobjVenda.objCupomFiscal.sNomeCliente = objCliente.sNomeReduzido
'        CGC.Text = objCliente.sCgc
'        gobjVenda.objCupomFiscal.sCPFCGC1 = CGC.FormattedText
'    End If
            
    Call Chama_TelaECF_Modal("ClientesLista", colSelecao, objCliente, objEventoCliente)
            
            
End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
    
    'Chama tela de vendedorLista
    Call Chama_TelaECF_Modal("VendedoresLista", colSelecao, objVendedor, objEventoVendedor)
    
'    If giRetornoTela = vbOK Then
'        CodVendedor.Text = objVendedor.iCodigo
'        Call CodVendedor_Validate(False)
'    End If
    
    Exit Sub

End Sub

Private Sub Menu_Click(Index As Integer)

Dim objTecladoProdutoItens As New ClassTecladoProdutoItem
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim sTecla As String
Dim objTecladoProdutoItensPai As New ClassTecladoProdutoItem

    Call Limpa_Teclado

    'se for um menu preenchido
    If Len(Menu(Index).Tag) > 0 Then

        'coloca o novo menu na tela

        'preenche os botoes de acordo com o que esta armazendo na treeview no nivel selecionado
        For Each objTecladoProdutoItens In gcolTecladoProdutoItens
            If left(objTecladoProdutoItens.sArvoreKey, Len(objTecladoProdutoItens.sArvoreKey) - 2) = Menu(Index).Tag Then
        
                iIndice = CInt(right(objTecladoProdutoItens.sArvoreKey, 2)) - 1
                
                If objTecladoProdutoItens.sTitulo <> "" Then
                    If objTecladoProdutoItens.iTecla <> 0 Then
                        If objTecladoProdutoItens.iTecla > vbKeyF2 Then
                            Call Acha_Tecla(objTecladoProdutoItens.iTecla, sTecla)
                        Else
                            sTecla = Chr(objTecladoProdutoItens.iTecla)
                        End If
                        Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo & "(" & sTecla & ")"
                    Else
                        Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo
                    End If
                Else
                    Produto(iIndice).Caption = ""
                End If
                
                Produto(iIndice).Tag = objTecladoProdutoItens.sArvoreKey
                Produto(iIndice).BackColor = objTecladoProdutoItens.lColor
                Produto(iIndice).ToolTipText = Produto(iIndice).Caption
                    
            End If
        Next
    
    Else
        
        Call Limpa_Teclado
                
        Call Traz_ColTecladoProdutoItens(gcolTecladoProdutoItens)
        
    End If

    'limpa os menus abaixo do selecionado
    For iIndice = Index + 1 To Menu.Count - 1
        Menu(iIndice).Caption = ""
        Menu(iIndice).ToolTipText = Menu(iIndice).Caption
        Menu(iIndice).Tag = ""
        Menu(iIndice).Visible = False
    Next


End Sub

Private Sub Produto_LostFocus(Index As Integer)
'    Quantidade.SelStart = 0
'    Quantidade.SelLength = Len(Quantidade.Text)
'    Quantidade.SetFocus
End Sub


Private Sub Timer1_Timer()
    
Dim vbMsgBox As VbMsgBoxResult
Dim bAchou As Boolean
Dim lErro As Long
Dim dTimerTemp As Double
Dim dtData As Date
Dim dtime As Double
Dim dtTime As Date
Dim lSequencial As Long
Dim sHora As String
Dim dtUltimaReducaoECF As Date
Dim sDataHora As String
Dim dtUltimoMovto As Date


On Error GoTo Erro_Timer1_Timer
    
'    dtime = Timer
'    If dtime > 3600 Then
'        'Coloca a hora atual do Sistema
'        sHora = CStr(dtime / (60 * 60))
'        iPosHora = InStr(1, sHora, ",")
'        If iPosHora > 0 Then sHora = Mid(sHora, 1, iPosHora - 1)
'    Else
'        sHora = 0
'    End If
'
'    If sHora <> 0 Then
'        dTimerTemp = dtime - (CLng(sHora * 3600))
'    Else
'        dTimerTemp = dtime
'    End If
'
'    If dTimerTemp > 60 Then
'        sMinuto = CStr(dtime / 60) - (CInt(sHora * 60))
'        iPosMinuto = InStr(1, sMinuto, ",")
'        If iPosMinuto > 0 Then sMinuto = Mid(sMinuto, 1, iPosMinuto - 1)
'    Else
'        sMinuto = 0
'    End If
'
'    ssegundo = CStr(dtime) - ((CLng(sMinuto * 60)) + (CLng(sHora * 3600)))
'    iPossegundo = InStr(1, ssegundo, ",")
'    If iPossegundo > 0 Then ssegundo = Mid(ssegundo, 1, iPossegundo - 1)
'
''    DataHora.Caption = Format(Date, "dd/mm/yyyy") & "   " & Format(sHora, "00") & ":" & Format(sMinuto, "00") & ":" & Format(ssegundo, "00")
'
'    Me.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", "Venda") & "Filial: " & giFilialEmpresa & "    Caixa: " & giCodCaixa & "    Operador: " & gsNomeOperador & _
'    "    Data: " & Format(Date, "dd/mm/yyyy") & "     Hora: " & Format(sHora, "00") & ":" & Format(sMinuto, "00") & ":" & Format(ssegundo, "00") & "    Empresa: " & Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
    
    If Me.ActiveControl Is Nothing Then Exit Sub
    
    Me.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", "Venda") & "Filial: " & giFilialEmpresa & "    Caixa: " & giCodCaixa & "    Operador: " & gsNomeOperador & _
    "    Data: " & Format(Date, "dd/mm/yyyy") & "     Hora: " & Format(Time, "Hh:Nn:Ss") & "    Empresa: " & Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
    
    'Me.Caption = Me.Caption & " R$ " & CStr(gdSaldoDinheiro)
    
    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
    
        If AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
            If gdtDataAnterior <> Date And gobjVenda.objCupomFiscal.lNumero = 0 Then
                
                'messagem avisando dá mudança de dia
                vbMsgBox = Rotina_AvisoECF(vbOKOnly, AVISO_INICIALIZAR_SISTEMA_AGORA)
                
                lErro = CF_ECF("Caixa_Executa_Fechamento")
                If lErro <> SUCESSO And lErro <> 118005 Then gError 118008
                                                        
                lErro = CF_ECF("Caixa_Executa_Abertura")
                If lErro <> SUCESSO Then gError 118011
                
                lErro = CF_ECF("Verifica_Arquivo_Loja", dtUltimoMovto)
                If lErro <> SUCESSO And lErro <> 53 Then gError ERRO_SEM_MENSAGEM
                
                'Função que Faz a Abertura de Sessão
                lErro = CF_ECF("Sessao_Executa_Abertura")
                If lErro <> SUCESSO Then gError 108013
                
            End If
        
        Else
    
            DoEvents
        
            'le a data/hora do ecf
            lErro = AFRAC_LerInformacaoImpressora("017", sDataHora)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
            If lErro <> SUCESSO Then gError 214513
        
            dtData = StrParaDate(left(sDataHora, 2) & "/" & Mid(sDataHora, 3, 2) & "/" & Mid(sDataHora, 5, 2))
    
            dtTime = StrParaDate(Mid(sDataHora, 7, 2) & ":" & Mid(sDataHora, 9, 2) & ":" & right(sDataHora, 2))
        
            If gdtDataAnterior <> dtData Then
                
                'se a hora >= 2  e tiver ocorrido uma reducao Z no ecf e ainda nao ocorreu no sistema ==> reducao Z no sistema
                If Hour(dtTime) >= 2 Then
                    
    '                lErro = AFRAC_DataReducao(dtUltimaReducaoECF)
    '                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data Ultima Reducao Z")
    '                If lErro <> SUCESSO Then gError 214032
    '
    '                If gdtUltimaReducao < dtUltimaReducaoECF Then
                
                        If gobjVenda.objCupomFiscal.lNumero = 0 Then
                            
                            lErro = CF_ECF("Requisito_XXII")
                            If lErro <> SUCESSO Then gError 210017
                            
                            'messagem avisando dá mudança de dia
                            vbMsgBox = Rotina_AvisoECF(vbOKOnly, AVISO_INICIALIZAR_SISTEMA_AGORA)
                            
                            lErro = CF_ECF("Caixa_Executa_Fechamento")
                            If lErro <> SUCESSO And lErro <> 118005 Then gError 118008
                                                                    
        '                    gdtDataHoje = Date
                                                                    
                            lErro = CF_ECF("Carrega_Arquivo_FonteDados")
                            If lErro <> SUCESSO Then gError 210199
                            
                            lErro = CF_ECF("Requisito_VII_Item8_E2")
                            If lErro <> SUCESSO Then gError 210200
        
                            lErro = CF_ECF("Caixa_Executa_Abertura")
                            If lErro <> SUCESSO Then gError 118011
                            
                            If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
                            
                                lErro = AFRAC_AbrirDia(Date)
                                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abertura do Dia")
                                If lErro <> SUCESSO Then gError 118012
                                
                            End If
                            
                            'Função que Faz a Abertura de Sessão
                            lErro = CF_ECF("Sessao_Executa_Abertura")
                            If lErro <> SUCESSO Then gError 108013
                        End If
                        
    '                End If
                        
                ElseIf Minute(dtTime) = 30 Or Minute(dtTime) = 0 Then
                    If Minute(dtTime) <> gsMinutoAnt Then
                        'messagem avisando dá mudança de dia
                        vbMsgBox = Rotina_AvisoECF(vbOKOnly, AVISO_INICIALIZAR_SISTEMA)
                        gsMinutoAnt = Minute(dtTime)
                    End If
                End If
            End If
    
        End If
    
    End If
    
    Exit Sub
    
Erro_Timer1_Timer:
    
    Select Case gErr
        
        Case 118008, 118011 To 118013, 210017, 210199, 210200, 214032, 57097, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175698)

    End Select
    
    Exit Sub

End Sub

Private Function Retorna_Count_ItensCupom() As Long

Dim objItens As ClassItemCupomFiscal
Dim lNum As Long
    
'    lNum = 0
'
'    For Each objItens In gobjVenda.objCupomFiscal.colItens
'        'se nao for um item cancelado
'        If objItens.iStatus = STATUS_ATIVO Then lNum = lNum + 1
'    Next
'
'    Retorna_Count_ItensCupom = lNum

    Retorna_Count_ItensCupom = gobjVenda.objCupomFiscal.colItens.Count

    
End Function

'Sub Inicializa_ColTecladoProdutoItens()
''Inicializa a col com os valores default
'
'Dim iIndice As Integer
'Dim objTecladoProdutoItens As ClassTecladoProdutoItem
'
'    Set gcolTecladoProdutoItens = New Collection
'
'    For iIndice = 0 To 34
'
'        Set objTecladoProdutoItens = New ClassTecladoProdutoItem
'
'        objTecladoProdutoItens.iTecla = 0
'        objTecladoProdutoItens.lColor = COR_DEFAULT
'        objTecladoProdutoItens.sProduto = ""
'        objTecladoProdutoItens.sTitulo = ""
'        objTecladoProdutoItens.iIndice = iIndice
'
'        gcolTecladoProdutoItens.Add objTecladoProdutoItens
'    Next
'
'End Sub

Sub Traz_ColTecladoProdutoItens(ColTecladoProdutoItens As Collection)
'joga a col passada na tela

Dim objTecladoProdutoItens1 As ClassTecladoProdutoItem
Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim sTecla As String
Dim iIndice As Integer

    'Pra cada item do TecladoProduto eu jogo seu correnspondente na tela
'    For Each objTecladoProdutoItens1 In ColTecladoProdutoItens
'
'        Set objTecladoProdutoItens = New ClassTecladoProdutoItem
'
'        Set objTecladoProdutoItens = gcolTecladoProdutoItens.Item(objTecladoProdutoItens1.iIndice + 1)
'
'        Produto(objTecladoProdutoItens1.iIndice).BackColor = objTecladoProdutoItens1.lColor
'        If objTecladoProdutoItens1.sTitulo <> "" Then
'            If objTecladoProdutoItens1.iTecla <> 0 Then
'                If objTecladoProdutoItens1.iTecla > vbKeyF2 Then
'                    Call Acha_Tecla(objTecladoProdutoItens1.iTecla, sTecla)
'                Else
'                    sTecla = Chr(objTecladoProdutoItens1.iTecla)
'                End If
'                Produto(objTecladoProdutoItens1.iIndice).Caption = objTecladoProdutoItens1.sTitulo & "(" & sTecla & ")"
'            Else
'                Produto(objTecladoProdutoItens1.iIndice).Caption = objTecladoProdutoItens1.sTitulo
'            End If
'        Else
'            Produto(objTecladoProdutoItens1.iIndice).Caption = ""
'        End If
'        objTecladoProdutoItens.sProduto = objTecladoProdutoItens1.sProduto
'        objTecladoProdutoItens.iTecla = objTecladoProdutoItens1.iTecla
'        objTecladoProdutoItens.lColor = objTecladoProdutoItens1.lColor
'        objTecladoProdutoItens.sTitulo = objTecladoProdutoItens1.sTitulo
'        objTecladoProdutoItens.sReferencia = objTecladoProdutoItens1.sReferencia
'        objTecladoProdutoItens.sCodigoBarras = objTecladoProdutoItens1.sCodigoBarras
'
'    Next
    
    
    'preenche os botoes de acordo com o que esta armazendo na treeview no nivel selecionado
    For Each objTecladoProdutoItens In ColTecladoProdutoItens
    
        If Len(objTecladoProdutoItens.sArvoreKey) = 3 Then
        
            iIndice = CInt(right(objTecladoProdutoItens.sArvoreKey, 2)) - 1
            
            If objTecladoProdutoItens.sTitulo <> "" Then
                If objTecladoProdutoItens.iTecla <> 0 Then
                    If objTecladoProdutoItens.iTecla > vbKeyF2 Then
                        Call Acha_Tecla(objTecladoProdutoItens.iTecla, sTecla)
                    Else
                        sTecla = Chr(objTecladoProdutoItens.iTecla)
                    End If
                    Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo & "(" & sTecla & ")"
                Else
                    Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo
                End If
            Else
                Produto(iIndice).Caption = ""
            End If
            
            Produto(iIndice).Tag = objTecladoProdutoItens.sArvoreKey
            Produto(iIndice).BackColor = objTecladoProdutoItens.lColor
            Produto(iIndice).ToolTipText = Produto(iIndice).Caption
            
        End If
        
    Next
    
    
End Sub

Public Sub Acha_Produto_Tecla(iTecla As Integer, bAchou As Boolean)

Dim objTecladoProduto As New ClassTecladoProdutoItem
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Acha_Produto_Tela

    iIndice = 0
    bAchou = False
    
    For Each objTecladoProduto In gcolTecladoProdutoItens
        If objTecladoProduto.iTecla = iTecla Then
            
            'Se a quantidade e o produto estão prenchidos
            If Len(Trim(Quantidade.Text)) > 0 Then
                
                lErro = Valor_Positivo_Critica(Quantidade.Text)
                If lErro <> SUCESSO Then gError 99888
                If Len(Trim(Fix(Quantidade.Text))) > 4 Then gError 112423
                Quantidade.Text = Format(Quantidade.Text, "0.000")
                If right(Quantidade.Text, 4) = ",000" Then Quantidade.Text = Format(Quantidade.Text, "#,#")
            End If
            
            Call Produto_Click(iIndice)
            bAchou = True
            Exit For
        End If
        iIndice = iIndice + 1
    Next
    
    Exit Sub

Erro_Acha_Produto_Tela:

    Select Case gErr
            
        Case 99888
        
        Case 112423
            Call Rotina_ErroECF(vbOKOnly, ERRO_QUANTIDADE_INVALIDA, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175699)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Acha_Tecla(iTecla As Integer, sTecla As String)

    Select Case iTecla
    
        Case vbKeyF2
            sTecla = "F2"
        Case vbKeyF3
            sTecla = "F3"
        Case vbKeyF4
            sTecla = "F4"
        Case vbKeyF5
            sTecla = "F5"
        Case vbKeyF6
            sTecla = "F6"
        Case vbKeyF7
            sTecla = "F7"
        Case vbKeyF8
            sTecla = "F8"
        Case vbKeyF9
            sTecla = "F9"
        Case vbKeyF10
            sTecla = "F10"
        Case vbKeyF11
            sTecla = "F11"
        Case vbKeyF12
            sTecla = "F12"
        Case vbKeyF13
            sTecla = "F13"
        Case vbKeyF14
            sTecla = "F14"
        Case vbKeyF15
            sTecla = "F15"
        Case vbKeyF16
            sTecla = "F16"
           
    End Select
    
End Sub


Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
        
    
End Sub

Private Sub LabelCodVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
    
    'Chama tela de vendedorLista
    Call Chama_TelaECF_Modal("VendedoresLista", colSelecao, objVendedor, objEventoVendedor)
    
'    If giRetornoTela = vbOK Then
'        CodVendedor.Text = objVendedor.iCodigo
'        Call CodVendedor_Validate(False)
'    End If
    
    Exit Sub

End Sub

Private Sub LabelOrcamento_Click()

Dim objOrcamento As New ClassOrcamentoLoja
Dim objVenda As New ClassVenda
Dim colOrcamento As New Collection
Dim iAchou As Integer
Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim iIndice As Integer
Dim objProduto As ClassProduto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelOrcamento_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210018

    If giDAV = 1 Then
        sSelecao = "NumeroDAV <> '0000000000'"
    ElseIf giPreVenda = 1 Then
        sSelecao = "NumeroDAV = '0000000000'"
    End If

    lErro = CF_ECF("ODBC_Testa_Acesso_Orc")
    If lErro = 216221 Then gError 210018
    
    'Chama tela de OrçamentoLista
    Call Chama_TelaECF_Modal("OrcamentoLista", colSelecao, objOrcamento, objEventoOrcamento, sSelecao)
    
'    If giRetornoTela = vbOK Then
'
'
'        objVenda.objCupomFiscal.lNumOrcamento = objOrcamento.lNumOrcamento
'
'        'Função Que le os orcamentos
'        lErro = CF_ECF("OrcamentoECF_Le", objVenda)
'        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 105868
'
'        'orcamento nao cadastrado
'        If lErro = 204690 Then gError 105869
'
'        If lErro = 210447 Then gError 210448
'
'
'        For Each objItens In objVenda.objCupomFiscal.colItens
'
'            Set objProduto = New ClassProduto
'
'            lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
'            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214852
'
'            If lErro = SUCESSO Then
'                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
'            End If
'
'
'        Next
'
'
'        'Traz ele para a tela
'        Set gobjVenda = New ClassVenda
'        Call Copia_Venda(gobjVenda, objVenda)
'        Call Traz_Orcamento
'
'        'se o cupom fiscal estiver ligado, cham OptionCF_Click para transformar o orcamento em cupom
'        If OptionCF.Value Then Call OptionCF_Click
'
'    End If
    
    Exit Sub

Erro_LabelOrcamento_Click:

    Select Case gErr

        Case 210018
'
'        Case 105869
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_CADASTRADO1, gErr, objVenda.objCupomFiscal.lNumOrcamento)
'
'        Case 210448
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda.objCupomFiscal.lNumOrcamento)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175700)

    End Select

    Exit Sub

End Sub

'Private Sub codCliente_Change()
''Determina se Houve Mudança
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'
'Private Sub CodCliente_Validate(Cancel As Boolean)
'
'On Error GoTo Erro_Cliente_Validate
'
'Dim objCliente As ClassCliente
'Dim lIndice As Long
'Dim lErro As Long
'
'On Error GoTo Erro_Cliente_Validate
'
'    If Len(Trim(CodCliente.ClipText)) > 0 Then
'
'        If IsNumeric(CodCliente.ClipText) Then
'
'            Set objCliente = gobjClienteCPF.Busca(CodCliente.ClipText, lIndice)
'
'            If objCliente Is Nothing Then gError 204307
'
'            Select Case Len(Trim(CodCliente.ClipText))
'
'                Case STRING_CPF 'CPF
'
'                    'Critica Cpf
'                    lErro = Cpf_Critica(CodCliente.ClipText)
'                    If lErro <> SUCESSO Then gError 204308
'
'                    'Formata e coloca na Tela
'                    CodCliente.Format = "000\.000\.000-00; ; ; "
'
'                Case STRING_CGC 'CGC
'
'                    'Critica CGC
'                    lErro = Cgc_Critica(CodCliente.ClipText)
'                    If lErro <> SUCESSO Then gError 204309
'
'                    'Formata e Coloca na Tela
'                    CodCliente.Format = "00\.000\.000\/0000-00; ; ; "
'
'                Case Else
'                    gError 204310
'
'            End Select
'
'        Else
'
'            Set objCliente = gobjClienteNome.Busca(CodCliente.ClipText, lIndice)
'
'            If objCliente Is Nothing Then gError 204311
'
'        End If
'
'        'joga o cliente no gobjvenda
'        gobjVenda.objCupomFiscal.sCPFCGC = objCliente.sCgc
'
'        CodCliente.Text = objCliente.sCgc
'        gobjVenda.objCupomFiscal.sCPFCGC1 = CodCliente.FormattedText
'
'        If Len(Trim(CGC.ClipText)) = 0 Then
'            CGC.Text = objCliente.sCgc
'        End If
'
'        If Len(Trim(NomeCliente.Text)) = 0 Then
'            NomeCliente.Text = objCliente.sNomeReduzido
'            gobjVenda.objCupomFiscal.sNomeCliente = objCliente.sNomeReduzido
'        End If
'
'
'    End If
'
'    Exit Sub
'
'Erro_Cliente_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 204307
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CPFCNPJ_NAO_CADASTRADO, gErr, CodCliente.Text)
'
'        Case 204308, 204309
'
'        Case 204310
'            Call Rotina_Erro(vbOKOnly, ERRO_TAMANHO_CGC_CPF1, gErr)
'
'        Case 204311
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_CADASTRADO3, gErr, CodCliente.Text)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175701)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub LabelCliente_Click()
'
'Dim objCliente As New ClassCliente
'Dim colSelecao As Collection
'
''     'Chama Tela ClienteLista
''    Call Chama_TelaECF_Modal("ClienteLista", objCliente)
''
''    If giRetornoTela = vbOK Then
''        CodCliente.Text = objCliente.sCgc
''    End If
'
'
'    objCliente.sNomeReduzido = CodCliente.Text
'
'    'Chama Tela ClienteLista
'    Call Chama_TelaECF_Modal("ClienteLista", objCliente)
'
'    If giRetornoTela = vbOK Then
'
'        Select Case Len(Trim(objCliente.sCgc))
'
'            Case STRING_CPF 'CPF
'
'                'Formata e coloca na Tela
'                CodCliente.Format = "000\.000\.000-00; ; ; "
'
'            Case STRING_CGC 'CGC
'
'                'Formata e Coloca na Tela
'                CodCliente.Format = "00\.000\.000\/0000-00; ; ; "
'
'        End Select
'
'        CodCliente.Text = objCliente.sCgc
'        NomeCliente.Text = objCliente.sNomeReduzido
'        gobjVenda.objCupomFiscal.sCPFCGC = objCliente.sCgc
'        gobjVenda.objCupomFiscal.sNomeCliente = objCliente.sNomeReduzido
'        CGC.Text = objCliente.sCgc
'    End If
'
'
'End Sub

Private Sub BotaoProdutos_LostFocus()
    Quantidade.SelStart = 0
    Quantidade.SelLength = Len(Quantidade.Text)
    Quantidade.SetFocus
End Sub

Public Sub BotaoProdutos_Click()
'Chama o browser do ProdutoLojaLista
'So traz produtos onde codigo de barras ou referencia está preenchida

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoProdutos_Click
    
'    objProduto.sNomeReduzido = ProdutoNomeRed.Text
'
'    'Chama tela de ProdutosLista
'    Call Chama_TelaECF_Modal("ProdutosLista", objProduto)
'
'    If giRetornoTela = vbOK Then
'        If objProduto.sReferencia <> "" Then
'            ProdutoNomeRed.Text = objProduto.sReferencia
'        Else
'            ProdutoNomeRed.Text = objProduto.sCodigoBarras
'        End If
'        Call ProdutoNomeRed_Validate(False)
'    End If
    
    
    objProduto.sNomeReduzido = ProdutoNomeRed.Text

    Call Chama_TelaECF_Modal("ProdutosLista", colSelecao, objProduto, objEventoProduto)
    
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175702)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
    
Dim lErro As Long
Dim lNumero As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210019

    
    If gobjVenda.iTipo = OPTION_CF Then Exit Sub
    
    'Função que obtém o próximo número
    lErro = CF_ECF("Venda_Obtem_Num_Automatico", lNumero)
    If lErro <> SUCESSO Then gError 99907

    'joga na tela
    Orcamento.Text = lNumero
    gobjVenda.objCupomFiscal.lNumOrcamento = lNumero
    
    If gobjVenda.objCupomFiscal.IdExterno <> "" Then
        gobjVenda.objCupomFiscal.IdExterno = ""
        gobjVenda.objCupomFiscal.lNumeroDAV = 0
        gobjVenda.objCupomFiscal.NumCaixa = 0
        gobjVenda.objCupomFiscal.StatusExterno = 0
    End If
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 99907, 210019
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175703)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCancelaItemAtual_Click()
    
Dim iItem As Integer
Dim lErro As Long
Dim iIndice As Integer
Dim lNum As Long
Dim objItens As New ClassItemCupomFiscal
Dim objVenda As ClassVenda
Dim objVendaParam As New ClassVenda
Dim iAchou As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_BotaoCancelaItemAtual_Click
           
    If gobjVenda.objCupomFiscal.bEditavel = False Then gError 201583
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210020
           
    'Retorna o count de gobjvenda.objcupomfiscal.colitens
    lNum = Retorna_Count_ItensCupom
    
    If lNum > 0 Then
        
        Set objVenda = gobjVenda
        
        'If (objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.lNumOrcamento = 0 Then gError 210501
    
        objVendaParam.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
        
        If objVendaParam.objCupomFiscal.lNumOrcamento <> 0 Then
        
            'le o orcamento em questao
            lErro = CF_ECF("OrcamentoECF_Le", objVendaParam)
            If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210502
        
        End If
        
        'se o orcamento é um DAV e ja foi impresso ==> nao pode alterar
        If objVendaParam.iTipo = OPTION_DAV And objVendaParam.objCupomFiscal.iDAVImpresso <> 0 Then gError 210503
        
        If gobjVenda.iTipo = OPTION_CF And gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then gError 214305
        
        'Se for necessária a autorização do Gerente
        If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
    
            'Chama a Tela de Senha
            Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
    
            'Sai de Função se a Tela de Login não Retornar ok
            If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
            
        End If
                    
        iAchou = 0
        
        For iIndice = gobjVenda.objCupomFiscal.colItens.Count To 1 Step -1
            Set objItens = gobjVenda.objCupomFiscal.colItens.Item(iIndice)
            If objItens.iStatus = STATUS_ATIVO Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 1 Then
        
        
            If gobjVenda.iTipo = OPTION_CF Then
                'cancelar o Item anterior
                lErro = AFRAC_CancelarItem(CInt(objItens.iItem))
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancela Item")
                If lErro <> SUCESSO Then gError 99732
                
                
    '            Call CF_ECF("Requisito_XXII_AtualizaGT")
                
                If AFRAC_ImpressoraCFe(giCodModeloECF) Then Call CF_ECF("Grava_Venda_Atual", gobjVenda)
                
            End If
            
    '        iItem = ListCF.ItemData(ListCF.ListCount - 1)
    '
    '        If iItem = 0 Then Exit Sub
            
            Exibe.Caption = "CANCELADO ITEM " & objItens.iItem
            Exibe1.Caption = ""
            lErro = Exclui_Item_ColItens(objItens.iItem)
            If lErro <> SUCESSO Then gError 214144
            
        
        Else
            gError 210519
        
        End If
        
    Else
        'Senão erro-->deve existir um item
        gError 99925
    End If
    
    
    Exit Sub

Erro_BotaoCancelaItemAtual_Click:

    Select Case gErr
                
        Case 201583
            Call Rotina_ErroECF(vbOKOnly, ERRO_EM_PROCESSAMENTO_SEFAZ, gErr)
        
        Case 99732, 210020, 210502, 214144, ERRO_SEM_MENSAGEM
                            
        Case 99925
            Call Rotina_ErroECF(vbOKOnly, ERRO_NAO_EXISTE_ITEM, gErr, Error$)
            
        Case 210501
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr)
        
        Case 210503
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
            
        Case 210519
            Call Rotina_ErroECF(vbOKOnly, ERRO_ITEM_NAO_ENCONTRADO_CANCELAR, gErr)
            
        Case 214305
            Call Rotina_ErroECF(vbOKOnly, ERRO_ITEM_ORCAMENTO_CUPOM_NAO_PODE_CANCELAR, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175704)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCancelaItem_Click()
    
Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim lNum As Long
Dim objVenda As ClassVenda
Dim objVendaParam As New ClassVenda
Dim objOperador As New ClassOperador

On Error GoTo Erro_BotaoCancelaItem_Click
    
    If gobjVenda.objCupomFiscal.bEditavel = False Then gError 201583
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210021

    
    'Retorna o count de gobjvenda.objcupomfiscal.colitens
    lNum = Retorna_Count_ItensCupom
    
    If lNum > 0 Then
    
        Set objVenda = gobjVenda
        
        'If (objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.lNumOrcamento = 0 Then gError 210495
    
        objVendaParam.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
        
        If objVendaParam.objCupomFiscal.lNumOrcamento <> 0 Then
        
            'le o orcamento em questao
            lErro = CF_ECF("OrcamentoECF_Le", objVendaParam)
            If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210496
        
        End If
        
        'se o orcamento é um DAV e ja foi impresso ==> nao pode alterar
        If objVendaParam.iTipo = OPTION_DAV And objVendaParam.objCupomFiscal.iDAVImpresso <> 0 Then gError 210497
        
        If gobjVenda.iTipo = OPTION_CF And gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then gError 214304
        
        
        If Grid.Width < 8000 Then
        
            'Se tiver algum item selecionado e ele não for do cabeçalho-->continua
'            If ListCF.ListIndex > 13 Then
                objVenda.objCupomFiscal.iItem = ListCF.ItemData(ListCF.ListIndex)
'            Else
'                objVenda.objCupomFiscal.iItem = 0
'            End If
        
        Else
            If Grid.RowSel > 1 And Grid.RowSel <= giUltimaLinhaGrid Then
                objVenda.objCupomFiscal.iItem = StrParaInt(Grid.TextMatrix(Grid.RowSel, GRID_COL_ITEM))
            Else
                objVenda.objCupomFiscal.iItem = 0
            End If
            
        End If
        
        'Se for necessária a autorização do Gerente
        If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
    
            'Chama a Tela de Senha
            Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
    
            'Sai de Função se a Tela de Login não Retornar ok
            If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call Chama_TelaECF_Modal("CancelaItem", gobjVenda)
        
        If giRetornoTela = vbOK Then
            If gobjVenda.iTipo = OPTION_CF Then
                'cancelar o item de Venda
                lErro = AFRAC_CancelarItem(gobjVenda.objCupomFiscal.iItem)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Item")
                If lErro <> SUCESSO Then gError 99733
                
'                Call CF_ECF("Requisito_XXII_AtualizaGT")
                
                If AFRAC_ImpressoraCFe(giCodModeloECF) Then Call CF_ECF("Grava_Venda_Atual", gobjVenda)
                
            End If
            
            lErro = Exclui_Item_ColItens(gobjVenda.objCupomFiscal.iItem)
            If lErro <> SUCESSO Then gError 214145
            
            
            Exibe.Caption = "CANCELADO ITEM " & gobjVenda.objCupomFiscal.iItem
            Exibe1.Caption = ""
         End If
         
    Else
        'Senão erro-->deve existir um item
        gError 99924
    End If
    
    Exit Sub

Erro_BotaoCancelaItem_Click:

    Select Case gErr
                
        Case 201583
            Call Rotina_ErroECF(vbOKOnly, ERRO_EM_PROCESSAMENTO_SEFAZ, gErr)
        
        Case 99733, 210021, 210496, 214145, ERRO_SEM_MENSAGEM
                                    
        Case 99896
            Call Rotina_ErroECF(vbOKOnly, ITEM_CUPOM_NAO_SELECIONADO, gErr, Error$)
                                    
        Case 99924
            Call Rotina_ErroECF(vbOKOnly, ERRO_NAO_EXISTE_ITEM, gErr, Error$)
            
        Case 210495
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr)
        
        Case 210497
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
            
        Case 214304
            Call Rotina_ErroECF(vbOKOnly, ERRO_ITEM_ORCAMENTO_CUPOM_NAO_PODE_CANCELAR, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175705)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCancelaCupom_Click()
    
Dim lErro As Long
Dim iCodigo As Integer
Dim lNum As Long
Dim iIndice As Integer
Dim objItens As New ClassItemCupomFiscal
Dim objAliquota As New ClassAliquotaICMS
Dim sRetorno As String
Dim vbMsgRes As VbMsgBoxResult
Dim lSequencial As Long
Dim sLog As String
Dim colRegistro As New Collection
Dim objCliente As ClassCliente
Dim sCPF As String
Dim lNumero As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer
Dim lRetorno As Long
Dim iFlag As Integer
Dim objTela As Object
Dim dtDataFinal As Date
Dim objVenda As New ClassVenda
Dim objVendaAux As ClassVenda
Dim lTamanho As Long
Dim objFormMsg As Object
Dim objMovCaixa As ClassMovimentoCaixa
Dim objMovCaixa1 As ClassMovimentoCaixa
Dim iCuponsVinculados As Integer
Dim colMeiosPag As New Collection

On Error GoTo Erro_BotaoCancelaCupom_Click

    Set objTela = Me

    If gobjVenda.objCupomFiscal.bEditavel = False Then gError 201583
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210022

    'se é uma PRE_VENDA
    If gobjVenda.iTipo = OPTION_PREVENDA Then
        Call BotaoLimpar_Click
    Else


        'se se trata de um orcamento DAV
        If gobjVenda.iTipo = OPTION_DAV Then
        
            If gobjVenda.objCupomFiscal.lNumOrcamento = 0 Then gError 105900
        
            If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
            
                'nao permitido imprimir em impressora nao fiscal dav no ES e na PB
                If giUsaImpressoraFiscal = 0 And (gsUF = "ES" Or gsUF = "PB") Then gError 214311
        
            End If
        
            objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
    
            'le os registros do orcamento e loca o arquivo
            lErro = CF_ECF("OrcamentoECF_Le", objVenda)
            If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 204714
            
            If lErro = 210447 Then gError 210449
            
            'se o orcamento ja esta cadastrado ==> vai imprimir o DAV se nao tiver ja sido impresso
            If lErro = SUCESSO Then
        
            
                    'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
                    If gobjVenda.iTipo = OPTION_DAV And objVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210991
            
                    gobjVenda.objCupomFiscal.dtDataEmissao = Date
                    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
                    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
        
        
                    'le os registros do orcamento e loca o arquivo
                    Timer1.Enabled = False
                    lErro = CF_ECF("Imprime_OrcamentoECF", dtDataFinal, objVenda.objCupomFiscal.lNumOrcamento, objTela, gobjVenda)
                    Timer1.Enabled = True
                    If lErro <> SUCESSO Then gError 105886
            
                    Set gobjVenda = New ClassVenda
                    gobjVenda.iCodModeloECF = giCodModeloECF

            End If
                    
        Else
        
            If AFRAC_ImpressoraCFe(giCodModeloECF) Then
            
                Call CancelaCupomEletronico(BOTAO_CANCELA_CUPOM_CFE)
                Exit Sub
                
            Else
            
                'se for um cupom e o
                'cupom a ser cancelado é um anterior(naum está na tela)
                If gobjVenda.objCupomFiscal.lNumero = 0 Then
                    If gcolVendas.Count = 0 Then gError 112075
                    For iIndice = gcolVendas.Count To 1 Step -1
                        Set objVendaAux = gcolVendas.Item(iIndice)
                        If objVendaAux.iTipo = OPTION_CF Then
                            
                            'CCF atual
                            lErro = AFRAC_CCF(sRetorno)
                            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "AFRAC_CCF")
                            If lErro <> SUCESSO Then gError 112061
                            
                            
                            lRetorno = StrParaLong(sRetorno)
                                
                            'se o último númeor de cupom é o da última venda executada--> pode cancelar esta venda
                            If lRetorno = objVendaAux.objCupomFiscal.lCCF Then
                        
                                'verifica se tem mais de um cupom vinculado impresso
                                'se tiver ==> nao pode cancelar por limitacao do ecf
                                For Each objMovCaixa In objVendaAux.colMovimentosCaixa
                                    
                                    lErro = CF_ECF("Trata_MovCaixa", objMovCaixa, colMeiosPag)
                                    If lErro <> SUCESSO Then gError 214735
                                    
                                Next
                                
                                iCuponsVinculados = 0
                                
                                For Each objMovCaixa1 In colMeiosPag
                                
                                    If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Or objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Or objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_TEF Then
                                        iCuponsVinculados = iCuponsVinculados + 1
                                    End If
                                
                                Next
                            
                                If iCuponsVinculados > 1 Then gError 214734
                        
                        
                                'Envia aviso perguntando se deseja cancelar o cupom
                                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM_ANTERIOR & vbNewLine & "No: " & CStr(objVendaAux.objCupomFiscal.lNumero) & " Data: " & Format(objVendaAux.objCupomFiscal.dtDataEmissao, "DD/MM/YYYY") & " R$: " & Format(objVendaAux.objCupomFiscal.dValorTotal, "STANDARD"))
                        
                                If vbMsgRes = vbNo Then Exit Sub
                                
                                'Se for Necessário a autorização do Gerente para abertura do Caixa
                                If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
                            
                                    'Chama a Tela de Senha
                                    Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
                            
                                    'Sai de Função se a Tela de Login não Retornar ok
                                    If giRetornoTela <> vbOK Then gError 102503
                                    
                                    'Se Operador for Gerente
                                    iCodGerente = objOperador.iCodigo
                            
                                End If
            
                                Call WritePrivateProfileString(APLICACAO_ECF, "COO", CStr(objVendaAux.objCupomFiscal.lNumero), NOME_ARQUIVO_CAIXA)
            
                                'cancelar o Cupom de Venda
                                lErro = AFRAC_CancelarCupom(Me, objVendaAux)
                                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                                If lErro <> SUCESSO Then gError 99610
                                
                                Exibe.Caption = "CANCELADO CUPOM CCF = " & objVendaAux.objCupomFiscal.lCCF
                                Exibe1.Caption = ""
                                
                                lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", objVendaAux)
                                If lErro <> SUCESSO Then gError 204540
                                
                                lErro = Alteracoes_CancelamentoCupom(objVendaAux)
                                If lErro <> SUCESSO Then gError 112078
                                
    '                            lTamanho = 10
    '                            sRetorno = String(lTamanho, 0)
    '
    '                            'Indica o status do TEF
    '                            Call GetPrivateProfileString(APLICACAO_ECF, "StatusTEF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    '
    '                            sRetorno = StringZ(sRetorno)
    '
    '                            If sRetorno = TEF_STATUS_VENDA Then
                                    
                                    Set objFormMsg = MsgTEF
                                    
                                    'cancela os cartoes ja confirmados e nao confirma o ultimo
                                    lErro = CF_ECF("TEF_CNC_PAYGO", objVendaAux, objFormMsg, objTela)
                                    If lErro <> SUCESSO Then gError 214588
                                    
                                    lErro = CF_ECF("TEF_Imprime_CNC_PAYGO", objFormMsg, objTela, objVendaAux)
                                    If lErro <> SUCESSO Then gError 214589
                                    
                                    
    '                            End If
                                
                                
                                
                                gcolVendas.Remove (iIndice)
                            Else
                                gError 112075
                            End If
                            
                            Set gobjVenda = New ClassVenda
                            gobjVenda.iCodModeloECF = giCodModeloECF
                    
                            gobjVenda.iTipo = OPTION_CF
                    
                            If gobjLojaECF.iAbreAposFechamento = MARCADO Then
                                sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
                                lErro = CF_ECF("Abre_Cupom", gobjVenda)
                                If lErro <> SUCESSO Then gError 99818
                            End If
    
                            Exit For
                        Else
                            If iIndice = 1 Then gError 112075
                        End If
                    Next
                'se esta cancelando um cupom aberto
                Else
                    'Envia aviso perguntando se deseja cancelar o cupom
                    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM_TELA)
            
                    If vbMsgRes = vbNo Then Exit Sub
                    
                    'Se for Necessário a autorização do Gerente para abertura do Caixa
                    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
                
                        'Chama a Tela de Senha
                        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
                
                        'Sai de Função se a Tela de Login não Retornar ok
                        If giRetornoTela <> vbOK Then gError 102504
                        
                        'Se Operador for Gerente
                        iCodGerente = objOperador.iCodigo
                
                    End If
                    
                    Exibe.Caption = "CANCELADO CUPOM CCF = " & gobjVenda.objCupomFiscal.lCCF
                    Exibe1.Caption = ""
                    
                    'cancelar o Cupom de Venda
                    lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                    If lErro <> SUCESSO Then gError 99742
                    
                    Call Move_Dados_Memoria_1
                    
                    'Realiza as operações necessárias para gravar
                    lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                    If lErro <> SUCESSO Then gError 204541
                    
                    lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                    If lErro <> SUCESSO Then gError 204542
                    
                    Set objFormMsg = MsgTEF
                    
                    lErro = CF_ECF("TEF_Imprime_CNC_PAYGO", objFormMsg, objTela, gobjVenda)
                    If lErro <> SUCESSO Then gError 214591
                    
                    Set gobjVenda = New ClassVenda
                    gobjVenda.iCodModeloECF = giCodModeloECF
                    
                    If gobjLojaECF.iAbreAposFechamento = MARCADO Then
                        sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
                        lErro = CF_ECF("Abre_Cupom", gobjVenda)
                        If lErro <> SUCESSO Then gError 99818
                    End If
                    
                End If
                
            End If
                
            Call Limpa_Tela_Venda
        
        End If
        
    End If
        
    Exit Sub

Erro_BotaoCancelaCupom_Click:

    Select Case gErr
                
        Case 201583
            Call Rotina_ErroECF(vbOKOnly, ERRO_EM_PROCESSAMENTO_SEFAZ, gErr)
        
        Case 99610, 99742, 112061, 112417, 112418, 112422, 99818, 102503, 102504, 105898, 199581, 204541, 204542, 204714, 210022, 214588, 214589, 214591, 214735
                            
        Case 105899
            Call Rotina_ErroECF(vbOKOnly, AVISO_ORCAMENTO_INEXISTENTE, gErr)
                            
        Case 105900
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr)
                            
        Case 112075
            Call Rotina_ErroECF(vbOKOnly, ERRO_CUPOM_NAO_CANCELADO, gErr)
                            
        Case 210449
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda.objCupomFiscal.lNumOrcamento)

        Case 214311
            Call Rotina_ErroECF(vbOKOnly, ERRO_NAO_PERMITIDO_IMPRIMIR_DAV_NAO_FISCAL, gErr)

        Case 214734
            Call Rotina_ErroECF(vbOKOnly, ERRO_NAO_PERMITIDO_CANCELAR_VARIOS_VINC, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175706)

    End Select
    
    Exit Sub
        
End Sub


Private Function Alteracoes_CancelamentoCupom(objVenda As ClassVenda) As Long

Dim objMovCaixa As ClassMovimentoCaixa
Dim objCheque As ClassChequePre
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim objCarne As ClassCarne
Dim lErro As Long
Dim sLog As String
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lSequencialCaixa As Long
Dim objAliquota As New ClassAliquotaICMS
Dim objItens As ClassItemCupomFiscal
Dim iIndice1 As Integer
Dim colRegistro As New Collection

On Error GoTo Erro_Alteracoes_CancelamentoCupom
    
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o Movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'se for um recebimento em cartão de crédito/Debito de TEF
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iTipoCartao = TIPO_TEF Then
            '''?????efetua caneclamento de TEF
        End If
    Next
    
    For Each objItens In objVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja - ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next

    
    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1
        Set objMovCaixa = gcolMovimentosCaixa.Item(iIndice)
        If objMovCaixa.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada movimento da venda
    For Each objMovCaixa In objVenda.colMovimentosCaixa
    
'??? 24/08/2016         If objMovCaixa.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor
    
'??? 24/08/2016         If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor

        'Se for de cartao de crédito ou débito especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto <> 0 Then
            'Busca em gcolCartão a ocorrencia de Cartão nao especificado
            For iIndice = gcolCartao.Count To 1 Step -1
                Set objAdmMeioPagtoCondPagto = gcolCartao.Item(iIndice)
                'Se encontrou
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovCaixa.iTipoCartao Then
                    'Atualiza o saldo do cartão
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolCartao.Remove (iIndice)
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de crédito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de débito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
    Next
    
    'Para cada movimento
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for um recebimento em ticket
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada obj de ticket da coleção global de tickets
                For Each objAdmMeioPagtoCondPagto In gcolTicket
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo de não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Ticket da coleção global
                For iIndice1 = gcolTicket.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolTicket.Item(iIndice1)
                    'Se encontrou o ticket/parcelamento
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolTicket.Remove (iIndice1)
                        'Sinaliza que encontrou
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    
    'Verifica se já existe movimentos de Outros\
    'Para cada MOvimento de Outros
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o MOvimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for do tipo outros
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada pagamento em outros na coleção global
                For Each objAdmMeioPagtoCondPagto In gcolOutros
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Pagamento em outros na col global
                For iIndice1 = gcolOutros.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolOutros.Item(iIndice1)
                    'Se for do mesmo tipo que o atual
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolOutros.Remove (iIndice1)
                        Exit For
                    End If
                Next
            End If
        End If
    Next
        
    'remove o Carne na col global
    If objVenda.objCarne.colParcelas.Count > 0 Then
        For iIndice = 1 To gcolCarne.Count
            Set objCarne = gcolCarne.Item(iIndice)
            If objCarne.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCarne.Remove (iIndice)
        Next
    End If
    
    'remove o Cheque na col global
    If objVenda.colCheques.Count > 0 Then
        For iIndice = gcolCheque.Count To 1 Step -1
            Set objCheque = gcolCheque.Item(iIndice)
            If objCheque.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCheque.Remove (iIndice)
        Next
    End If
    
    
    Alteracoes_CancelamentoCupom = SUCESSO
    
    Exit Function
    
Erro_Alteracoes_CancelamentoCupom:
    
    Alteracoes_CancelamentoCupom = gErr
    
    Select Case gErr
    
        Case 99901, 99953, 99952
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175707)

    End Select
        
    Exit Function
    
End Function

Private Sub Move_Dados_Memoria()

Dim objMovCaixa As ClassMovimentoCaixa
Dim objItens As New ClassItemCupomFiscal
Dim objAliquota As New ClassAliquotaICMS

    'se nao veio por um pedido externo (ntk)
    If gobjVenda.objCupomFiscal.lNumeroDAV < 1000000 Or gobjVenda.colMovimentosCaixa.Count = 0 Then
    
        Set gobjVenda.colCheques = Nothing
        Set gobjVenda.colTroca = Nothing
        Set gobjVenda.objCarne = Nothing
    
        Set gobjVenda.colMovimentosCaixa = New Collection
    
        gobjVenda.objCupomFiscal.dValorProdutos = StrParaDbl(Subtotal.Caption)
        'tem que incluir acrescimos e descontos lancados na tela de pagto
        gobjVenda.objCupomFiscal.dValorTotal = Arredonda_Moeda(gobjVenda.objCupomFiscal.dValorProdutos + gobjVenda.objCupomFiscal.dValorAcrescimo - gobjVenda.objCupomFiscal.dValorDesconto - gobjVenda.objCupomFiscal.dValorDesconto1)

        Set objMovCaixa = New ClassMovimentoCaixa
        
        objMovCaixa.dtDataMovimento = Date
        objMovCaixa.dValor = gobjVenda.objCupomFiscal.dValorTotal
        objMovCaixa.iFilialEmpresa = giFilialEmpresa
        objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO
        objMovCaixa.iCaixa = giCodCaixa
        objMovCaixa.iCodOperador = giCodOperador
        objMovCaixa.dHora = CDbl(Time)
        objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                    
        gobjVenda.colMovimentosCaixa.Add objMovCaixa
            
        'Atualiza o saldo em dinheiro
'        gdSaldoDinheiro = gdSaldoDinheiro + gobjVenda.objCupomFiscal.dValorTotal
            
    Else
    
        If gobjVenda.colMovimentosCaixa.Item(1).iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then
            'Atualiza o saldo em dinheiro
'            gdSaldoDinheiro = Arredonda_Moeda(gdSaldoDinheiro + gobjVenda.objCupomFiscal.dValorTotal)
        End If
    
    End If

'    gobjVenda.objCupomFiscal.dtDataEmissao = Date
'    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iECF = giCodECF
    
    For Each objItens In gobjVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja + ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next


'    If OptionDAV.Value = True Then
'        gobjVenda.iTipo = OPTION_DAV
'        gobjVenda.objCupomFiscal.iStatus = STATUS_BAIXADO
'        gobjVenda.objCupomFiscal.iTipo = OPTION_DAV
'    ElseIf OptionPreVenda.Value = True Then
'        gobjVenda.iTipo = OPTION_PREVENDA
'        gobjVenda.objCupomFiscal.iStatus = STATUS_BAIXADO
'        gobjVenda.objCupomFiscal.iTipo = OPTION_PREVENDA
'    'modelo nao possui ECF
'    ElseIf giCodModeloECF = 4 Then
'        gobjVenda.objCupomFiscal.iStatus = STATUS_BAIXADO
'        gobjVenda.iTipo = OPTION_CF
'        gobjVenda.objCupomFiscal.iTipo = OPTION_CF
'    Else
        gobjVenda.iTipo = OPTION_CF
        gobjVenda.objCupomFiscal.iTipo = OPTION_CF
'    End If

End Sub

Private Sub Alteracoes_CancelamentoOrcamento(objVenda As ClassVenda)

Dim objMovCaixa As ClassMovimentoCaixa
Dim objCheque As ClassChequePre
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim objCarne As ClassCarne
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lSequencialCaixa As Long
Dim objAliquota As New ClassAliquotaICMS
Dim objItens As ClassItemCupomFiscal
Dim iIndice1 As Integer

    For Each objItens In objVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja - ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next
    
    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1
        Set objMovCaixa = gcolMovimentosCaixa.Item(iIndice)
        If objMovCaixa.lNumIntExt = objVenda.objCupomFiscal.lNumOrcamento Then gcolMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada movimento da venda
    For Each objMovCaixa In objVenda.colMovimentosCaixa
'        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor
        'Se for de cartao de crédito ou débito especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto <> 0 Then
            'Busca em gcolCartão a ocorrencia de Cartão nao especificado
            For iIndice = gcolCartao.Count To 1 Step -1
                Set objAdmMeioPagtoCondPagto = gcolCartao.Item(iIndice)
                'Se encontrou
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovCaixa.iTipoCartao Then
                    'Atualiza o saldo do cartão
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolCartao.Remove (iIndice)
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de crédito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de débito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
    Next
    
    'Para cada movimento
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for um recebimento em ticket
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada obj de ticket da coleção global de tickets
                For Each objAdmMeioPagtoCondPagto In gcolTicket
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo de não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Ticket da coleção global
                For iIndice1 = gcolTicket.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolTicket.Item(iIndice1)
                    'Se encontrou o ticket/parcelamento
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolTicket.Remove (iIndice1)
                        'Sinaliza que encontrou
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    
    'Verifica se já existe movimentos de Outros\
    'Para cada MOvimento de Outros
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o MOvimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for do tipo outros
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada pagamento em outros na coleção global
                For Each objAdmMeioPagtoCondPagto In gcolOutros
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Pagamento em outros na col global
                For iIndice1 = gcolOutros.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolOutros.Item(iIndice1)
                    'Se for do mesmo tipo que o atual
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolOutros.Remove (iIndice1)
                        Exit For
                    End If
                Next
            End If
        End If
    Next
        
    'remove o Carne na col global
    If objVenda.objCarne.colParcelas.Count > 0 Then
        For iIndice = 1 To gcolCarne.Count
            Set objCarne = gcolCarne.Item(iIndice)
            If objCarne.lNumIntExt = objVenda.objCupomFiscal.lNumOrcamento Then gcolCarne.Remove (iIndice)
        Next
    End If
    
    'remove o Cheque na col global
    If objVenda.colCheques.Count > 0 Then
        For iIndice = gcolCheque.Count To 1 Step -1
            Set objCheque = gcolCheque.Item(iIndice)
            If objCheque.lNumIntExt = objVenda.objCupomFiscal.lNumOrcamento Then gcolCheque.Remove (iIndice)
        Next
    End If
    
    Exit Sub
    
End Sub

Private Sub BotaoAbrirGaveta_Click()
    
Dim iCodigo As Integer
Dim lErro As Long
Dim objMovCaixa As ClassMovimentoCaixa
Dim bAchou As Boolean
Dim iIndice As Integer
Dim sIndice As String
Dim sMsg As String
Dim lNum As Long
Dim sDesc As String
Dim objTiposMeioPagto As ClassTMPLoja
Dim objVenda As New ClassVenda
Dim objAliquota As New ClassAliquotaICMS
Dim objItens As ClassItemCupomFiscal
Dim lSequencial As Long
Dim sLog As String
Dim colRegistro As New Collection
Dim objCliente As ClassCliente
Dim sCPF As String
Dim lNumero As Long
Dim dtDataFinal As Date
Dim vbMsgRes As VbMsgBoxResult
Dim objProdutoNomeRed As Object
Dim objTela As Object
Dim objVendedor As ClassVendedor
Dim sMsgVendedor As String
Dim sOrcamento As String
Dim sRetorno As String
Dim lTamanho As Long
Dim objFormMsg As Object
Dim sDescricao As String


On Error GoTo Erro_BotaoAbrirGaveta_Click
        
    Set objTela = Me
        
    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210023
            
        'se for um DAV tem q ter o nome do cliente e cpf/cnpj preenchido
        If gobjVenda.iTipo = OPTION_DAV And Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
            
            If Len(Trim(gobjVenda.objCupomFiscal.sNomeCliente)) = 0 Then gError 214177
            
            If Len(Trim(gobjVenda.objCupomFiscal.sCPFCGC)) = 0 Then gError 214178
                
        End If
            
    End If
            
    If giCodModeloECF = IMPRESSORA_NFE Then
    
        If Len(Trim(gobjVenda.objCupomFiscal.sNomeCliente)) = 0 Then gError 214179
        
        If Len(Trim(gobjVenda.objCupomFiscal.sCPFCGC)) = 0 Then gError 214180
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntLogradouro)) = 0 Then gError 214181
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntNúmero)) = 0 Then gError 214182
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntComplemento)) = 0 Then gError 214183
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntBairro)) = 0 Then gError 2141784
        
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntCidade)) = 0 Then gError 214185
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntUF)) = 0 Then gError 214186
    
    End If
        
    'Se não há valor para pagar --> erro.
    If StrParaDbl(Subtotal.Caption) = 0 Then gError 99890
        
    
    'sevé obrigatório o preenchimento do vendedor
    If gobjLojaECF.iVendedorObrigatorio = 1 Then
        If Len(Trim(CodVendedor.Text)) = 0 Then gError 112072
    End If
    
    If gobjVenda.iTipo = OPTION_PREVENDA Or gobjVenda.iTipo = OPTION_DAV Then

        gobjVenda.objCupomFiscal.iTipo = gobjVenda.iTipo
        gobjVenda.objCupomFiscal.dtDataEmissao = Date
        gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
        gobjVenda.objCupomFiscal.dValorTroco = 0
        gobjVenda.objCupomFiscal.lDuracao = 0
        gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
        gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
        gobjVenda.objCupomFiscal.iECF = giCodECF
        gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
        gobjVenda.objCupomFiscal.dValorProdutos = StrParaDbl(Subtotal.Caption)
        
        'tem que incluir acrescimos e descontos lancados na tela de pagto
        gobjVenda.objCupomFiscal.dValorTotal = Arredonda_Moeda(gobjVenda.objCupomFiscal.dValorProdutos + gobjVenda.objCupomFiscal.dValorAcrescimo - gobjVenda.objCupomFiscal.dValorDesconto - gobjVenda.objCupomFiscal.dValorDesconto1)
        
        gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
        
        'resseta estas variaveis para ficar o deposito todo o pagamento em dinheiro
        Set gobjVenda.colMovimentosCaixa = New Collection
        Set gobjVenda.colCheques = Nothing
        Set gobjVenda.colTroca = Nothing
        Set gobjVenda.objCarne = Nothing
        
        Set objMovCaixa = New ClassMovimentoCaixa
        
        objMovCaixa.dtDataMovimento = gobjVenda.objCupomFiscal.dtDataEmissao
        objMovCaixa.dValor = gobjVenda.objCupomFiscal.dValorTotal
        objMovCaixa.iFilialEmpresa = giFilialEmpresa
        objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO
        objMovCaixa.iCaixa = giCodCaixa
        objMovCaixa.iCodOperador = giCodOperador
        objMovCaixa.dHora = gobjVenda.objCupomFiscal.dHoraEmissao
        objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                    
        gobjVenda.colMovimentosCaixa.Add objMovCaixa
        
        If gobjVenda.iTipo = OPTION_PREVENDA Then
        
            'grava o orcamento.
            lErro = CF_ECF("Grava_Orcamento_ECF1", gobjVenda)
            If lErro <> SUCESSO Then gError 105914
        
        Else
        
'        'Envia aviso perguntando se deseja imprimir o orçamento
'        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ORCAMENTO_IMPRESSAO)
'
'        If vbMsgRes = vbYes Then
'        If giImprimeOrc = 1 Then   'PAFECF

            'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
'            If gobjVenda.objCupomFiscal.lNumeroDAV <> 0 Then gError 210505


            objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento

            lErro = CF_ECF("OrcamentoECF_Le", objVenda)
            If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210894

            'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
            If gobjVenda.iTipo = OPTION_DAV And objVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210505

            'se o orcamento ja tiver sido usado para gerar cupom fiscal ==> nao pode alterar
            If lErro = 210447 Then gError 210998

            'se for DAV grava e imprime
            dtDataFinal = gobjVenda.objCupomFiscal.dtDataEmissao
            
            'le os registros do orcamento e loca o arquivo
'            lErro = CF_ECF("Imprime_OrcamentoECF", dtDataFinal, gobjVenda.objCupomFiscal.lNumOrcamento, objTela, gobjVenda)
'            If lErro <> SUCESSO Then gError 105915
        
            lErro = CF_ECF("Grava_Orcamento_ECF1", gobjVenda)
            If lErro <> SUCESSO Then gError 105915
        
        
        End If
        
        Set gobjVenda = New ClassVenda
        gobjVenda.iCodModeloECF = giCodModeloECF
        
        Call Limpa_Tela_Venda
        
        
    Else
    
        
        'se for um cupom fiscal
        
        Exibe.Caption = "ENCERRAMENTO DE VENDA"
        Exibe1.Caption = ""
        DoEvents
        
        Call Move_Dados_Memoria
        
        'informar se for um CF
        If gobjVenda.iTipo = OPTION_CF Then
            
            Call Venda_AjustaTrib
            
            If gobjVenda.objCupomFiscal.lNumeroDAV < 1000000 Then
            
                sIndice = TIPOMEIOPAGTOLOJA_DINHEIRO
                'Recolhe a descrição
                For Each objTiposMeioPagto In gcolTiposMeiosPagtos
                    If objTiposMeioPagto.iTipo = StrParaInt(sIndice) Then
                        sDesc = objTiposMeioPagto.sDescricao
                        Exit For
                    End If
                Next
                
                'If gobjlojaecf.iImprimeItemAItem = DESMARCADO Then
        '            lErro = Transforma_Cupom
        '            If lErro <> SUCESSO Then gError 112073
        '        'End If
                
                lErro = AFRAC_AcrescimoDescontoCupom(0, 0, "", "")
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acrescimo - Desconto")
                If lErro <> SUCESSO Then gError 109691
                    
                'Informa o pagamento
                lErro = AFRAC_FormaPagamento(sDesc, sIndice, Subtotal.Caption, sMsg)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Forma Pagamento")
                If lErro <> SUCESSO Then gError 99734
                    
            End If
                    
            If gobjVenda.objCupomFiscal.iVendedor <> 0 Then

'                For Each objVendedor In gcolVendedores
'                    'verifica se existe o vendedor na col
'                    If objVendedor.iCodigo = gobjVenda.objCupomFiscal.iVendedor Then
'                        Exit For
'                    End If
'                Next
            
                Set objVendedor = New ClassVendedor
                        
                lErro = CF_ECF("Vendedores_Le_Codigo", gobjVenda.objCupomFiscal.iVendedor, objVendedor)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214948
            
                sMsgVendedor = VENDEDOR_ECF_MSG & Formata_Campo(ALINHAMENTO_DIREITA, 38, " ", gobjVenda.objCupomFiscal.iVendedor & " - " & objVendedor.sNomeReduzido)
                
            End If
            
            'se o cupom esta sendo originado de um orcamento
            If gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then
            
                'se for uma PREVENDA
                If gobjVenda.objCupomFiscal.lNumeroDAV = 0 Then
                    sOrcamento = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "PV:" & Format(gobjVenda.objCupomFiscal.lNumOrcamento, "0000000000"))
                Else
                    sOrcamento = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "DAV:" & Format(gobjVenda.objCupomFiscal.lNumeroDAV, "0000000000"))
                End If
                
            End If
            
            'Fecha cupom
            Timer1.Enabled = False
            lErro = AFRAC_FecharCupom(Me, gobjVenda, False, gobjVenda.objCupomFiscal.sCPFCGC1, NomeCliente.Text, gobjVenda.objCupomFiscal.sEndereco, False, sOrcamento, sMsgVendedor)
            Timer1.Enabled = True
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fechar Cupom")
            If lErro <> SUCESSO Then gError 99735
                                
            If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
                        
                'Abri a Gaveta
                lErro = AFRAC_AbrirGaveta()
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Gaveta")
                If lErro <> SUCESSO Then gError 99736
                    
                'Realiza as operações necessárias para gravar
                lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 109822
        
            Else
            
'                'Realiza as operações necessárias para gravar
'                lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
'                If lErro <> SUCESSO Then gError 109822
        
                'Abri a Gaveta
                lErro = AFRAC_AbrirGaveta()
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Gaveta")
                If lErro <> SUCESSO Then gError 99736
                    
            End If
        
        Else
        
            'Realiza as operações necessárias para gravar
            lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 109822
        
        End If
                
        'depois de abrir a gaveta  vai tentar imprimir todos os registros de TEF
        'caso existam pois ja poderia ter passado na tela de pagamento e os pagamentos em cartao
        'terem sido cancelados
        
            
        Set objFormMsg = MsgTEF
        
        lErro = CF_ECF("TEF_Imprime_CNC_PAYGO", objFormMsg, objTela, gobjVenda)
        If lErro <> SUCESSO Then gError 214562
        
        
        'Atribui para a coleção global o objvenda
        gcolVendas.Add gobjVenda
        
        'Para cada movimento da venda
'??? 24/08/2016         For Each objMovCaixa In gobjVenda.colMovimentosCaixa
'??? 24/08/2016             'Se for de cartao de crédito ou débito especificado
'??? 24/08/2016             If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro + objMovCaixa.dValor
'??? 24/08/2016         Next

        If gobjVenda.iTipo = OPTION_CF Then Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)
        
        Set gobjVenda = New ClassVenda
        gobjVenda.iCodModeloECF = giCodModeloECF
                
        Call Limpa_Tela_Venda
        
        If gobjLojaECF.iAbreAposFechamento = MARCADO Then
            sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
            lErro = CF_ECF("Abre_Cupom", gobjVenda)
            If lErro <> SUCESSO Then gError 99818
        End If
    
    End If
    
    Exit Sub

Erro_BotaoAbrirGaveta_Click:

    Select Case gErr
                
        Case 99734 To 99736, 109822, 109691, 112073, 99901, 99952, 99953, 99818, 105870, 105914, 105915, 204346, 204347, 210023, 210894, 214562
                            
        Case 99811
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr, Error$)
                            
        Case 99890
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_EXISTENTE, gErr)
                            
        Case 112072
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDEDOR_NAO_PREENCHIDO, gErr, Error$)
        
        Case 210505
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
        
        Case 210998
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr)
        
        Case 214177
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NOME_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214178
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_CPFCNPJ_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214179
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_NOME_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214180
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_CPFCNPJ_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214181
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_LOGRADOURO_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214182
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_NUMERO_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214183
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_COMPL_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214184
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_BAIRRO_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214185
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_CIDADE_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214186
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_UF_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175708)

    End Select
    
    Exit Sub
        
End Sub


Private Sub BotaoSuspender_Click()

Dim lErro As Long
Dim lNumItens As Long
Dim sCaption As String
Dim sCaption1 As String

On Error GoTo Erro_BotaoSuspender_Click
    
    'Retorna o count de gobjvenda.objcupomfiscal.colitens
    lNumItens = Retorna_Count_ItensCupom()
    
    'Se tiver uma venda acontecendo -> erro.
    If lNumItens > 0 And OptionCF Then gError 99903
          
    sCaption = Exibe.Caption
    sCaption1 = Exibe1.Caption
          
    Exibe.Caption = "AGUARDANDO RETORNO..."
    Exibe1.Caption = ""
    'Função que Executa a Suspenção da Sessão
    lErro = CF_ECF("Sessao_Executa_Suspensao")
    If lErro <> SUCESSO Then gError 99826
    
    'funcao que executa o termino da suspensao se a senha for digitada.
    lErro = CF_ECF("Sessao_Executa_Termino_Susp")
    If lErro <> SUCESSO Then gError 117546
    
    Exibe.Caption = sCaption
    Exibe1.Caption = sCaption1
    DoEvents
    
    Exit Sub

Erro_BotaoSuspender_Click:

    Select Case gErr
                
        Case 99826, 117546
        
        Case 99903
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDA_ANDAMENTO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175676)

    End Select
    
    Exit Sub

End Sub

Private Function Exclui_Item_ColItens(ByVal iItem As Integer) As Long

Dim objItens As ClassItemCupomFiscal
Dim iIndice As Integer
Dim sSubtotal As String
Dim lErro As Long

On Error GoTo Erro_Exclui_Item_ColItens

    'Percorre toda a lista
'    For iIndice = (ListCF.ListCount - 1) To 12 Step -1
'        'Se tiver o itemdata do código passado
'        If ListCF.ItemData(iIndice) = iItem Then
'            'Exclui este item
'            ListCF.RemoveItem (iIndice)
'        End If
'    Next
        
        
        
'    ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "           ***** ITEM " & iItem & " CANCELADO *****")
'    ListCF.ItemData(ListCF.NewIndex) = iItem
'
'    'Para rolar automaticamente a barra de rolagem
'    ListCF.ListIndex = ListCF.NewIndex
        
    For Each objItens In gobjVenda.objCupomFiscal.colItens
        If objItens.iItem = iItem And objItens.iStatus = STATUS_ATIVO Then
            objItens.iStatus = STATUS_CANCELADO
            
            If gobjVenda.iTipo = OPTION_CF And Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
            
                lErro = AFRAC_LerInformacaoImpressora("001", sSubtotal)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
                If lErro <> SUCESSO Then gError 214140
                
                sSubtotal = StrParaDbl(sSubtotal) / 100
                
                Subtotal.Caption = Format(sSubtotal, "standard")
            
            Else
                Subtotal.Caption = Format(Subtotal.Caption - ((objItens.dPrecoUnitario * objItens.dQuantidade) - objItens.dValorDesconto), "standard")
            
            End If
            
            
            If objItens.dValorDesconto = 0 Then
                Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(objItens.dQuantidade, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
            Else
                Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(objItens.dQuantidade, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & "-" & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(objItens.dPrecoUnitario - objItens.dValorDesconto, "standard"))
            End If
            
            If Grid.Width < 8000 Then
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", "           ***** ITEM " & iItem & " CANCELADO *****") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = iItem
                
                'Para rolar automaticamente a barra de rolagem
                ListCF.ListIndex = ListCF.NewIndex
                
            Else
                Call Proxima_Linha_Grid
            
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "***** ITEM CANCELADO *****"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, iItem)
                
            End If
                                    
            Exit For
            
        End If
    Next
    
    Exclui_Item_ColItens = SUCESSO
    
    Exit Function

Erro_Exclui_Item_ColItens:

    Exclui_Item_ColItens = gErr

    Select Case gErr
                
        Case 214140
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214141)

    End Select
    
    Exit Function
    
    
End Function
        
Private Sub CodVendedor_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub CodVendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As ClassVendedor
Dim bAchou As Boolean

On Error GoTo Erro_Codvendedor_Validate

    'Verifica se o vendedor foi preenchidO
    If Len(Trim(CodVendedor.Text)) = 0 Then
        'joga o vendedor no gobjvenda
        gobjVenda.objCupomFiscal.iVendedor = StrParaInt(CodVendedor.Text)
        LabelVendedor.Caption = ""
        Exit Sub
    End If
    
    bAchou = False
    
'    For Each objVendedor In gcolVendedores
'        'verifica se existe o vendedor na col
'        If objVendedor.iCodigo = StrParaInt(CodVendedor.Text) Then
'            LabelVendedor.Caption = objVendedor.sNomeReduzido
'            bAchou = True
'            Exit For
'        End If
'    Next
            
    Set objVendedor = New ClassVendedor
            
    lErro = CF_ECF("Vendedores_Le_Codigo", StrParaInt(CodVendedor.Text), objVendedor)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214947
            
    If lErro = SUCESSO Then
        LabelVendedor.Caption = objVendedor.sNomeReduzido
        bAchou = True
    End If
            
    'Não encontrou o vendedor
    If bAchou = False Then gError 99737
    
    'joga o vendedor no gobjvenda
    gobjVenda.objCupomFiscal.iVendedor = StrParaInt(CodVendedor.Text)
    
    lErro = AFRAC_InformarVendedor(CodVendedor.Text)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informar Vendedor")
        If lErro <> SUCESSO Then gError 99922
    End If
    Exit Sub

Erro_Codvendedor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 99737
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDEDOR_NAO_CADASTRADO2, gErr)
            
        Case 99922, 214947
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175710)

    End Select
    
    Exit Sub

End Sub
        
Private Sub Orcamento_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Orcamento_GotFocus()
    Call MaskEdBox_TrataGotFocus(Orcamento, iAlterado)
End Sub

Private Sub Orcamento_Validate(Cancel As Boolean)

'Dim objVenda As ClassVenda
Dim colOrcamento As New Collection
Dim objItens As ClassItemCupomFiscal
Dim iIndice As Integer
Dim lErro As Long
Dim objProduto As ClassProduto
Dim objVenda1 As New ClassVenda

On Error GoTo Erro_Orcamento_Validate

    If Len(Trim(Orcamento.Text)) > 0 Then
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210024
    End If
    
    If Len(Trim(Orcamento.Text)) = 0 Then gobjVenda.objCupomFiscal.lNumOrcamento = 0 'Desvincula ao orçamento trazido para permitir gravar um novo (baseado nos itens do antigo)

    'Se existe um número
    If Len(Trim(Orcamento.Text)) > 0 And gobjVenda.objCupomFiscal.lNumOrcamento <> StrParaLong(Orcamento.Text) Then
    
        'Não pode marcar o cupom atual com o número do orçamento antigo por se der erro no processo vai ficar apontando para o cara errado e não vai validar de novo ... será feito ao final em Copia_Venda
        'gobjVenda.objCupomFiscal.lNumOrcamento = StrParaLong(Orcamento.Text)
        
        objVenda1.objCupomFiscal.lNumOrcamento = StrParaLong(Orcamento.Text)
        
        'Função Que le os orcamentos
        lErro = CF_ECF("OrcamentoECF_Le", objVenda1)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 204715
        
        'If lErro = 204690 And OptionCF.Value = True Then gError 210457
        If lErro = 204690 Then gError 210457  'dará erro para DAV agora também pois o mesmo não poderá ter o número digitado salvo para ser consultado
        
        If lErro = 210447 Then gError 210450
        
        If lErro = SUCESSO Then
            
            'descobre o nome reduzido do produto
'            For Each objItens In objVenda1.objCupomFiscal.colItens
'                For iIndice = 1 To gaobjProdutosNome.Count
'                    Set objProduto = gaobjProdutosNome.Item(iIndice)
'                    If objItens.sProduto = objProduto.sCodigo Then
'                        objItens.sProdutoNomeRed = objProduto.sNomeReduzido
'                        Exit For
'                    End If
'                Next
'            Next
            
            'objVenda era Nothing
            For Each objItens In objVenda1.objCupomFiscal.colItens
                
                Set objProduto = New ClassProduto
                
                lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214853
                
                If lErro = SUCESSO Then
                    objItens.sProdutoNomeRed = objProduto.sNomeReduzido
                End If
                
            Next
            
            Call Copia_Venda(gobjVenda, objVenda1)
            Call Traz_Orcamento
            
            If OptionCF.Value Then Call OptionCF_Click
                
        End If
        
    End If
    
    Exit Sub
    
Erro_Orcamento_Validate:
    
    Cancel = True
    
    Select Case gErr
                
        Case 204715, 210024, 214853
                
        Case 210450
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda1.objCupomFiscal.lNumOrcamento)

        Case 210457
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_CADASTRADO1, gErr, objVenda1.objCupomFiscal.lNumOrcamento)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175711)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Traz_Orcamento()
    
Dim objItens As New ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long
Dim sProduto1 As String

On Error GoTo Erro_Traz_Orcamento
    
    Exibe.Caption = "CONSULTA DE ORÇAMENTO."
    Exibe1.Caption = ""
    DoEvents
    
    ListCF.Clear
    Subtotal.Caption = ""
    
    Call Inicia_Cupom_Tela
    
    Orcamento.Text = CStr(gobjVenda.objCupomFiscal.lNumOrcamento)
    
    If gobjVenda.objCupomFiscal.iVendedor > 0 Then
        CodVendedor.Text = gobjVenda.objCupomFiscal.iVendedor
    Else
        CodVendedor.Text = ""
    End If
    Call CodVendedor_Validate(False)
    
    NomeCliente.Text = gobjVenda.objCupomFiscal.sNomeCliente
    Email.Text = gobjVenda.objCupomFiscal.sEndEntEmail
'    CodCliente.Text = gobjVenda.objCupomFiscal.sCPFCGC
    CGC.Text = gobjVenda.objCupomFiscal.sCPFCGC1
'    Endereco.Text = gobjVenda.objCupomFiscal.sEndereco
    
'    Call CodCliente_Validate(False)
    
    'Para cada Item --> inclui no Cupom
    For Each objItens In gobjVenda.objCupomFiscal.colItens
    
        ProdutoNomeRed.Text = objItens.sProdutoNomeRed
        
        Quantidade.Text = objItens.dQuantidade
        
        sProduto1 = ProdutoNomeRed.Text
                
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
        
        'caso o produto não seja encontrado
        If objProduto Is Nothing Then gError 99884
        
        ProdutoNomeRed.Text = objProduto.sNomeReduzido
        'PrecoUnitario.Caption = Format(objProduto.dPrecoLoja, "standard")
        PrecoUnitario.Caption = Format(objItens.dPrecoUnitario, "standard")
        
        'Prenche a col de itens do cupom com os dados do mesmo
        PrecoItem.Caption = Format(StrParaDbl(Quantidade.Text) * StrParaDbl(PrecoUnitario.Caption), "Standard")
        
        If objItens.iStatus = STATUS_ATIVO Then
            Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto), "standard")
        End If
        
        If Len(Trim(objProduto.sReferencia)) > 0 Then
            sProduto = objProduto.sReferencia
        Else
            sProduto = objProduto.sCodigoBarras
        End If
        
        If Grid.Width < 8000 Then
            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "     " & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", objProduto.sCodigo) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", objProduto.sDescricao)
'            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "     " & Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", objProduto.sDescricao)
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", Format(Quantidade.Text, "#0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 12, " ", Format(PrecoUnitario.Caption, "standard")) & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dAliquotaICMS * 100, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", Format(PrecoItem.Caption, "standard"))
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
        Else
            Call Proxima_Linha_Grid
    
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_CODIGO) = objProduto.sCodigo
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = objProduto.sCodigo & " - " & objProduto.sDescricao
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_QUANTIDADE) = Format(Quantidade.Text, "#0.000") & " " & objProduto.sSiglaUMVenda
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_X) = "x"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(PrecoUnitario.Caption, "standard")
            'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = objProduto.sSituacaoTribECF & Format(objItens.dAliquotaICMS * 100, "fixed") & "%"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(PrecoItem.Caption, "standard")
            
        End If
        
        
        'se existir desconto sobre o item...
        If objItens.dValorDesconto > 0 Then
            If Grid.Width < 8000 Then
        
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 21, " ", "-" & Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
        '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - (StrParaDbl(PrecoItem.Caption)), "standard")
        '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - objProduto.dDescontoValor), "standard")
        
            Else
            
                Call Proxima_Linha_Grid
                
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "DESCONTO: - R$ " & Format(objItens.dValorDesconto, "standard")
                'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = "-" & Format(objItens.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            
            End If
        End If
        
        If objItens.iStatus = STATUS_CANCELADO Then
            If Grid.Width < 8000 Then
'                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "           ***** ITEM " & iItem & " CANCELADO *****")
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", "           ***** ITEM " & objItens.iItem & " CANCELADO *****") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            Else
                Call Proxima_Linha_Grid
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "***** ITEM CANCELADO *****"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
                
            End If
        End If
        
        
        'Para rolar automaticamente a barra de rolagem
        If Grid.Width < 8000 Then
            ListCF.ListIndex = ListCF.NewIndex
        End If
        
        Call Limpa_Cupom_Tela
    Next
    
    
'    If giUsaImpressoraFiscal = 0 Then
'        If gobjVenda.iTipo = OPTION_PREVENDA Then
'            OptionPreVenda.Value = True
'        ElseIf gobjVenda.iTipo = OPTION_DAV Then
'            OptionDAV.Value = True
'        End If
'    End If
    
    Exit Sub

Erro_Traz_Orcamento:

    Select Case gErr
                
        Case 99884
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175712)

    End Select
    
    Exit Sub

End Sub

'Private Sub Outros_Click()
'
'Dim objTeclado As New ClassTecladoProduto
'Dim objTeclado1 As New ClassTecladoProduto
'
'    'Chama tela de TecladoLista
'    Call Chama_TelaECF_Modal("TecladoLojaLista", objTeclado)
'
'    If giRetornoTela = vbOK Then
'
'        Exibe.Caption = "TROCA DE TECLADO."
'        Exibe1.Caption = ""
'        DoEvents
'
'        Call Inicializa_ColTecladoProdutoItens
'
'        Call Traz_ColTecladoProdutoItens(gcolTecladoProdutoItens)
'
'        For Each objTeclado1 In gcolTeclados
'        If objTeclado1.iCodigo = objTeclado.iCodigo Then
'            Call Traz_ColTecladoProdutoItens(objTeclado1.colTecladoProdutoItem)
'        End If
'
'    Next
'
'    End If
'
'    Exit Sub
'
'End Sub

Private Sub Produto_Click(Index As Integer)
    
Dim objTecladoProdutoItens As New ClassTecladoProdutoItem
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim sTecla As String
Dim objTecladoProdutoItensPai As New ClassTecladoProdutoItem
Dim iIndiceMenu As Integer
Dim sKey As String

On Error GoTo Erro_Produto_Click

    'se for um botao preenchido
    If Len(Produto(Index).Tag) > 0 Then

        Set objTecladoProdutoItens = gcolTecladoProdutoItens(Produto(Index).Tag)
    
        'se for um botao que abre outro menu
        If Len(objTecladoProdutoItens.sProduto) = 0 Then
    
            iIndiceMenu = (Len(Produto(Index).Tag) - 1) / 2
            
            If iIndiceMenu > 0 Then
                Menu(iIndiceMenu).Visible = True
                Set objTecladoProdutoItensPai = gcolTecladoProdutoItens(Produto(Index).Tag)
                Menu(iIndiceMenu).Caption = objTecladoProdutoItensPai.sTitulo
                Menu(iIndiceMenu).ToolTipText = Menu(iIndiceMenu).Caption
                Menu(iIndiceMenu).Tag = objTecladoProdutoItensPai.sArvoreKey
            End If
            
    
            sKey = Produto(Index).Tag
            'coloca o novo menu na tela
    
            Call Limpa_Teclado
        
                    
        
            'preenche os botoes de acordo com o que esta armazendo na treeview no nivel selecionado
            For Each objTecladoProdutoItens In gcolTecladoProdutoItens
                If left(objTecladoProdutoItens.sArvoreKey, Len(objTecladoProdutoItens.sArvoreKey) - 2) = sKey Then
            
                    iIndice = CInt(right(objTecladoProdutoItens.sArvoreKey, 2)) - 1
                    
                    If objTecladoProdutoItens.sTitulo <> "" Then
                        If objTecladoProdutoItens.iTecla <> 0 Then
                            If objTecladoProdutoItens.iTecla > vbKeyF2 Then
                                Call Acha_Tecla(objTecladoProdutoItens.iTecla, sTecla)
                            Else
                                sTecla = Chr(objTecladoProdutoItens.iTecla)
                            End If
                            Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo & "(" & sTecla & ")"
                        Else
                            Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo
                        End If
                    Else
                        Produto(iIndice).Caption = ""
                    End If
                    
                    Produto(iIndice).Tag = objTecladoProdutoItens.sArvoreKey
                    Produto(iIndice).BackColor = objTecladoProdutoItens.lColor
                    Produto(iIndice).ToolTipText = Produto(iIndice).Caption
                        
                End If
            Next
    
        Else
            'se for um botao que tem codigo de produto associado

            'Verifica se o produto tem referencia para fazer a busca
            If Len(Trim(objTecladoProdutoItens.sReferencia)) <> 0 Then
                'Joga na tela os dados do produto selecionado
                ProdutoNomeRed.Text = objTecladoProdutoItens.sReferencia
            'senão busca pelo código de Barras
            Else
                'Joga na tela os dados do produto selecionado
                ProdutoNomeRed.Text = objTecladoProdutoItens.sCodigoBarras
            End If
            
            Call ProdutoNomeRed_LostFocus
            Call ProdutoNomeRed_Validate(False)
    
        End If
    
    End If
    
    Exit Sub

Erro_Produto_Click:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 665544)

    End Select
    
    Exit Sub

End Sub



Private Sub ProdutoNomeRed_LostFocus()
    Quantidade.SelStart = 0
    Quantidade.SelLength = Len(Quantidade.Text)
    Quantidade.SetFocus
End Sub

Private Sub ProdutoNomeRed_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoNomeRed_Validate(Cancel As Boolean)

Dim lErro As Long
Dim bAchou As Boolean
Dim objProduto As ClassProduto
Dim objItens As New ClassItemCupomFiscal
Dim sProduto As String

On Error GoTo Erro_ProdutoNomeRed_Validate
    
    If Len(Trim(ProdutoNomeRed.Text)) > 0 Then
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210025

    End If

    'Se o produto não está preenchido
    If Len(Trim(ProdutoNomeRed.Text)) = 0 Then
        PrecoUnitario.Caption = Format(0, "standard")
        PrecoItem.Caption = Format(0, "standard")
    'caso contrário
    Else
        'Verifica a quantidade
        If Len(Trim(Quantidade)) <> 0 Then
            'Rotina de cupom
            lErro = Adiciona_Cupom(0)
            If lErro <> SUCESSO Then gError 99887
        End If
    End If
    
    Exit Sub

Erro_ProdutoNomeRed_Validate:

    Cancel = True

    Select Case gErr
                
        Case 99887, 210025
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175713)

    End Select
    
    Exit Sub

End Sub

Private Sub Tecla_Ativa()
    
Dim objTecladoProdutoItens As New ClassTecladoProdutoItem

    For Each objTecladoProdutoItens In gcolTecladoProdutoItens
        If (Quantidade.Text) = Chr(objTecladoProdutoItens.iTecla) Then
            'Verifica se o produto tem referencia para fazer a busca
            If Len(Trim(objTecladoProdutoItens.sReferencia)) <> 0 Then
                'Joga na tela os dados do produto selecionado
                ProdutoNomeRed.Text = objTecladoProdutoItens.sReferencia
            'senão busca pelo código de Barras
            Else
                'Joga na tela os dados do produto selecionado
                ProdutoNomeRed.Text = objTecladoProdutoItens.sCodigoBarras
            End If
            
            Call ProdutoNomeRed_Validate(False)
            Exit For
        End If
    Next
    
    Quantidade.Text = ""
    
End Sub

Private Sub Quantidade_Change()
    
    iAlterado = REGISTRO_ALTERADO
  
End Sub

Private Sub Quantidade_LostFocus()

    PrecoUnitario.Caption = Format(0, "standard")
    PrecoItem.Caption = Format(0, "standard")
    ProdutoNomeRed.Text = ""
    
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate
    
     'Se a quantidade e o produto estão prenchidos
    If Len(Trim(Quantidade.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 99888
        If Len(Trim(Fix(Quantidade.Text))) > 4 Then gError 112423
        Quantidade.Text = Format(Quantidade.Text, "0.000")
        If right(Quantidade.Text, 4) = ",000" Then Quantidade.Text = Format(Quantidade.Text, "#,#")
        
        If Len(Trim(ProdutoNomeRed.Text)) <> 0 Then
            
            lErro = CF_ECF("Requisito_XXII")
            If lErro <> SUCESSO Then gError 214147
            
            'Rotina de cupom
            lErro = Adiciona_Cupom(0)
            If lErro <> SUCESSO Then gError 210882

        End If
        
        
    End If
    
    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr
            
        Case 99888, 210882, 214147
        
        Case 112423
            Call Rotina_ErroECF(vbOKOnly, ERRO_QUANTIDADE_INVALIDA, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175714)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoPagamento_Click()

Dim iCodigo As Integer
Dim lErro As Long
Dim objGenerico As New AdmGenerico
Dim objTela As Object
Dim lTamanho As Long
Dim sRetorno As String
Dim objFormMsg As Object
Dim sDescricao As String

On Error GoTo Erro_BotaoPagamento_Click
    
    Set objTela = Me
    Set objFormMsg = MsgTEF
                
    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210026
    
        'se for um DAV tem q ter o nome do cliente e cpf/cnpj preenchido
        If gobjVenda.iTipo = OPTION_DAV Then
            If Len(Trim(gobjVenda.objCupomFiscal.sNomeCliente)) = 0 Then gError 214197
            
            If Len(Trim(gobjVenda.objCupomFiscal.sCPFCGC)) = 0 Then gError 214198
                
        End If
            
    End If
            
    If giCodModeloECF = IMPRESSORA_NFE Then
    
        If Len(Trim(gobjVenda.objCupomFiscal.sNomeCliente)) = 0 Then gError 204199
        
        If Len(Trim(gobjVenda.objCupomFiscal.sCPFCGC)) = 0 Then gError 204200
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntLogradouro)) = 0 Then gError 204201
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntNúmero)) = 0 Then gError 204202
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntComplemento)) = 0 Then gError 204203
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntBairro)) = 0 Then gError 2041204
        
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntCidade)) = 0 Then gError 204205
    
        If Len(Trim(gobjVenda.objCupomFiscal.sEndEntUF)) = 0 Then gError 204206
    
    End If
    
    'Se não há valor para pagar --> erro.
    If StrParaDbl(Subtotal.Caption) = 0 Then gError 99739
        
    'sevé obrigatório o preenchimento do vendedor
    If gobjLojaECF.iVendedorObrigatorio = 1 Then
        If Len(Trim(CodVendedor.Text)) = 0 Then gError 112072
    End If
            
    gobjVenda.objCupomFiscal.dValorProdutos = StrParaDbl(Subtotal.Caption)
    'tem que incluir acrescimos e descontos lancados na tela de pagto
    gobjVenda.objCupomFiscal.dValorTotal = Arredonda_Moeda(gobjVenda.objCupomFiscal.dValorProdutos + gobjVenda.objCupomFiscal.dValorAcrescimo - gobjVenda.objCupomFiscal.dValorDesconto - gobjVenda.objCupomFiscal.dValorDesconto1)
    
    If gobjVenda.iTipo = OPTION_CF Then
                
'        Exibe.Caption = ""
'        Exibe1.Caption = ""
'        DoEvents
        Call Chama_TelaECF_Modal("Pagamento", gobjVenda, objGenerico)
        'SE não deu tudo certo
        If objGenerico.vVariavel = vbCancel Then
        
            Exibe.Caption = "CANCELADO CUPOM CCF = " & gobjVenda.objCupomFiscal.lCCF
            Exibe1.Caption = ""
            
            If AFRAC_ImpressoraCFe(giCodModeloECF) Then
            
                Call CancelaCupomEletronico(BOTAO_PAGTO_CANCELA_CUPOM_CFE)
                Exit Sub
                
            Else
            
                'cancelar o Cupom de Venda
                lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                If lErro <> SUCESSO Then gError 99738
                
                Call Move_Dados_Memoria_1
                
                'Realiza as operações necessárias para gravar
                lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204543
                
                lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204544
                
                'foi feito o pedido de cancelamento portanto vai tentar imprimir todos os registros de TEF
                'caso existam
                lErro = CF_ECF("TEF_Imprime_CNC_PAYGO", objFormMsg, objTela, gobjVenda)
                If lErro <> SUCESSO Then gError 214563
                
            End If
                
        End If
        
    Else
    
        Set objTela = Me
    
        lErro = CF_ECF("Valida_Orcamento", objTela)
        If lErro <> SUCESSO Then gError 204716
        
        Call Chama_TelaECF_Modal("FechaOrcamento", gobjVenda, objGenerico)
    
    End If
    
    
    If objGenerico.vVariavel <> vbAbort Then
    
        If giRetornoTela <> vbAbort Then

            Set gobjVenda = New ClassVenda
            gobjVenda.iCodModeloECF = giCodModeloECF
            
            Call Limpa_Tela_Venda
        End If
    
    End If
    
    Exit Sub

Erro_BotaoPagamento_Click:

    Select Case gErr
                
        Case 99738
            
            Set gobjVenda = New ClassVenda
            gobjVenda.iCodModeloECF = giCodModeloECF
                
            gobjVenda.iTipo = 0
                
            Call Limpa_Tela_Venda
                
        Case 99739
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_EXISTENTE, gErr)
                    
        Case 99897
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr, Error$)
                        
        Case 112072
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDEDOR_NAO_PREENCHIDO, gErr, Error$)
        
        Case 204543, 204544, 204716, 210026, 214563
        
        Case 214197
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NOME_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214198
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_CPFCNPJ_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214199
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_NOME_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214200
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_CPFCNPJ_CLIENTE_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214201
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_LOGRADOURO_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214202
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_NUMERO_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214203
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_COMPL_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214204
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_BAIRRO_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214205
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_CIDADE_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
            
        Case 214206
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFE_UF_ENTREGA_NAO_PREENCHIDO, gErr, Error$)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175715)

    End Select
    
    Exit Sub
    
End Sub

Private Sub OptionCF_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim lNum As Long
Dim lNumero As Long
Dim objCliente As ClassCliente
Dim sCPF As String
Dim objVendaAux As New ClassVenda
Dim colItens As New Collection
Dim colItens1 As New Collection
Dim iIndice As Integer
Dim objVendaOrc As New ClassVenda

On Error GoTo Erro_OptionCF_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210027

    lNum = Retorna_Count_ItensCupom()
    
    'Se tiver um orçamento selecionado
    If lNum <> 0 And (gobjVenda.iTipo = OPTION_PREVENDA Or gobjVenda.iTipo = OPTION_DAV) Then
        
        'Envia aviso perguntando se deseja transforma o orçamento em venda
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ORCAMENTO_VENDA)

        If vbMsgRes = vbNo Then
            Call Limpa_Tela_Venda

            Set gobjVenda = New ClassVenda
            gobjVenda.iCodModeloECF = giCodModeloECF
            
        Else
            Exibe.Caption = "AGUARDE..."
            Exibe1.Caption = ""
            DoEvents
            
            
'            If giRemoveOrc = REMOVER_ORC Then
            
                'exclui o orcamento que está sendo transformado em cupom
'                lErro = CF_ECF("Caixa_Exclui_Orcamento", gobjVenda)
'                If lErro <> SUCESSO And lErro <> 105761 Then gError 105879
            
'            End If
            
            'se o orcamento nao estava gravado ==> zera o numero do orcamento
'            If lErro <> SUCESSO Then gobjVenda.objCupomFiscal.lNumOrcamento = 0
            
            
             'se o numero do orcamento esta preenchido verifica se esta gravado. Se nao estiver gravado ==> limpa o numero do orcamento
            If gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then
                
                objVendaOrc.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
            
                'Função Que le os orcamentos
                lErro = CF_ECF("OrcamentoECF_Le", objVendaOrc)
                If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210513
            
                If lErro = 210447 Then gError 210514
                
                'se nao esta cadatrado o orcamento, limpa o numero
                If lErro = 204690 Then gobjVenda.objCupomFiscal.lNumOrcamento = 0
            
            End If
           
            
            Call Copia_Venda(objVendaAux, gobjVenda)
    
            
            'guardo o número do orçamento que agora está relacionado com o cupom
            Set gobjVenda = objVendaAux
            gobjVenda.objCupomFiscal.dtDataOrcamento = gobjVenda.objCupomFiscal.dtDataEmissao
            
            
            
            For iIndice = gobjVenda.objCupomFiscal.colItens.Count To 1 Step -1
                colItens1.Add gobjVenda.objCupomFiscal.colItens.Item(iIndice)
                gobjVenda.objCupomFiscal.colItens.Remove (iIndice)
            Next
            
            For iIndice = colItens1.Count To 1 Step -1
                colItens.Add colItens1.Item(iIndice)
            Next
            
            If Len(Trim(Orcamento.Text)) > 0 Then gobjVenda.objCupomFiscal.lNumOrcamento = StrParaLong(Orcamento.Text)
            
            gobjVenda.iTipo = OPTION_CF
'            gobjVenda.objCupomFiscal.dtDataEmissao = Date
'            gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
            gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
            gobjVenda.objCupomFiscal.iECF = giCodECF
    
            
            'Abre o cupom
'            sCPF = CodCliente.Text
'            lErro = CF_ECF("Abre_Cupom", gobjVenda, sCPF)
'            If lErro <> SUCESSO Then gError 99747
    
            lErro = Transforma_Cupom(colItens)
            If lErro <> SUCESSO Then gError 99867
            Exibe1.Caption = ""
            Exibe.Caption = "TRANSF.: ORÇAMENTO EM VENDA."
        End If
        
    End If
    
    
    Orcamento.Text = ""
'    Orcamento.Enabled = False
'    LabelOrcamento.Enabled = False

    BotaoProxNum.Enabled = False
'    OptionCF.Value = True
    gobjVenda.iTipo = OPTION_CF
    
    BotaoAbrirGaveta.Caption = "(F10)   Abrir Gaveta"
    BotaoCancelaCupom.Caption = "(Esc)   Cancela Cupom"
    
    Exit Sub
    
Erro_OptionCF_Click:

    Select Case gErr
                
        Case 99747, 99867, 109396, 99818, 105879, 210027, 210513
                            
        Case 105880
            Call Rotina_ErroECF(vbOKOnly, AVISO_ORCAMENTO_INEXISTENTE, gErr)
                            
        Case 210514
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, gobjVenda.objCupomFiscal.lNumOrcamento)
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175716)

    End Select
    
    Exit Sub
        
End Sub


Private Sub Copia_Venda(objVendaAux As ClassVenda, objVenda As ClassVenda)
    
Dim iIndice As Integer
Dim objCheque As New ClassChequePre
Dim objChequeAux As New ClassChequePre
Dim objMVCX As New ClassMovimentoCaixa
Dim objMVCXAux As New ClassMovimentoCaixa
Dim objTroca As New ClassTroca
Dim objTrocaAux As New ClassTroca
Dim objCarne As New ClassCarne
Dim objCarneAux As New ClassCarne
Dim objCarneParc As New ClassCarneParcelas
Dim objCarneParcAux As New ClassCarneParcelas
Dim objCupom As New ClassCupomFiscal
Dim objCupomAux As New ClassCupomFiscal
Dim objCupomItem As New ClassItemCupomFiscal
Dim objCupomItemAux As New ClassItemCupomFiscal

    Set objVendaAux = New ClassVenda
    
    For iIndice = 1 To objVenda.colCheques.Count
        Set objCheque = objVenda.colCheques.Item(iIndice)
        Set objChequeAux = New ClassChequePre
        
        objChequeAux.dtDataDeposito = objCheque.dtDataDeposito
        objChequeAux.dValor = objCheque.dValor
        objChequeAux.iAprovado = objCheque.iAprovado
        objChequeAux.iBanco = objCheque.iBanco
        objChequeAux.iCaixa = giCodCaixa ' objCheque.iCaixa
        objChequeAux.iChequeSel = objCheque.iChequeSel
        objChequeAux.iECF = giCodECF ' objCheque.iECF
        objChequeAux.iFilial = objCheque.iFilial
        objChequeAux.iFilialEmpresa = giFilialEmpresa ' objCheque.iFilialEmpresa
        objChequeAux.iFilialEmpresaLoja = objCheque.iFilialEmpresaLoja
        objChequeAux.iNaoEspecificado = objCheque.iNaoEspecificado
        objChequeAux.iStatus = objCheque.iStatus
        objChequeAux.iTipoBordero = objCheque.iTipoBordero
        objChequeAux.lCliente = objCheque.lCliente
        objChequeAux.lCupomFiscal = objCheque.lCupomFiscal
        objChequeAux.lNumBordero = objCheque.lNumBordero
        objChequeAux.lNumBorderoLoja = objCheque.lNumBorderoLoja
        objChequeAux.lNumero = objCheque.lNumero
        objChequeAux.lNumIntCheque = objCheque.lNumIntCheque
        objChequeAux.lNumIntChequeBord = objCheque.lNumIntChequeBord
        objChequeAux.lNumIntExt = objCheque.lNumIntExt
        objChequeAux.lNumMovtoCaixa = objCheque.lNumMovtoCaixa
        objChequeAux.lNumMovtoSangria = objCheque.lNumMovtoSangria
        objChequeAux.lSequencial = objCheque.lSequencial
        objChequeAux.lSequencialBack = objCheque.lSequencialBack
        objChequeAux.lSequencialCaixa = objCheque.lSequencialCaixa
        objChequeAux.lSequencialLoja = objCheque.lSequencialLoja
        objChequeAux.sAgencia = objCheque.sAgencia
        objChequeAux.sCarne = objCheque.sCarne
        objChequeAux.sContaCorrente = objCheque.sContaCorrente
        objChequeAux.sCPFCGC = objCheque.sCPFCGC
        
        objVendaAux.colCheques.Add objChequeAux
    Next
    
    For iIndice = 1 To objVenda.colMovimentosCaixa.Count
        Set objMVCX = objVenda.colMovimentosCaixa.Item(iIndice)
        Set objMVCXAux = New ClassMovimentoCaixa
        
        objMVCXAux.dHora = objMVCX.dHora
        objMVCXAux.dtDataMovimento = objMVCX.dtDataMovimento
        objMVCXAux.dValor = objMVCX.dValor
        objMVCXAux.iAdmMeioPagto = objMVCX.iAdmMeioPagto
        objMVCXAux.iCaixa = giCodCaixa 'objMVCX.iCaixa
        objMVCXAux.iCodConta = objMVCX.iCodConta
        objMVCXAux.iCodOperador = objMVCX.iCodOperador
        objMVCXAux.iExcluiu = objMVCX.iExcluiu
        objMVCXAux.iFilialEmpresa = giFilialEmpresa 'objMVCX.iFilialEmpresa
        objMVCXAux.iGerente = objMVCX.iGerente
        objMVCXAux.iParcelamento = objMVCX.iParcelamento
        objMVCXAux.iQuantLog = objMVCX.iQuantLog
        objMVCXAux.iTipo = objMVCX.iTipo
        objMVCXAux.iTipoCartao = objMVCX.iTipoCartao
        objMVCXAux.lCupomFiscal = objMVCX.lCupomFiscal
        objMVCXAux.lMovtoEstorno = objMVCX.lMovtoEstorno
        objMVCXAux.lMovtoTransf = objMVCX.lMovtoTransf
        objMVCXAux.lNumero = objMVCX.lNumero
        objMVCXAux.lNumIntDocLog = objMVCX.lNumIntDocLog
        objMVCXAux.lNumIntExt = objMVCX.lNumIntExt
        objMVCXAux.lNumMovto = objMVCX.lNumMovto
        objMVCXAux.lNumRefInterna = objMVCX.lNumRefInterna
        objMVCXAux.lSequencial = objMVCX.lSequencial
        objMVCXAux.lSequencialConta = objMVCX.lSequencialConta
        objMVCXAux.lTransferencia = objMVCX.lTransferencia
        objMVCXAux.sFavorecido = objMVCX.sFavorecido
        objMVCXAux.sHistorico = objMVCX.sHistorico
        
        objVendaAux.colMovimentosCaixa.Add objMVCXAux
    Next
    
    For iIndice = 1 To objVenda.colTroca.Count
        Set objTroca = objVenda.colTroca.Item(iIndice)
        Set objTrocaAux = New ClassTroca
        
        objTrocaAux.dQuantidade = objTroca.dQuantidade
        objTrocaAux.dValor = objTroca.dValor
        objTrocaAux.iFilialEmpresa = giFilialEmpresa 'objTroca.iFilialEmpresa
        objTrocaAux.lNumIntDoc = objTroca.lNumIntDoc
        objTrocaAux.lNumMovtoCaixa = objTroca.lNumMovtoCaixa
        objTrocaAux.sProduto = objTroca.sProduto
        objTrocaAux.sUnidadeMed = objTroca.sUnidadeMed
        
        objVendaAux.colTroca.Add objTrocaAux
    Next
    
    objVendaAux.iTipo = objVenda.iTipo
    
    Set objCarne = objVenda.objCarne
    
    objCarneAux.dtDataReferencia = objCarne.dtDataReferencia
    objCarneAux.iFilialEmpresa = giFilialEmpresa 'objCarne.iFilialEmpresa
    objCarneAux.iStatus = objCarne.iStatus
    objCarneAux.lCliente = objCarne.lCliente
    objCarneAux.lCupomFiscal = objCarne.lCupomFiscal
    objCarneAux.lNumIntDoc = objCarne.lNumIntDoc
    objCarneAux.lNumIntExt = objCarne.lNumIntExt
    objCarneAux.sAutorizacao = objCarne.sAutorizacao
    objCarneAux.sCodBarrasCarne = objCarne.sCodBarrasCarne
        
    Set objVendaAux.objCarne = objCarneAux
    
    For iIndice = 1 To objVenda.objCarne.colParcelas.Count
        Set objCarneParc = objVenda.objCarne.colParcelas.Item(iIndice)
        Set objCarneParcAux = New ClassCarneParcelas
        
        objCarneParcAux.dtDataVencimento = objCarneParc.dtDataVencimento
        objCarneParcAux.dValor = objCarneParc.dValor
        objCarneParcAux.iFilialEmpresa = objCarneParc.iFilialEmpresa
        objCarneParcAux.iParcela = objCarneParc.iParcela
        objCarneParcAux.iStatus = objCarneParc.iStatus
        objCarneParcAux.lNumIntCarne = objCarneParc.lNumIntCarne
        objCarneParcAux.lNumIntDoc = objCarneParc.lNumIntDoc
        
        objVendaAux.objCarne.colParcelas.Add objCarneParcAux
    Next
    
    Set objCupom = objVenda.objCupomFiscal
    
    objCupomAux.dHoraEmissao = objCupom.dHoraEmissao
    objCupomAux.dtDataEmissao = objCupom.dtDataEmissao
    objCupomAux.dValorAcrescimo = objCupom.dValorAcrescimo
    objCupomAux.dValorDesconto = objCupom.dValorDesconto
    objCupomAux.dValorProdutos = objCupom.dValorProdutos
    objCupomAux.dValorTotal = objCupom.dValorTotal
    objCupomAux.dValorTroco = objCupom.dValorTroco
    objCupomAux.iCodCaixa = giCodCaixa 'objCupom.iCodCaixa
    objCupomAux.iECF = giCodECF 'objCupom.iECF
    objCupomAux.iFilialEmpresa = giFilialEmpresa 'objCupom.iFilialEmpresa
    objCupomAux.iStatus = objCupom.iStatus
    objCupomAux.iTabelaPreco = objCupom.iTabelaPreco
    objCupomAux.iTipo = objCupom.iTipo
    objCupomAux.iVendedor = objCupom.iVendedor
    objCupomAux.lCliente = objCupom.lCliente
    objCupomAux.lDuracao = objCupom.lDuracao
    objCupomAux.lGerenteCancel = objCupom.lGerenteCancel
    objCupomAux.lNumero = objCupom.lNumero
    objCupomAux.lNumIntDoc = objCupom.lNumIntDoc
    objCupomAux.lNumOrcamento = objCupom.lNumOrcamento
    objCupomAux.sCPFCGC = objCupom.sCPFCGC
    objCupomAux.sMotivoCancel = objCupom.sMotivoCancel
    objCupomAux.sNaturezaOp = objCupom.sNaturezaOp
    objCupomAux.sCPFCGC1 = objCupom.sCPFCGC1
    objCupomAux.sNomeCliente = objCupom.sNomeCliente
    objCupomAux.lNumeroDAV = objCupom.lNumeroDAV
    objCupomAux.sEndereco = objCupom.sEndereco
    objCupomAux.sEndEntLogradouro = objCupom.sEndEntLogradouro
    objCupomAux.sEndEntNúmero = objCupom.sEndEntNúmero
    objCupomAux.sEndEntComplemento = objCupom.sEndEntComplemento
    objCupomAux.sEndEntBairro = objCupom.sEndEntBairro
    objCupomAux.sEndEntCidade = objCupom.sEndEntCidade
    objCupomAux.sEndEntUF = objCupom.sEndEntUF
    objCupomAux.lEndEntIBGECidade = objCupom.lEndEntIBGECidade
    objCupomAux.iDAVImpresso = objCupom.iDAVImpresso
    objCupomAux.sEndEntEmail = objCupom.sEndEntEmail
    objCupomAux.lCOOCupomOrigDAV = objCupom.lCOOCupomOrigDAV
        
    objCupomAux.IdExterno = objCupom.IdExterno
    objCupomAux.NumCaixa = objCupom.NumCaixa
    objCupomAux.StatusExterno = objCupom.StatusExterno
        
    Set objVendaAux.objCupomFiscal = objCupomAux
    
    For iIndice = 1 To objVenda.objCupomFiscal.colItens.Count
        Set objCupomItem = objVenda.objCupomFiscal.colItens.Item(iIndice)
        Set objCupomItemAux = New ClassItemCupomFiscal
        
        objCupomItemAux.dAliquotaICMS = objCupomItem.dAliquotaICMS
        objCupomItemAux.dPercDesc = objCupomItem.dPercDesc
        objCupomItemAux.dPrecoUnitario = objCupomItem.dPrecoUnitario
        objCupomItemAux.dQuantidade = objCupomItem.dQuantidade
        objCupomItemAux.dValorDesconto = objCupomItem.dValorDesconto
'        objCupomItemAux.icancel = objCupomItem.icancel
        objCupomItemAux.iCodCaixa = giCodCaixa 'objCupomItem.iCodCaixa
        objCupomItemAux.iFilialEmpresa = giFilialEmpresa 'objCupomItem.iFilialEmpresa
        objCupomItemAux.iItem = objCupomItem.iItem
        objCupomItemAux.iStatus = objCupomItem.iStatus
        objCupomItemAux.lNumIntCupom = objCupomItem.lNumIntCupom
        objCupomItemAux.lNumIntDoc = objCupomItem.lNumIntDoc
        objCupomItemAux.sProduto = objCupomItem.sProduto
        objCupomItemAux.sSituacaoTrib = objCupomItem.sSituacaoTrib
        objCupomItemAux.sUnidadeMed = objCupomItem.sUnidadeMed
        objCupomItemAux.sProdutoNomeRed = objCupomItem.sProdutoNomeRed
        
        objVendaAux.objCupomFiscal.colItens.Add objCupomItemAux
    Next
    
End Sub

Private Function Transforma_Cupom(colItens As Collection) As Long
    
Dim objItemCupom As ClassItemCupomFiscal
Dim lErro As Long
Dim bAchou As Boolean
Dim objProduto As ClassProduto
Dim sProduto As String
Dim lNum As Long
Dim lNumero As Long

On Error GoTo Erro_Transforma_Cupom
        
    Call Limpa_Tela_Venda
    
    For Each objItemCupom In colItens
        
'        If objItemCupom.iStatus = STATUS_ATIVO Or gobjVenda.objCupomFiscal.lNumeroDAV <> 0 Then
            ProdutoNomeRed.Text = objItemCupom.sProdutoNomeRed
            Quantidade.Text = objItemCupom.dQuantidade
            
            If objItemCupom.dValorDesconto <> 0 Then
                DescontoValor.Text = Format(objItemCupom.dValorDesconto / objItemCupom.dQuantidade, "standard")
            Else
                DescontoValor.Text = ""
            End If
            
            lErro = Adiciona_Cupom(1, objItemCupom)
            If lErro <> SUCESSO Then gError 210509
            
'        End If
                
'        If objItemCupom.iStatus = STATUS_CANCELADO And gobjVenda.objCupomFiscal.lNumeroDAV <> 0 Then
        If objItemCupom.iStatus = STATUS_CANCELADO Then
            'cancelar o Item anterior
            lErro = AFRAC_CancelarItem(CInt(objItemCupom.iItem))
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Item")
            If lErro <> SUCESSO Then gError 210517
            
            'Atualiza o subtotal
            Subtotal.Caption = Format(Subtotal.Caption - ((objItemCupom.dPrecoUnitario * objItemCupom.dQuantidade) - objItemCupom.dValorDesconto), "standard")
            
            If Grid.Width < 8000 Then
'                ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", " ***** ITEM " & objItemCupom.iItem & " CANCELADO *****")
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", "           ***** ITEM " & objItemCupom.iItem & " CANCELADO *****") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", "-" & Format(objItemCupom.dPrecoUnitario * objItemCupom.dQuantidade, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItemCupom.iItem
            Else
                Call Proxima_Linha_Grid
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "***** ITEM CANCELADO *****"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = "-" & Format(objItemCupom.dPrecoUnitario * objItemCupom.dQuantidade, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItemCupom.iItem)
            
            End If
            
            gobjVenda.objCupomFiscal.colItens.Item(gobjVenda.objCupomFiscal.colItens.Count).iStatus = STATUS_CANCELADO
            
        End If
                
    Next
    
    gobjVenda.objCupomFiscal.dValorProdutos = StrParaDbl(Subtotal.Caption)
    'tem que incluir acrescimos e descontos lancados na tela de pagto
    gobjVenda.objCupomFiscal.dValorTotal = Arredonda_Moeda(gobjVenda.objCupomFiscal.dValorProdutos + gobjVenda.objCupomFiscal.dValorAcrescimo - gobjVenda.objCupomFiscal.dValorDesconto - gobjVenda.objCupomFiscal.dValorDesconto1)
    
    Call Venda_AjustaTrib
    
    If gobjVenda.objCupomFiscal.iVendedor <> 0 Then
        CodVendedor.Text = gobjVenda.objCupomFiscal.iVendedor
    Else
        CodVendedor.Text = ""
    End If
    
    Call CodVendedor_Validate(False)
    
'    CodCliente.Text = gobjVenda.objCupomFiscal.sCPFCGC
    NomeCliente.Text = gobjVenda.objCupomFiscal.sNomeCliente
    Email.Text = gobjVenda.objCupomFiscal.sEndEntEmail
    CGC.Text = gobjVenda.objCupomFiscal.sCPFCGC1
'    Endereco.Text = gobjVenda.objCupomFiscal.sEndereco
        
    Call Limpa_Cupom_Tela
    
    If AFRAC_ImpressoraCFe(giCodModeloECF) Then Call CF_ECF("Grava_Venda_Atual", gobjVenda)
    
    Transforma_Cupom = SUCESSO
    
    Exit Function

Erro_Transforma_Cupom:
    
    Transforma_Cupom = gErr
    
    Select Case gErr
                  
        Case 99818, 99884, 99912, 210509, 210517
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175717)

    End Select
    
    Exit Function
    
End Function

'Private Sub OptionOrcamento_Click()
'
'Dim vbMsgRes As VbMsgBoxResult
'Dim iCodigo As Integer
'Dim lErro As Long
'Dim lNum As Long
'Dim iCodGerente As Integer
'Dim objOperador As New ClassOperador
'
'On Error GoTo Erro_OptionOrcamento_Click
'
'    lNum = Retorna_Count_ItensCupom
'
'    'Se tiver um Cupom em andamento
'    If gobjVenda.objCupomfiscal.lNumero <> 0 And gobjVenda.iTipo <> OPTION_ORCAMENTO Then
'
'        'Envia aviso perguntando se deseja cancelar o cupom em andamento
'        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM)
'
'        If vbMsgRes = vbYes Then
'            'Se for Necessário a autorização do Gerente para abertura do Caixa
'            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
'
'                'Chama a Tela de Senha
'                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
'
'                'Sai de Função se a Tela de Login não Retornar ok
'                If giRetornoTela <> vbOK Then gError 102501
'
'                'Se Operador for Gerente
'                iCodGerente = objOperador.iCodigo
'
'            End If
'
'            Exibe.Caption = "CANCELADO CUPOM " & gobjVenda.objCupomfiscal.lNumero
'            Exibe1.Caption = ""
'            DoEvents
'            'Cancelar o Cupom de Venda
'            lErro = AFRAC_CancelarCupom()
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancela Cupom")
'            If lErro <> SUCESSO Then gError 99748
'
'            'Realiza as operações necessárias para gravar
'            lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
'            If lErro <> SUCESSO Then gError 204545
'
'            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
'            If lErro <> SUCESSO Then gError 204546
'
'            Exibe1.Caption = ""
'            Exibe.Caption = "TRANSF.: VENDA EM ORÇAMENTO."
'        Else
''            Exibe.Caption = "PRÓXIMO CLIENTE..."
''            Exibe1.Caption = ""
'            'Seleciona o cupom
'            giNaoTrataCF = 1
'            OptionCF.Value = True
'            Exit Sub
'        End If
'    End If
'
'    'Quando eu clico no Orçamento reativa o número do orcamento(campo)
'    Orcamento.Enabled = True
'    LabelOrcamento.Enabled = True
'    BotaoProxNum.Enabled = True
'    gobjVenda.objCupomfiscal.lNumero = 0
'    gobjVenda.iTipo = OPTION_ORCAMENTO
'
'    Exit Sub
'
'Erro_OptionOrcamento_Click:
'
'    Select Case gErr
'
'        Case 99748, 102501
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175718)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub BotaoPreco_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iTipo As Integer
        
    'Chama tela de Preco
    Call Chama_TelaECF_Modal("Preco", objProduto)
    
    If Not (objProduto Is Nothing) Then
        ProdutoNomeRed.Text = objProduto.sCodigo
        Call ProdutoNomeRed_Validate(False)
    End If
     
End Sub

Private Function Adiciona_Cupom(ByVal iTransfOrcCF As Integer, Optional ByVal objItemCupomOrc As ClassItemCupomFiscal) As Long
    
Dim lErro As Long
Dim bAchou As Boolean
Dim objProduto As ClassProduto
Dim objItemCupom As New ClassItemCupomFiscal
Dim sProduto As String
Dim lNum As Long
Dim lNumero As Long
Dim objAliquota As ClassAliquotaICMS
Dim objCliente As ClassCliente
Dim sCPF As String
Dim sRet As String
Dim lErro1 As Long
Dim sPrecoItem As String
Dim sAliquota As String
Dim sPeso As String
Dim sPrecoKilo As String
Dim sPrecoTotal As String
Dim sProduto1 As String
Dim objAliquota1 As ClassAliquotaICMS
Dim iTotalizador As Integer
Dim objVenda As New ClassVenda
Dim sSubtotal As String
Dim sPrecoItem1 As String, bDescontoValor As Boolean

On Error GoTo Erro_Adiciona_Cupom
    
    If gobjVenda.objCupomFiscal.bEditavel = False Then gError 201583
    
    If gobjVenda.iTipo = 0 Then gError 214767
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 214333
    
    'se se tratar de DAV e ja tiver sido impresso ==> nao pode alterar
'    If gobjVenda.iTipo = OPTION_DAV And gobjVenda.objCupomFiscal.lNumero <> 0 Then gError 210508
    
    
    If gobjVenda.iTipo = OPTION_DAV Then
    
        objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
    
        lErro = CF_ECF("OrcamentoECF_Le", objVenda)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210896
    
        'se o orcamento é um DAV e ja foi impresso ==> nao pode alterar
        If objVenda.iTipo = OPTION_DAV And objVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210508
    
    End If
    
    giItem = gobjVenda.objCupomFiscal.colItens.Count + 1
    
    sProduto1 = ProdutoNomeRed.Text
    
    Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
            
    'caso o produto não seja encontrado
    If objProduto Is Nothing Then
        ProdutoNomeRed.Text = ""
        gError 99606
    End If
    
    'verifica se a figura foi preenchida
    If objProduto.sFigura <> "" Then
        'verifica se o arquivo é do tipo imagem
        sRet = Dir(objProduto.sFigura, vbNormal)
        If sRet <> "" Then
            If GetAttr(objProduto.sFigura) = vbArchive Then
                'coloca a figura na tela
                Figura.Picture = LoadPicture(objProduto.sFigura)
            End If
        Else
            gError 99607
        End If
    Else
        Figura.Picture = LoadPicture
    End If
    
    If iTransfOrcCF = 0 Then
    
        'se esta tentando acrescentar um item num dav ou prevenda ao emitir o cupom fiscal relativo ==> erro
        If gobjVenda.iTipo = OPTION_CF And gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then gError 214301
    
        If objProduto.iUsaBalanca = USA_BALANCA And Len(Trim(gsBalancaPorta)) > 0 And giBalancaModelo > 0 Then
            lErro = AFRAC_Le_Balanca(gsBalancaNome, gsBalancaPorta, giBalancaModelo, sPeso, sPrecoKilo, sPrecoTotal)
            lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Leitura Balança")
            If lErro1 <> SUCESSO Or lErro <> SUCESSO Then gError 133754
            
            Quantidade.Text = sPeso
        End If
    End If
    
    
    For Each objAliquota In gcolAliquotasTotal
        If objAliquota.sSigla = objProduto.sICMSAliquota Then
            objItemCupom.dAliquotaICMS = objAliquota.dAliquota
            Exit For
        End If
    Next
    
    
    objItemCupom.dQuantidade = StrParaDbl(Quantidade.Text)
    If iTransfOrcCF = 0 Or objItemCupomOrc Is Nothing Then
        objItemCupom.dPrecoUnitario = objProduto.dPrecoLoja
    Else
        objItemCupom.dPrecoUnitario = objItemCupomOrc.dPrecoUnitario
    End If
    objItemCupom.sUnidadeMed = objProduto.sSiglaUMVenda
    objItemCupom.sSituacaoTrib = objProduto.sSituacaoTribECF
    objItemCupom.sProduto = objProduto.sCodigo
    objItemCupom.sProdutoNomeRed = objProduto.sNomeReduzido
    objItemCupom.sProdutoDescricao = objProduto.sDescricao
    
    objItemCupom.iItem = giItem
    objItemCupom.iStatus = STATUS_ATIVO
    
    lNum = Retorna_Count_ItensCupom
    
    If gobjVenda.iTipo = OPTION_CF And lNum = 0 Then
        sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
        lErro = CF_ECF("Abre_Cupom", gobjVenda)
        
        'se saiu porque deseja preencher o cpf na nota
        If lErro = 214505 Then
            ProdutoNomeRed.Text = ""
            CGC.SetFocus
        End If
        
        If lErro <> SUCESSO Then gError 99818
    End If

    sPrecoItem = Format(StrParaDbl(Format(Quantidade.Text, "0.000")) * StrParaDbl(Format(objItemCupom.dPrecoUnitario, "standard")), "Standard")

    bDescontoValor = False
    If Len(Trim(DescontoValor.Text)) > 0 And objItemCupom.dPrecoUnitario > StrParaDbl(DescontoValor.Text) Then
    
        objItemCupom.dValorDesconto = Arredonda_Moeda(objItemCupom.dQuantidade * StrParaDbl(DescontoValor.Text), 2)
        bDescontoValor = True
    
    Else
    
        If objProduto.dDescontoValor > 0 Then
            objItemCupom.dValorDesconto = objProduto.dDescontoValor
            bDescontoValor = True
        ElseIf objProduto.dPercentMenosReceb > 0 Then
            objItemCupom.dValorDesconto = Fix(StrParaDbl(sPrecoItem) * objProduto.dPercentMenosReceb) / 100
        End If

    End If
        
    If objItemCupom.dAliquotaICMS > 0 Then
        If objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
            
            sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL & Format(objItemCupom.dAliquotaICMS * 10000, "0000")
            
'                For Each objAliquota1 In gobjLojaECF.colAliquotaICMS
'                    If objAliquota1.iISS = 1 And objItemCupom.dAliquotaICMS = objAliquota1.dAliquota Then
'                        Exit For
'                    End If
'                    iTotalizador = iTotalizador + 1
'                Next

'                objItemCupom.sSituacaoTrib = Format(iTotalizador, "00") & sAliquota
            objItemCupom.sSituacaoTrib = sAliquota
            
        Else
            
'                For Each objAliquota1 In gobjLojaECF.colAliquotaICMS
'                    If objAliquota1.iISS = 0 And objItemCupom.dAliquotaICMS = objAliquota1.dAliquota Then
'                        Exit For
'                    End If
'                    iTotalizador = iTotalizador + 1
'                Next
            
            sAliquota = Format(objItemCupom.dAliquotaICMS * 10000, "0000")
'                objItemCupom.sSituacaoTrib = Format(iTotalizador, "00") & TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL & sAliquota
            objItemCupom.sSituacaoTrib = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL & sAliquota
            
        End If
    Else
        sAliquota = left(objProduto.sSituacaoTribECF, 1)
        objItemCupom.sSituacaoTrib = sAliquota & "1"
    End If
        
    If gobjVenda.iTipo = OPTION_CF And Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
        
        lErro = AFRAC_VenderItem(objProduto.sCodigo, objProduto.sDescricao, StrParaDbl(Format(Quantidade.Text, "#0.000")), Format(objItemCupom.dPrecoUnitario, "standard"), 1, 1, objItemCupom.dValorDesconto, StrParaDbl(Format(StrParaDbl(Format(Quantidade.Text, "0.000")) * StrParaDbl(Format(objItemCupom.dPrecoUnitario, "standard")), "Standard")), sAliquota, objProduto.sSiglaUMVenda, False)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Vender Item")
        If lErro <> SUCESSO Then gError 99913
        
        lErro = CF_ECF("Requisito_XXII_AtualizaGT")
        If lErro <> SUCESSO Then gError 210424
        
        lErro = AFRAC_LerInformacaoImpressora("062", sPrecoItem1)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
        If lErro <> SUCESSO Then gError 214138
        
        sPrecoItem1 = Format(StrParaDbl(sPrecoItem1) / 100, "standard")
       
        lErro = AFRAC_LerInformacaoImpressora("001", sSubtotal)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
        If lErro <> SUCESSO Then gError 214139
       
        Subtotal.Caption = Format(StrParaDbl(sSubtotal) / 100, "standard")

    Else
    
        PrecoItem.Caption = sPrecoItem
        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - objItemCupom.dValorDesconto), "standard")
       
    End If

    
    'Prenche a col de itens do cupom com os dados do mesmo
    PrecoUnitario.Caption = Format(objItemCupom.dPrecoUnitario, "standard")
    ProdutoNomeRed.Text = objProduto.sNomeReduzido
    PrecoItem.Caption = sPrecoItem

    


    'Joga na col
    gobjVenda.objCupomFiscal.colItens.Add objItemCupom
    
    'Joga no cupom o item
    
    If Grid.Width < 8000 Then

        ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItemCupom.iItem) & "     " & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", objProduto.sCodigo) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", objProduto.sDescricao)
        ListCF.ItemData(ListCF.NewIndex) = objItemCupom.iItem
        ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", Format(Quantidade.Text, "#0.000")) & "  " & Formata_Campo(ALINHAMENTO_DIREITA, 4, " ", objProduto.sSiglaUMVenda) & " x " & Formata_Campo(ALINHAMENTO_DIREITA, 12, " ", Format(PrecoUnitario.Caption, "standard")) & objProduto.sSituacaoTribECF & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItemCupom.dAliquotaICMS * 100, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", Format(PrecoItem.Caption, "standard"))
        ListCF.ItemData(ListCF.NewIndex) = objItemCupom.iItem
        
    Else
    
        Call Proxima_Linha_Grid

        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItemCupom.iItem)
        'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_CODIGO) = objProduto.sCodigo
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = objProduto.sCodigo & " - " & objProduto.sDescricao
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_QUANTIDADE) = Format(Quantidade.Text, "#0.000") & " " & objProduto.sSiglaUMVenda
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_X) = "x"
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(PrecoUnitario.Caption, "standard")
        'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = objProduto.sSituacaoTribECF & Format(objItemCupom.dAliquotaICMS * 100, "fixed") & "%"
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(PrecoItem.Caption, "standard")
    
    End If
    
    
    'se existir desconto sobre o item...
    If objItemCupom.dValorDesconto <> 0 Then
    
        'se o desconto for em percentual
        If Not bDescontoValor Then
        
            If Grid.Width < 8000 Then
            
    '        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(objProduto.dPercentMenosReceb, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", "-" & Format(Fix(StrParaDbl(PrecoItem.Caption) * (objProduto.dPercentMenosReceb)) / 100, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItemCupom.dValorDesconto, "standard"))
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(objProduto.dPercentMenosReceb, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", "-" & Format(objItemCupom.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItemCupom.dValorDesconto, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItemCupom.iItem
    '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - (StrParaDbl(PrecoItem.Caption)), "standard")
    '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - (StrParaDbl(PrecoItem.Caption) * objProduto.dPercentMenosReceb / 100)), "standard")
    
            Else
            
                Call Proxima_Linha_Grid
                
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "DESCONTO: " & Format(objProduto.dPercentMenosReceb, "fixed") & "% = - R$ " & Format(objItemCupom.dValorDesconto, "standard")
                'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(objProduto.dPercentMenosReceb, "fixed") & "%"
                'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = "-" & Format(objItemCupom.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(StrParaDbl(PrecoItem.Caption) - objItemCupom.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItemCupom.iItem)
    
            End If
            
    '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - objItemCupom.dValorDesconto, "standard")
            Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption), "standard")
            
            
        'desconto em valor
        Else
        
            If Grid.Width < 8000 Then
        
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 21, " ", "-" & Format(objItemCupom.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItemCupom.dValorDesconto, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItemCupom.iItem
        '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - (StrParaDbl(PrecoItem.Caption)), "standard")
        '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - objProduto.dDescontoValor), "standard")
        
            Else
            
                Call Proxima_Linha_Grid
                
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "DESCONTO: - R$ " & Format(objItemCupom.dValorDesconto, "standard")
                'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = "-" & Format(objItemCupom.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(StrParaDbl(PrecoItem.Caption) - objItemCupom.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItemCupom.iItem)
            
            End If
            
    '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - objItemCupom.dValorDesconto, "standard")
            Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption), "standard")
            
        
        End If
    
    End If
    
    'Para rolar automaticamente a barra de rolagem
    If Grid.Width < 8000 Then
        ListCF.ListIndex = ListCF.NewIndex
    End If
    
    Exibe.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItemCupom.iItem) & "   " & Formata_Campo(ALINHAMENTO_DIREITA, 20, " ", objProduto.sNomeReduzido)
    
    If objItemCupom.dValorDesconto > 0 Then
        Exibe1.Caption = Format(Quantidade.Text, "0.000") & " x " & Format(PrecoUnitario.Caption, "standard") & " - " & Format(objItemCupom.dValorDesconto, "standard") & " = " & Format(StrParaDbl(PrecoItem.Caption) - objItemCupom.dValorDesconto, "standard")
    Else
        Exibe1.Caption = Format(Quantidade.Text, "0.000") & " x " & Format(PrecoUnitario.Caption, "standard") & " = " & Format(PrecoItem.Caption, "standard")
    End If
    
    Set objItemCupom.objTributacaoDocItem = New ClassTributacaoDocItem
    Call objItemCupom.objTributacaoDocItem.Copia(objProduto.objTributacaoDocItem)
    'ajusta de acordo com o que efetivamente foi vendido
    If iTransfOrcCF = 0 Then Call ItemCupom_AjustaTrib(objItemCupom)
    
    Call Limpa_Cupom_Tela
        
    DoEvents
        
    If gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) Then Call CF_ECF("Grava_Venda_Atual", gobjVenda)
    
    Adiciona_Cupom = SUCESSO
    
    Exit Function

Erro_Adiciona_Cupom:
    
    Adiciona_Cupom = gErr
    
    Select Case gErr
                
        Case 99606
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_SEM_PRECO, gErr, sProduto)
                
        Case 99607
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FIGURA_INVALIDO, gErr, objProduto.sFigura)
                
        Case 99740
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_INVALIDO, gErr, objProduto.sFigura)
                 
        Case 99818, 99913, 210896, 214138, 214139, 214333
        
        Case 214767
            Call Rotina_ErroECF(vbOKOnly, ERRO_TIPOCF_NAO_ESCOLHIDO, gErr)
        
        Case 210508
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
        
        Case 214301
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PERMITE_INCLUSAO_ITENS, gErr)
        
        Case 201583
            Call Rotina_ErroECF(vbOKOnly, ERRO_EM_PROCESSAMENTO_SEFAZ, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175719)

    End Select
    
    Exit Function
    
End Function

Private Sub Limpa_Tela_Venda()


Dim lErro As Long
Dim objOrcamento As Object

On Error GoTo Erro_Limpa_Tela_Venda

    Call Limpa_Cupom_Tela
    
    CGC.Text = ""
'    CodCliente.Text = ""
    NomeCliente.Text = ""
    Email.Text = ""
'    Endereco.Text = ""

    If giCodModeloECF = IMPRESSORA_NFCE Then
    
        ImprimirNFCE.Value = IIf(gobjVenda.objNFeInfo.iNFCeImprimir = 0, vbUnchecked, vbChecked)
        
        If ImprimirNFCE.Visible Then gobjVenda.objCupomFiscal.bNFCEImprimir = IIf(ImprimirNFCE.Value = vbChecked, True, False)
        
    End If
        
    Call Limpa_Tela_Venda_1

    Call Menu_Click(0)

           
    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then
    
'        If giPrevenda = 1 And giUsaImpressoraFiscal = 0 Then
'            OptionPreVenda.Value = False
'            OptionPreVenda.Value = True
'        ElseIf giDAV = 1 Then
'            OptionDAV.Value = False
'            OptionDAV.Value = True
'        End If
            
        
        Set objOrcamento = Orcamento
        
        lErro = CF_ECF("Retorna_Prox_Orcamento", objOrcamento, gobjVenda)
        If lErro <> SUCESSO Then gError 210412
    Else
'        OptionCF.Value = False
'        OptionCF.Value = True
    End If
    

    If OptionPreVenda.Value Then gobjVenda.iTipo = OPTION_PREVENDA
    
    If OptionDAV.Value Then gobjVenda.iTipo = OPTION_DAV

    If OptionCF.Value Then gobjVenda.iTipo = OPTION_CF

    Exit Sub

Erro_Limpa_Tela_Venda:
    
    Select Case gErr
                
        Case 210412
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210413)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Limpa_Cupom_Tela()

    PrecoUnitario.Caption = Format(0, "standard")
    PrecoItem.Caption = Format(0, "standard")
    ProdutoNomeRed.Text = ""
    Quantidade.Text = 1
    DescontoValor.Text = ""
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim bAchou As Boolean, ShiftKey As Integer
    
    ShiftKey = Shift And 7
    
    If ShiftKey = 0 Then

        Select Case KeyCode
        
            Case 13
                KeyCode = vbKeyTab
        
            Case KEYCODE_BROWSER
                If Me.ActiveControl Is CodVendedor Then
                    Call LabelCodVendedor_Click
'                ElseIf Me.ActiveControl Is CodCliente Then
'                    Call LabelCliente_Click
                ElseIf Me.ActiveControl Is Orcamento Then
                    Call LabelOrcamento_Click
                End If
        
            Case KEYCODE_PROXIMO_NUMERO
                If Not TrocaFoco(Me, BotaoProxNum) Then Exit Sub
                Call BotaoProxNum_Click
            
            Case vbKeyF4
                If Not TrocaFoco(Me, BotaoCancelaItemAtual) Then Exit Sub
                Call BotaoCancelaItemAtual_Click
        
            Case vbKeyF5
                If Not TrocaFoco(Me, BotaoPreco) Then Exit Sub
                Call BotaoPreco_Click
        
            Case vbKeyF6
                If Not TrocaFoco(Me, BotaoCancelaItem) Then Exit Sub
                Call BotaoCancelaItem_Click
        
            Case vbKeyF7
                If Not TrocaFoco(Me, BotaoSuspender) Then Exit Sub
                Call BotaoSuspender_Click
        
            Case vbKeyF8
                If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
                Call BotaoFechar_Click
        
            Case vbKeyF9
                If Not TrocaFoco(Me, BotaoProdutos) Then Exit Sub
                Call BotaoProdutos_Click
        
            Case vbKeyF10
                If Not TrocaFoco(Me, BotaoAbrirGaveta) Then Exit Sub
                Call BotaoAbrirGaveta_Click
        
            Case vbKeyF11
                If Not TrocaFoco(Me, BotaoPagamento) Then Exit Sub
                Call BotaoPagamento_Click
                
            Case vbKeyF12
                Call LabelOrcamento_Click
                
            Case vbKeyEscape
                If Not TrocaFoco(Me, BotaoCancelaCupom) Then Exit Sub
                Call BotaoCancelaCupom_Click
        
        End Select
                
    ElseIf ShiftKey = 2 Then
    
        Call Acha_Produto_Tecla(KeyCode, bAchou)
        If bAchou Then
        
            KeyCode = 0
            Call ProdutoNomeRed_LostFocus
        
        Else
        
            If KeyCode = vbKeyF2 Or KeyCode = vbKeyP Then
                Call BotaoReimprimir_Click
            End If

        End If
        
    End If
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    If LeitoraCodBarras.PortOpen = True Then LeitoraCodBarras.PortOpen = False
    
    'Libera a referência da tela
    Set gobjVenda = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoOrcamento = Nothing
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim lNum As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer

On Error GoTo Erro_Form_QueryUnload

    If (ListCF.ListIndex <> -1 Or giUltimaLinhaGrid > 1) And gobjVenda.objCupomFiscal.bEditavel Then
    
        Timer1.Enabled = False
        
'        lErro = CF_ECF("Requisito_XXII")
'        If lErro <> SUCESSO Then gError 210028
        
        If gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
            lErro = CancelaCupomEletronico(BOTAO_FECHAR_CUPOM_CFE)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Else
        
            lNum = Retorna_Count_ItensCupom
            
            'Se tiver uma venda acontecendo
            If gobjVenda.objCupomFiscal.lNumero <> 0 And OptionCF.Value Then
                'Envia aviso perguntando se deseja cancelar a venda
                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_VENDA)
        
                If vbMsgRes = vbYes Then
                    'Se for Necessário a autorização do Gerente para abertura do Caixa
                    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
                
                        'Chama a Tela de Senha
                        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
                
                        'Sai de Função se a Tela de Login não Retornar ok
                        If giRetornoTela <> vbOK Then gError 102507
                        
                        'Se Operador for Gerente
                        iCodGerente = objOperador.iCodigo
                
                    End If
        
        
                    'Cancelar o Cupom de Venda
                    lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                    If lErro <> SUCESSO Then gError 99617
                    
                    Call Move_Dados_Memoria_1
                    
                    'Realiza as operações necessárias para gravar
                    lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                    If lErro <> SUCESSO Then gError 204547
                    
                    lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                    If lErro <> SUCESSO Then gError 204548
                    
                    
                Else
                    Cancel = True
                End If
                
            End If
        
            'Se tiver um orçamento na tela
            If lNum > 0 And (OptionDAV.Value Or OptionPreVenda.Value) Then
                
                'Envia aviso que o orçamento será cancelado
                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_ORCAMENTO)
        
                If vbMsgRes = vbYes Then
                    'Se for Necessário a autorização do Gerente para abertura do Caixa
                    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
                
                        'Chama a Tela de Senha
                        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
                
                        'Sai de Função se a Tela de Login não Retornar ok
                        If giRetornoTela <> vbOK Then gError 102506
                        
                        'Se Operador for Gerente
                        iCodGerente = objOperador.iCodigo
                
                    End If
                    
                Else
                   Cancel = True
                End If
                
                    
            End If
        
        End If
        
        If Cancel = True Then Timer1.Enabled = True
    
    End If
    
    Exit Sub

Erro_Form_QueryUnload:

    Select Case gErr
                
        Case 99617, 102507, 204547, 204548, 210028, ERRO_SEM_MENSAGEM
            Cancel = True
            Timer1.Enabled = True
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175720)

    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
Dim objOperador As ClassOperador
Dim sOper As String

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    For Each objOperador In gcolOperadores
        If objOperador.iCodigo = giCodOperador Then sOper = objOperador.sNome
    Next
    
    Caption = Formata_Campo(ALINHAMENTO_DIREITA, 20, " ", "Venda") & "Filial : " & giFilialEmpresa & "    Caixa : " & giCodCaixa & "    Operador : " & sOper
    Call Form_Load

    
End Function

Public Function Name() As String

    Name = "VendaP"
    
End Function

Public Function objParent() As Object

    Set objParent = Parent
    
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

Private Sub LeitoraCodBarras_OnComm()

Dim sInput As String
Dim sInput1 As String
Dim iInput As Integer
Dim i As Long
Dim j As Long

On Error GoTo Erro_LeitoraCodBarras_OnComm

    Select Case LeitoraCodBarras.CommEvent
    
        ' Handle each event or error by placing
        ' code below each case statement
        
        ' Errors
    
        Case comEventBreak   ' A Break was received.
            
        Case comEventFrame   ' Framing Error
            MsgBox "Erro de Frame"
            
        Case comEventOverrun   ' Data Lost.
            MsgBox "Dados Perdidos"
      
        Case comEventRxOver   ' Receive buffer overflow.
            MsgBox "Buffer Overflow"
      
        Case comEventRxParity   ' Parity Error.
            MsgBox "Erro de Paridade"
      
        Case comEventTxFull   ' Transmit buffer full.
        
        Case comEventDCB   ' Unexpected error retrieving DCB]
            MsgBox "Erro de DCB"

        ' Events
        
        Case comEvCD   ' Change in the CD line.
      
        Case comEvCTS   ' Change in the CTS line.
      
        Case comEvDSR   ' Change in the DSR line.
      
        Case comEvRing   ' Change in the Ring Indicator.
            
        Case comEvReceive   ' Received RThreshold # of chars.
         
         
                     
'            For i = 1 To 1000000
'               j = j + 1
'            Next

            Call Sleep(1000)
         
            LeitoraCodBarras.InputLen = 0
            sInput = LeitoraCodBarras.Input
            ProdutoNomeRed.Text = left(sInput, Len(sInput) - 1)
            LeitoraCodBarras.InBufferCount = 0
      
            Call ProdutoNomeRed_Validate(False)
      
      
        Case comEvSend   ' There are SThreshold number of characters in the transmit buffer.
      
        Case comEvEOF   ' An EOF charater was found in the input stream
            
    End Select

    Exit Sub
    
Erro_LeitoraCodBarras_OnComm:
    
    Select Case gErr
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 175695)

    End Select
    
    Exit Sub

End Sub


Public Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As ClassCliente
Dim lIndice As Long

On Error GoTo Erro_CGC_Validate
    
'    'Se CGC/CPF não foi preenchido -- Exit Sub
'    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
'
'    Select Case Len(Trim(CGC.Text))
'
'        Case STRING_CPF 'CPF
'
'            'Critica Cpf
'            lErro = Cpf_Critica(CGC.Text)
'            If lErro <> SUCESSO Then gError 199469
'
'            'Formata e coloca na Tela
'            CGC.Format = "000\.000\.000-00; ; ; "
'            CGC.Text = CGC.Text
'
'        Case STRING_CGC 'CGC
'
'            'Critica CGC
'            lErro = Cgc_Critica(CGC.Text)
'            If lErro <> SUCESSO Then gError 199470
'
'            'Formata e Coloca na Tela
'            CGC.Format = "00\.000\.000\/0000-00; ; ; "
'            CGC.Text = CGC.Text
'
'        Case Else
'
'            gError 199471
'
'    End Select
'
'    gobjVenda.objCupomFiscal.sCPFCGC1 = CGC.FormattedText
'    gobjVenda.objCupomFiscal.sCPFCGC = CGC.ClipText
'
'    Exit Sub

    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC.Text)) = 0 Then
        gobjVenda.objCupomFiscal.sCPFCGC1 = ""
        gobjVenda.objCupomFiscal.sCPFCGC = ""
        Exit Sub
    End If
    
    Set objCliente = gobjClienteCPF.Busca(CGC.ClipText, lIndice)
    
    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 199469
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 199470
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            gError 199471

    End Select

    If Not objCliente Is Nothing Then
        
        NomeCliente.Text = objCliente.sNomeReduzido
        gobjVenda.objCupomFiscal.sNomeCliente = objCliente.sNomeReduzido

    End If

    gobjVenda.objCupomFiscal.sCPFCGC1 = CGC.FormattedText
    gobjVenda.objCupomFiscal.sCPFCGC = CGC.ClipText
    
    Exit Sub


Erro_CGC_Validate:

    Cancel = True

    Select Case gErr

        Case 199469, 199470

        Case 199471
            Call Rotina_Erro(vbOKOnly, ERRO_TAMANHO_CGC_CPF, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 199472)

    End Select

    Exit Sub

End Sub

Private Sub NomeCliente_Validate(Cancel As Boolean)
    gobjVenda.objCupomFiscal.sNomeCliente = NomeCliente.Text
End Sub

Private Sub Email_Validate(Cancel As Boolean)
    
Dim sEmail As String

On Error GoTo Erro_Email_Validate

    sEmail = Trim(Email.Text)

    If Len(sEmail) <> 0 Then
        If Not ValidEmail(sEmail) Then gError 201581
    End If
    
    gobjVenda.objCupomFiscal.sEndEntEmail = sEmail
    
    If ImprimirNFCE.Visible And gobjVenda.objNFeInfo.iNFCeImprimir <> 0 Then
        
        If Len(sEmail) <> 0 Then
            ImprimirNFCE.Value = vbUnchecked
        Else
            ImprimirNFCE.Value = vbChecked
        End If
    
    End If
    
    Exit Sub

Erro_Email_Validate:

    Cancel = True

    Select Case gErr

        Case 201581
            Call Rotina_ErroECF(vbOKOnly, ERRO_EMAIL_INVALIDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 201580)

    End Select

    Exit Sub

End Sub

Private Sub OptionDAV_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_OptionDAV_Click
    
    If gobjVenda.objCupomFiscal.bEditavel = False Then gError 201583
    
    If gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = CancelaCupomEletronico(BOTAO_DAV_CANCELA_CUPOM_CFE)
        If lErro <> SUCESSO Then
        
            'Seleciona o cupom
            OptionCF.Value = True
        
            Exit Sub
        
        End If
    
    Else
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210029
        
        'Se tiver um Cupom em andamento
        If gobjVenda.objCupomFiscal.lNumero <> 0 And gobjVenda.iTipo <> OPTION_DAV And gobjVenda.iTipo <> OPTION_PREVENDA Then
            'Envia aviso perguntando se deseja cancelar o cupom em andamento
            vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM)
        
            If vbMsgRes = vbYes Then
                'Se for Necessário a autorização do Gerente para abertura do Caixa
                If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
            
                    'Chama a Tela de Senha
                    Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
            
                    'Sai de Função se a Tela de Login não Retornar ok
                    If giRetornoTela <> vbOK Then gError 204717
                    
                    'Se Operador for Gerente
                    iCodGerente = objOperador.iCodigo
                
                End If
                
                Exibe.Caption = "CANCELADO CUPOM CCF = " & gobjVenda.objCupomFiscal.lCCF
                Exibe1.Caption = ""
                DoEvents
                'Cancelar o Cupom de Venda
                lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancela Cupom")
                If lErro <> SUCESSO Then gError 204718
                
                Call Move_Dados_Memoria_1
    
                'Realiza as operações necessárias para gravar
                lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204719
                
                lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204720
                
                Exibe1.Caption = ""
                Exibe.Caption = "TRANSF.: VENDA EM ORÇAMENTO."
            Else
                Exibe.Caption = "PRÓXIMO CLIENTE..."
                Exibe1.Caption = ""
                'Seleciona o cupom
                OptionCF.Value = True
                Exit Sub
            End If
        End If
    
    End If
    
    'Quando eu clico no Orçamento reativa o número do orcamento(campo)
'    Orcamento.Enabled = True
'    LabelOrcamento.Enabled = True
    BotaoProxNum.Enabled = True
    gobjVenda.objCupomFiscal.lNumero = 0
    gobjVenda.iTipo = OPTION_DAV
    Orcamento.Text = ""
    
    BotaoAbrirGaveta.Caption = "(F10)   Grava Orçamento"
    BotaoCancelaCupom.Caption = "(Esc)   Imprime Orçamento"
    
    Exit Sub

Erro_OptionDAV_Click:

    Select Case gErr
                
        Case 201583
            Call Rotina_ErroECF(vbOKOnly, ERRO_EM_PROCESSAMENTO_SEFAZ, gErr)
                
        Case 204717 To 204720, 210029
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204721)

    End Select
    
    Exit Sub
        
End Sub

Private Sub OptionPreVenda_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_OptionPreVenda_Click
    
    If gobjVenda.objCupomFiscal.bEditavel = False Then gError 201583
    
    If gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = CancelaCupomEletronico(BOTAO_PREVENDA_CANCELA_CUPOM_CFE)
        If lErro <> SUCESSO Then
        
            'Seleciona o cupom
            OptionCF.Value = True
        
            Exit Sub
        
        End If
    
    Else
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210030
    
        'Se tiver um Cupom em andamento
        If gobjVenda.objCupomFiscal.lNumero <> 0 And gobjVenda.iTipo <> OPTION_DAV And gobjVenda.iTipo <> OPTION_PREVENDA Then
            'Envia aviso perguntando se deseja cancelar o cupom em andamento
            vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM)
        
            If vbMsgRes = vbYes Then
                'Se for Necessário a autorização do Gerente para abertura do Caixa
                If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
            
                    'Chama a Tela de Senha
                    Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
            
                    'Sai de Função se a Tela de Login não Retornar ok
                    If giRetornoTela <> vbOK Then gError 204722
                    
                    'Se Operador for Gerente
                    iCodGerente = objOperador.iCodigo
                
                End If
                
                Exibe.Caption = "CANCELADO CUPOM CCF = " & gobjVenda.objCupomFiscal.lCCF
                Exibe1.Caption = ""
                DoEvents
                'Cancelar o Cupom de Venda
                lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancela Cupom")
                If lErro <> SUCESSO Then gError 204723
                
                Call Move_Dados_Memoria_1
                
                'Realiza as operações necessárias para gravar
                lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204724
                
                lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204725
                
                Exibe1.Caption = ""
                Exibe.Caption = "TRANSF.: VENDA EM ORÇAMENTO."
            Else
                Exibe.Caption = "PRÓXIMO CLIENTE..."
                Exibe1.Caption = ""
                'Seleciona o cupom
                OptionCF.Value = True
                Exit Sub
            End If
        End If
    
    End If
    
    'Quando eu clico no Orçamento reativa o número do orcamento(campo)
'    Orcamento.Enabled = True
'    LabelOrcamento.Enabled = True
    BotaoProxNum.Enabled = True
    gobjVenda.objCupomFiscal.lNumero = 0
    gobjVenda.iTipo = OPTION_PREVENDA
    Orcamento.Text = ""
    
    BotaoAbrirGaveta.Caption = "(F10)   Grava Orçamento"
    BotaoCancelaCupom.Caption = "(Esc)   Limpa Orçamento"
    
    Exit Sub

Erro_OptionPreVenda_Click:

    Select Case gErr
                
        Case 204722 To 204725, 210030
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204726)

    End Select
    
    Exit Sub
        
End Sub

'Private Sub Endereco_Validate(Cancel As Boolean)
'    gobjVenda.objCupomFiscal.sEndereco = Endereco.Text
'End Sub

Private Sub Move_Dados_Memoria_1()

    gobjVenda.objCupomFiscal.dValorProdutos = StrParaDbl(Subtotal.Caption)
    'tem que incluir acrescimos e descontos lancados na tela de pagto
    gobjVenda.objCupomFiscal.dValorTotal = Arredonda_Moeda(gobjVenda.objCupomFiscal.dValorProdutos + gobjVenda.objCupomFiscal.dValorAcrescimo - gobjVenda.objCupomFiscal.dValorDesconto - gobjVenda.objCupomFiscal.dValorDesconto1)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iECF = giCodECF

End Sub

Private Sub ItemCupom_AjustaTrib(objItem As ClassItemCupomFiscal)
'Ajusta a tributacao de acordo com o que foi efetivamente vendido
'??? falta tratar ST

Dim objTributacaoDocItem As ClassTributacaoDocItem
Dim dValorLiquido As Double

    Set objTributacaoDocItem = objItem.objTributacaoDocItem

    dValorLiquido = Arredonda_Moeda(Arredonda_Moeda(objItem.dQuantidade * objItem.dPrecoUnitario) - Arredonda_Moeda(objItem.dValorDesconto) + objTributacaoDocItem.dValorFreteItem)
    
    'dados gerais
    objTributacaoDocItem.dQuantidade = objItem.dQuantidade
    objTributacaoDocItem.dQtdTrib = objItem.dQuantidade
    objTributacaoDocItem.dPrecoUnitario = Arredonda_Moeda(objItem.dPrecoUnitario)
    objTributacaoDocItem.dValorUnitTrib = Arredonda_Moeda(objItem.dPrecoUnitario)
    objTributacaoDocItem.dDescontoGrid = Arredonda_Moeda(objItem.dValorDesconto)
    
    If objTributacaoDocItem.dTotTrib <> 0 Then
        objTributacaoDocItem.dTotTrib = Arredonda_Moeda(objTributacaoDocItem.dTotTrib * dValorLiquido / 1000)
    End If
    
    'icms
    If objTributacaoDocItem.dICMSAliquota <> 0 Then
        objTributacaoDocItem.dICMSBase = dValorLiquido
        objTributacaoDocItem.dICMSValor = Arredonda_Moeda(objTributacaoDocItem.dICMSBase * objTributacaoDocItem.dICMSAliquota)
    Else
        objTributacaoDocItem.dICMSBase = 0
        objTributacaoDocItem.dICMSValor = 0
    End If
    
    'FCP
    If objTributacaoDocItem.dICMSpFCP <> 0 Then
        objTributacaoDocItem.dICMSvBCFCP = dValorLiquido
        objTributacaoDocItem.dICMSvFCP = Arredonda_Moeda(objTributacaoDocItem.dICMSvBCFCP * objTributacaoDocItem.dICMSpFCP)
    Else
        objTributacaoDocItem.dICMSvBCFCP = 0
        objTributacaoDocItem.dICMSvFCP = 0
    End If
    
    'pis
    If objTributacaoDocItem.dPISAliquota <> 0 Then
        objTributacaoDocItem.dPISBase = dValorLiquido
        objTributacaoDocItem.dPISValor = Arredonda_Moeda(objTributacaoDocItem.dPISBase * objTributacaoDocItem.dPISAliquota)
    Else
        If objTributacaoDocItem.dPISAliquotaValor <> 0 Then
            objTributacaoDocItem.dPISBase = 0 '??? rever
            objTributacaoDocItem.dPISQtde = objItem.dQuantidade  '??? rever pq pode ter que converter unidade de medida
            objTributacaoDocItem.dPISValor = Arredonda_Moeda(objTributacaoDocItem.dPISQtde * objTributacaoDocItem.dPISAliquota)
        Else
            objTributacaoDocItem.dPISBase = 0
            objTributacaoDocItem.dPISValor = 0
    
        End If
    End If
    
    'cofins
    If objTributacaoDocItem.dCOFINSAliquota <> 0 Then
        objTributacaoDocItem.dCOFINSBase = dValorLiquido
        objTributacaoDocItem.dCOFINSValor = Arredonda_Moeda(objTributacaoDocItem.dCOFINSBase * objTributacaoDocItem.dCOFINSAliquota)
    Else
        If objTributacaoDocItem.dCOFINSAliquotaValor <> 0 Then
            objTributacaoDocItem.dCOFINSQtde = objItem.dQuantidade  '??? rever pq pode ter que converter unidade de medida
            objTributacaoDocItem.dCOFINSValor = Arredonda_Moeda(objTributacaoDocItem.dCOFINSQtde * objTributacaoDocItem.dCOFINSAliquota)
        Else
            objTributacaoDocItem.dCOFINSBase = 0
            objTributacaoDocItem.dCOFINSValor = 0
    
        End If
    End If
    
    If objTributacaoDocItem.dISSAliquota <> 0 Then
        objTributacaoDocItem.dISSBase = dValorLiquido
        objTributacaoDocItem.dISSValor = Arredonda_Moeda(objTributacaoDocItem.dISSBase * objTributacaoDocItem.dISSAliquota)
    Else
        objTributacaoDocItem.dISSBase = 0
        objTributacaoDocItem.dISSValor = 0
    End If
    
End Sub

Private Sub Limpa_Tela_Venda_1()

    If giCodModeloECF = IMPRESSORA_NFCE And gobjNFeInfo.iEmContingencia <> 0 Then
        Exibe.Caption = "PRÓXIMO CLIENTE...(EM CONTINGENCIA)"
    Else
        Exibe.Caption = "PRÓXIMO CLIENTE..."
    End If
    Exibe1.Caption = ""
    
    giItem = 0
    Figura.Picture = LoadPicture
    
    Subtotal.Caption = Format(0, "standard")
    CodVendedor.Text = ""
'    CodCliente.Text = ""
    LabelVendedor.Caption = ""
    NomeCliente.Text = ""
    Email.Text = ""
    
    Orcamento.Text = ""
'    Orcamento.Enabled = False
'    LabelOrcamento.Enabled = False
    
    'Limpa toda a list
    ListCF.Clear
    
    'Inicia o Cupom
    Call Inicia_Cupom_Tela

End Sub

Public Sub Form_Resize(iLargura As Integer, iAltura As Integer)

Dim iIndice As Integer, lErro As Long
Dim lGridTopOrig As Long

On Error GoTo Erro_Form_Resize

    If gbCarregandoTela Then
    
        UserControl.Size iLargura, iAltura
            
        Frame3.Width = iLargura
    
    '        ListCF.Width = giLarguraListCF + (iLargura - giLarguraOrig)
    '        ListCF.Height = giAlturaListCF + (iAltura - giAlturaOrig)
    
        
    
'        Figura.Height = IIf(giAlturaFigura + (iAltura - giAlturaOrig) < 0, 1, giAlturaFigura + (iAltura - giAlturaOrig))
        
        
        Grid.Width = giLarguraGrid + (iLargura - giLarguraOrig)
    '        Grid.Height = giAlturaGrid + (iAltura - giAlturaOrig)
        
        Grid1.left = Grid.left
        Grid1.Width = Grid.Width
        ListCF.Width = Grid.Width
        Picture1.Width = Grid1.Width
        
        
        If Grid.Width < 12000 Then
            Grid.Font.Size = 8
            Grid1.Font.Size = 8
            ListCF.Font.Size = 8
            Picture1.Font.Size = 8
        Else
            Grid.Font.Size = 14
            Grid1.Font.Size = 14
            ListCF.Font.Size = 14
            Picture1.Font.Size = 14
        End If
        
        Grid1.Visible = True
        
        If Grid.Width < 8000 Then
            Grid.Visible = False
            ListCF.Visible = True
            Grid1.BorderStyle = flexBorderSingle
            Grid1.Height = 13 * Grid1.RowHeight(0)
            
        Else
            Grid.Visible = True
            ListCF.Visible = False
            Grid1.BorderStyle = flexBorderNone
            Grid1.Height = 11 * Grid1.RowHeight(0)
        
        End If
        
        lGridTopOrig = Grid.top
        Grid.top = Grid1.top + Grid1.Height
        ListCF.top = Grid1.top + Grid1.Height - 15
        ListCF.left = Grid1.left
        
        iAltura = iAltura - FrameProdutos.Height 'alteracao de layout colocando o teclado de produtos na parte de baixo
        
        Grid.Height = iAltura - GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelX - Grid.top
        ListCF.Height = iAltura - ListCF.top
        
        'Limpa toda a list
        ListCF.Clear
        
        Call Inicia_Cupom_Tela
        
        If AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
            'carregar venda atual, se houver
            lErro = CF_ECF("Carrega_Movimento_Atual", gobjVenda)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If gobjVenda.objCupomFiscal.colItens.Count <> 0 Then
            
                lErro = Carrega_Venda_Atual()
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
            
        End If

        gbCarregandoTela = False

    End If

    Exit Sub

Erro_Form_Resize:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175669)

    End Select

    Exit Sub

End Sub

Private Function ControleMaxTamVisivel(ByVal objControle As Object, sCaracter As String) As Integer
'Retorna o maximo de caracteres que cabem na largura de uma picturebox ou outro controle que tenha hdc
'o fonte utilizado tem que ser de caracteres com a mesma largura

Dim hdc As Long, sTexto As String, szTexto As typeSize, iTam As Integer

    iTam = 0
    
    hdc = objControle.hdc
    If hdc <> 0 Then
        
'        sTexto = "1234567890123456789012345678901234567890"
        sTexto = String(100, sCaracter)
        
        If GetTextExtentPoint32(objControle.hdc, sTexto, Len(sTexto), szTexto) <> 0 Then
            
'            iTam = Fix(objControle.Width * 40 / (szTexto.cx * Screen.TwipsPerPixelX))
            iTam = Fix(objControle.Width * 100 / (szTexto.cx * (Screen.TwipsPerPixelX)))
            
        End If
        
    End If
    
    ControleMaxTamVisivel = iTam
    
End Function

Private Function Proxima_Linha_Grid()

Dim iLinha As Integer

        giUltimaLinhaGrid = giUltimaLinhaGrid + 1
        If giUltimaLinhaGrid + 1 > Grid.Rows Then Grid.Rows = giUltimaLinhaGrid + 1
        
        If giUltimaLinhaGrid + 1 > giLinhasVisiveisGrid Then
            Grid.TopRow = giUltimaLinhaGrid - giLinhasVisiveisGrid + 3
        End If

End Function

Private Sub IEIsento_Click()
    gobjVenda.objCupomFiscal.iIEIsento = IEIsento.Value
End Sub

Public Sub Trata_IE()
    If Len(Trim(InscricaoEstadual.Text)) > 0 Then
        IEIsento.Value = vbUnchecked
        IEIsento.Enabled = False
        If InscricaoEstadual.Text <> sIEAnt Then
            IENaoContrib.Value = vbUnchecked
        End If
    Else
        If InscricaoEstadual.Text <> sIEAnt Then
            IEIsento.Value = vbChecked
            IENaoContrib.Value = vbChecked
        End If
        IEIsento.Enabled = True
    End If
    sIEAnt = InscricaoEstadual.Text
End Sub

Public Sub InscricaoEstadual_Validate(Cancel As Boolean)
    Call Trata_IE
    gobjVenda.objCupomFiscal.sInscricaoEstadual = InscricaoEstadual.Text
End Sub

Sub Limpa_Teclado()

Dim iIndice As Integer

    For iIndice = 0 To 34
    
        Produto(iIndice).Caption = ""
        Produto(iIndice).Tag = ""
        Produto(iIndice).BackColor = COR_DEFAULT
        Produto(iIndice).ToolTipText = Produto(iIndice).Caption
    
    Next

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(objProduto.sReferencia)) > 0 Then
        ProdutoNomeRed.Text = objProduto.sReferencia
    Else
        ProdutoNomeRed.Text = objProduto.sCodigoBarras
    End If
    Call ProdutoNomeRed_Validate(False)

'    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214934)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objCliente = obj1

    Select Case Len(Trim(objCliente.sCgc))

        Case STRING_CPF 'CPF
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "

        Case STRING_CGC 'CGC
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "

    End Select

    NomeCliente.Text = objCliente.sNomeReduzido
    gobjVenda.objCupomFiscal.sCPFCGC = objCliente.sCgc
    gobjVenda.objCupomFiscal.sNomeCliente = objCliente.sNomeReduzido
    CGC.Text = objCliente.sCgc
    gobjVenda.objCupomFiscal.sCPFCGC1 = CGC.FormattedText
    
    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214935)

    End Select

    Exit Sub
    

End Sub


Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1

    CodVendedor.Text = objVendedor.iCodigo
    Call CodVendedor_Validate(False)

'    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214956)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOrcamento_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrcamento As ClassOrcamentoLoja
Dim objVenda As New ClassVenda
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoOrcamento_evSelecao

        Set objOrcamento = obj1

        objVenda.objCupomFiscal.lNumOrcamento = objOrcamento.lNumOrcamento

        'Função Que le os orcamentos
        lErro = CF_ECF("OrcamentoECF_Le", objVenda)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 214959

        If lErro = 210447 Then gError 214960

        'orcamento nao cadastrado
        If lErro <> SUCESSO Then gError 214961


        For Each objItens In objVenda.objCupomFiscal.colItens

            Set objProduto = New ClassProduto

            lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214962

            If lErro = SUCESSO Then
                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
            End If


        Next

        'Traz ele para a tela
        Call Copia_Venda(gobjVenda, objVenda)
        Call Traz_Orcamento

        'se o cupom fiscal estiver ligado, cham OptionCF_Click para transformar o orcamento em cupom
        If OptionCF.Value Then Call OptionCF_Click

    Exit Sub

Erro_objEventoOrcamento_evSelecao:

    Select Case gErr

        Case 214959, 214962

        Case 214960
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda.objCupomFiscal.lNumOrcamento)

        Case 214961
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_CADASTRADO1, gErr, objOrcamento.lNumOrcamento)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214963)

    End Select

    Exit Sub

End Sub

Private Sub DescontoValor_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim objOperador As New ClassOperador
    
On Error GoTo Erro_DescontoValor_Validate
    
    'Se o valor foi preenchido
    If Len(Trim(DescontoValor.Text)) > 0 Then
    
        'Verifica se é um valor aceito
        lErro = Valor_NaoNegativo_Critica(DescontoValor.Text)
        If lErro <> SUCESSO Then gError 99608
        
        For Each objOperador In gcolOperadores
    
            If objOperador.iCodigo = giCodOperador Then
    
                objOperador.iLimiteDesconto = objOperador.iLimiteDesconto

                Exit For
    
            End If
    
        Next
        
        'Se for necessária a autorização do Gerente
        If objOperador.iLimiteDesconto <> 100 Then
    
            'Não pode passar o obj que aponta para o mesmo da coleção porque vai trocar o usuário
            Set objOperador = New ClassOperador
            
            'Chama a Tela de Senha
            Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
    
            'Sai de Função se a Tela de Login não Retornar ok
            If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        
    End If
        
    Exit Sub
    
Erro_DescontoValor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99608, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164142)

    End Select

    Exit Sub
    
End Sub


Private Sub BotaoReimprimir_Click()
    
Dim lErro As Long
Dim iCodigo As Integer
Dim lNumItens As Long
Dim iIndice As Integer
Dim objItens As New ClassItemCupomFiscal
Dim objAliquota As New ClassAliquotaICMS
Dim objVenda As New ClassVenda
Dim sRetorno As String
Dim vbMsgRes As VbMsgBoxResult
Dim lSequencial As Long
Dim colRegistro As New Collection
Dim sLog As String
Dim objCliente As ClassCliente
Dim sCPF As String
Dim lNumero As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer
Dim iFlag As Integer
Dim lRetorno As Long
Dim dtDataFinal As Date
Dim objTela As Object
Dim objVendaAux As ClassVenda
Dim lTamanho As Long
Dim objFormMsg As Object
Dim objMovCaixa As ClassMovimentoCaixa
Dim objMovCaixa1 As ClassMovimentoCaixa
Dim iCuponsVinculados As Integer
Dim colMeiosPag As New Collection

On Error GoTo Erro_BotaoReimprimir_Click
    
    Set objTela = Me
    
    'se for um cupom e o
    'cupom a ser impresso é um anterior(naum está na tela)
    If gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) And gobjVenda.objCupomFiscal.lNumero = 0 Then
        
        If gcolVendas.Count = 0 Then gError ERRO_SEM_MENSAGEM
        
        For iIndice = gcolVendas.Count To 1 Step -1
            
            Set objVendaAux = gcolVendas.Item(iIndice)
            
            If objVendaAux.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) Then
            
                lRetorno = objVendaAux.objCupomFiscal.lNumero
                
                'se o último número de cupom é o da última venda executada--> pode reimprimir esta venda
                If lRetorno = objVendaAux.objCupomFiscal.lCCF Then
                
                    'reimprimir o Cupom de Venda
                    Call CF_ECF("NFCE_Imprime", objTela, objVendaAux)

                End If
                
                Exit For
                
            Else
                
                If iIndice = 1 Then gError ERRO_SEM_MENSAGEM
            
            End If
        
        Next
            
    End If
            
    Exit Sub

Erro_BotaoReimprimir_Click:

    Select Case gErr
                
        Case ERRO_SEM_MENSAGEM
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175677)

    End Select
    
    Exit Sub
        
End Sub


Private Function CancelaCupomEletronico(ByVal iOrigem As Integer) As Long

Dim lErro As Long, iIndice As Integer
Dim lNum As Long, lRetorno As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iCodGerente As Integer
Dim objFormMsg As Object, objTela As Object
Dim objOperador As New ClassOperador
Dim objVendaAux As ClassVenda

On Error GoTo Erro_CancelaCupomEletronico

    Set objTela = Me
    
    lNum = Retorna_Count_ItensCupom
    
    If gobjVenda.objCupomFiscal.lNumero = 0 And lNum = 0 And iOrigem <> BOTAO_CANCELA_CUPOM_CFE Then
    
        lErro = Alteracoes_CancelamentoCupom(gobjVenda)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
    Else
    
        'Se tiver uma venda ainda nao autorizada acontecendo
        If gobjVenda.objCupomFiscal.lNumero = 0 And lNum <> 0 Then
        
            'já pergunta uma vez na propria tela de pagto e por isso nao precisa perguntar aqui
            If iOrigem <> BOTAO_PAGTO_CANCELA_CUPOM_CFE Then
            
                'Envia aviso perguntando se deseja cancelar a venda
                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_VENDA)
        
                If vbMsgRes <> vbYes Then gError ERRO_SEM_MENSAGEM
                    
                'Se for Necessário a autorização do Gerente para abertura do Caixa
                If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
            
                    'Chama a Tela de Senha
                    Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
            
                    'Sai de Função se a Tela de Login não Retornar ok
                    If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
                    
                    'Se Operador for Gerente
                    iCodGerente = objOperador.iCodigo
            
                End If
    
            End If
    
            lErro = Alteracoes_CancelamentoCupom(gobjVenda)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Else
        
            'Se tiver uma venda já autorizada na tela
            If gobjVenda.objCupomFiscal.lNumero <> 0 Then
            
                '?????
                gError 102506
                
            Else
        
                'nunca deveria chegar aqui
                If iOrigem <> BOTAO_CANCELA_CUPOM_CFE Then gError ERRO_SEM_MENSAGEM
        
                'a tela está limpa e então vai buscar a venda anterior para cancelar
                        
                For iIndice = gcolVendas.Count To 1 Step -1
                    
                    Set objVendaAux = gcolVendas.Item(iIndice)
                    
                    If objVendaAux.iTipo = OPTION_CF Then
                    
                        lRetorno = objVendaAux.objCupomFiscal.lNumero
                        
                        'se o último número de cupom é o da última venda executada--> pode cancelar esta venda
                        If lRetorno = objVendaAux.objCupomFiscal.lCCF Then
                        
                            'Envia aviso perguntando se deseja cancelar o cupom
                            vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM_ANTERIOR & vbNewLine & "No: " & CStr(objVendaAux.objCupomFiscal.lNumero) & " Data: " & Format(objVendaAux.objCupomFiscal.dtDataEmissao, "DD/MM/YYYY") & " R$: " & Format(objVendaAux.objCupomFiscal.dValorTotal, "STANDARD"))
                    
                            If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM
                            
                            'Se for Necessário a autorização do Gerente para abertura do Caixa
                            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
                        
                                'Chama a Tela de Senha
                                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
                        
                                'Sai de Função se a Tela de Login não Retornar ok
                                If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
                                
                                'Se Operador for Gerente
                                iCodGerente = objOperador.iCodigo
                            
                            End If
                            
                            Call WritePrivateProfileString(APLICACAO_ECF, "COO", CStr(objVendaAux.objCupomFiscal.lNumero), NOME_ARQUIVO_CAIXA)
                            
                            'cancelar o Cupom de Venda
                            lErro = AFRAC_CancelarCupom(Me, objVendaAux)
                            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                            Exibe.Caption = "CANCELADO CUPOM CCF = " & objVendaAux.objCupomFiscal.lCCF
                            Exibe1.Caption = ""
                            
                            'Fecha a Transação
                            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", objVendaAux)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                            lErro = Alteracoes_CancelamentoCupom(objVendaAux)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                            Set objFormMsg = MsgTEF
                            
                            'cancela os cartoes ja confirmados
                            lErro = CF_ECF("TEF_CNC_PAYGO", objVendaAux, objFormMsg, objTela)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                            lErro = CF_ECF("TEF_Imprime_CNC_PAYGO", objFormMsg, objTela, objVendaAux)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                            gcolVendas.Remove (iIndice)
                        
                        Else
                            
                            gError 112075
                        
                        End If
                        
                        Exit For
                        
                    Else
                        If iIndice = 1 Then gError 112075
                    End If
                
                Next
            
            End If
        
        End If
        
    End If
        
    'limpar movimentodecaixaatual
    Call CF_ECF("Apaga_Venda_Atual")
            
    Select Case iOrigem
    
        Case BOTAO_DAV_CANCELA_CUPOM_CFE, BOTAO_PREVENDA_CANCELA_CUPOM_CFE
            'mantem o que estiver na tela
        
        Case Else
        
            Set gobjVenda = New ClassVenda
            gobjVenda.iCodModeloECF = giCodModeloECF
                
            gobjVenda.iTipo = OPTION_CF
        
            If iOrigem <> BOTAO_FECHAR_CUPOM_CFE Then
            
                Call Limpa_Tela_Venda
        
                Quantidade.SelStart = 0
                Quantidade.SelLength = Len(Quantidade.Text)
                Quantidade.SetFocus
            
            End If
    
    End Select
    
    CancelaCupomEletronico = SUCESSO
    
    Exit Function
    
Erro_CancelaCupomEletronico:
    
    CancelaCupomEletronico = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            
        Case 112075, 102506
            Call Rotina_ErroECF(vbOKOnly, ERRO_CUPOM_NAO_CANCELADO, gErr)
                    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175678)

    End Select
        
    Exit Function
    
End Function

Private Function Carrega_Venda_Atual() As Long
    
Dim objItens As ClassItemCupomFiscal
Dim lErro As Long
Dim bAchou As Boolean
Dim objProduto As ClassProduto
Dim sProduto As String, sProduto1 As String
Dim lNum As Long, sPrecoItem As String
Dim lNumero As Long
Dim objMovCaixa As ClassMovimentoCaixa

On Error GoTo Erro_Carrega_Venda_Atual
    
    For Each objItens In gobjVenda.objCupomFiscal.colItens
            
        'Joga no cupom o item
        ProdutoNomeRed.Text = objItens.sProdutoNomeRed
        Quantidade.Text = objItens.dQuantidade
        PrecoUnitario.Caption = objItens.dPrecoUnitario
        
        If objItens.dValorDesconto <> 0 Then
            DescontoValor.Text = Format(objItens.dValorDesconto / objItens.dQuantidade, "standard")
        Else
            DescontoValor.Text = ""
        End If
        
        sProduto1 = ProdutoNomeRed.Text
        
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
        
        sPrecoItem = Format(StrParaDbl(Format(Quantidade.Text, "0.000")) * StrParaDbl(Format(objItens.dPrecoUnitario, "standard")), "Standard")
        PrecoItem.Caption = sPrecoItem
        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto), "standard")
                
        If Grid.Width < 8000 Then
    
            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "     " & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", objProduto.sCodigo) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", objProduto.sDescricao)
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", Format(Quantidade.Text, "#0.000")) & "  " & Formata_Campo(ALINHAMENTO_DIREITA, 4, " ", objProduto.sSiglaUMVenda) & " x " & Formata_Campo(ALINHAMENTO_DIREITA, 12, " ", Format(PrecoUnitario.Caption, "standard")) & objProduto.sSituacaoTribECF & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dAliquotaICMS * 100, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", Format(PrecoItem.Caption, "standard"))
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            
        Else
        
            Call Proxima_Linha_Grid
    
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_CODIGO) = objProduto.sCodigo
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = objProduto.sCodigo & " - " & objProduto.sDescricao
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_QUANTIDADE) = Format(Quantidade.Text, "#0.000") & " " & objProduto.sSiglaUMVenda
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_X) = "x"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(PrecoUnitario.Caption, "standard")
            'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = objProduto.sSituacaoTribECF & Format(objItens.dAliquotaICMS * 100, "fixed") & "%"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(PrecoItem.Caption, "standard")
        
        End If
        
        'se existir desconto sobre o item...
        If objItens.dValorDesconto <> 0 Then
        
            If Grid.Width < 8000 Then
        
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 21, " ", "-" & Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
        
            Else
            
                Call Proxima_Linha_Grid
                
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "DESCONTO: - R$ " & Format(objItens.dValorDesconto, "standard")
                'Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = "-" & Format(objItens.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            
            End If
            
            Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption), "standard")
            
        End If
        
        If objItens.iStatus = STATUS_CANCELADO Then
                
            'Atualiza o subtotal
            Subtotal.Caption = Format(Subtotal.Caption - ((objItens.dPrecoUnitario * objItens.dQuantidade) - objItens.dValorDesconto), "standard")
            
            If Grid.Width < 8000 Then
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", "           ***** ITEM " & objItens.iItem & " CANCELADO *****") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            Else
                Call Proxima_Linha_Grid
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "***** ITEM CANCELADO *****"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            
            End If
            
        End If
    
        'Para rolar automaticamente a barra de rolagem
        If Grid.Width < 8000 Then
            ListCF.ListIndex = ListCF.NewIndex
        End If
        
        Exibe.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "   " & Formata_Campo(ALINHAMENTO_DIREITA, 20, " ", objProduto.sNomeReduzido)
    
        If objProduto.dPercentMenosReceb > 0 Or objProduto.dDescontoValor > 0 Then
            Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(Quantidade.Text, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(PrecoUnitario.Caption, "standard")) & "-" & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
        Else
            Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(Quantidade.Text, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(PrecoUnitario.Caption, "standard")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "=") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(PrecoItem.Caption, "standard"))
        End If
        
        Set objItens.objTributacaoDocItem = New ClassTributacaoDocItem
        Call objItens.objTributacaoDocItem.Copia(objProduto.objTributacaoDocItem)
        'ajusta de acordo com o que efetivamente foi vendido
        Call ItemCupom_AjustaTrib(objItens)
    
    Next
    
    If gobjVenda.objCupomFiscal.iVendedor <> 0 Then
        CodVendedor.Text = gobjVenda.objCupomFiscal.iVendedor
    Else
        CodVendedor.Text = ""
    End If
    
    Call CodVendedor_Validate(False)
    
    NomeCliente.Text = gobjVenda.objCupomFiscal.sNomeCliente
    CGC.Text = gobjVenda.objCupomFiscal.sCPFCGC1
    Email.Text = gobjVenda.objCupomFiscal.sEndEntEmail
    
    Call Limpa_Cupom_Tela
    
    If giCodModeloECF = IMPRESSORA_NFCE Then
    
        If gobjVenda.objNFeInfo.iNFCeEnviarEmail <> 0 Then
        
            If gobjVenda.objNFeInfo.iNFCeImprimir <> 0 Then
                gobjVenda.objCupomFiscal.bNFCEImprimir = True
            Else
                gobjVenda.objCupomFiscal.bNFCEImprimir = False
            End If
            
        Else
            
            gobjVenda.objCupomFiscal.bNFCEImprimir = True
    
        End If
    
    End If
    
    Carrega_Venda_Atual = SUCESSO
    
    Exit Function

Erro_Carrega_Venda_Atual:
    
    Carrega_Venda_Atual = gErr
    
    Select Case gErr
                  
        Case 99818, 99884, 99912, 204210, 210510, 210516
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175688)

    End Select
    
    Exit Function
        
End Function

Private Sub TimerIntegracao_Timer()

Dim lErro As Long, lNum As Long
Dim objTela As Object

On Error GoTo Erro_TimerIntegracao_Timer
    
    lNum = Retorna_Count_ItensCupom
    
    If gobjVenda.objCupomFiscal.lNumero = 0 And lNum = 0 Then
    
        Set objTela = Me
    
        lErro = CF_ECF("Importar_Pedidos", objTela)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
            
    Exit Sub

Erro_TimerIntegracao_Timer:

    Select Case gErr
                
        Case ERRO_SEM_MENSAGEM
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175677)

    End Select
    
    Exit Sub
        
End Sub

Private Sub Venda_AjustaTrib()

Dim dSaldoFrete As Double, dValorFreteItem As Double, dValorLiquido As Double
Dim iItem As Integer, iUltItemNaoCancelado As Integer
Dim objItemCupom As ClassItemCupomFiscal

    iItem = 0
    
    For Each objItemCupom In gobjVenda.objCupomFiscal.colItens
    
        iItem = iItem + 1
        
        If objItemCupom.iStatus <> STATUS_CANCELADO Then
            
            iUltItemNaoCancelado = iItem
            
        End If
            
    Next
            
    dSaldoFrete = gobjVenda.objCupomFiscal.dValorAcrescimo
    
    iItem = 0
    
    For Each objItemCupom In gobjVenda.objCupomFiscal.colItens
    
        iItem = iItem + 1
        
        If objItemCupom.iStatus <> STATUS_CANCELADO Then
        
            'se for o ultimo item
            If iItem = iUltItemNaoCancelado Then
            
                dValorFreteItem = dSaldoFrete
                
            Else

                dValorLiquido = Arredonda_Moeda(Arredonda_Moeda(objItemCupom.dQuantidade * objItemCupom.dPrecoUnitario) - Arredonda_Moeda(objItemCupom.dValorDesconto))
                dValorFreteItem = Arredonda_Moeda(gobjVenda.objCupomFiscal.dValorAcrescimo * dValorLiquido / gobjVenda.objCupomFiscal.dValorTotal)
                dSaldoFrete = Arredonda_Moeda(dSaldoFrete - dValorFreteItem)
            
            End If
        
            objItemCupom.objTributacaoDocItem.dValorFreteItem = dValorFreteItem
                
        End If
        
    Next
    
    dValorFreteItem = 0
    For Each objItemCupom In gobjVenda.objCupomFiscal.colItens
        If objItemCupom.iStatus <> STATUS_CANCELADO Then
            If objItemCupom.objTributacaoDocItem.dValorFreteItem < 0 Then objItemCupom.objTributacaoDocItem.dValorFreteItem = 0
            dValorFreteItem = dValorFreteItem + objItemCupom.objTributacaoDocItem.dValorFreteItem
        End If
    Next
    dSaldoFrete = Arredonda_Moeda(gobjVenda.objCupomFiscal.dValorAcrescimo - dValorFreteItem)
    Do While Abs(dSaldoFrete) > DELTA_VALORMONETARIO
        For Each objItemCupom In gobjVenda.objCupomFiscal.colItens
            If objItemCupom.iStatus <> STATUS_CANCELADO Then
                If dSaldoFrete > 0 Then
                    objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem + 0.01
                    dSaldoFrete = dSaldoFrete - 0.01
                Else
                    If objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01 > -DELTA_VALORMONETARIO2 Then
                        objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01
                    End If
                    dSaldoFrete = dSaldoFrete + 0.01
                End If
                If Abs(dSaldoFrete) < DELTA_VALORMONETARIO Then Exit For
            End If
        Next
    Loop

    For Each objItemCupom In gobjVenda.objCupomFiscal.colItens
    
        If objItemCupom.iStatus <> STATUS_CANCELADO Then Call ItemCupom_AjustaTrib(objItemCupom)
    
    Next

End Sub
