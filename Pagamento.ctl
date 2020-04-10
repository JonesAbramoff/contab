VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl Pagamento 
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   11655
   Begin VB.Frame Frame2 
      Caption         =   "Pagamento"
      Height          =   3345
      Index           =   0
      Left            =   360
      TabIndex        =   38
      Top             =   840
      Width           =   8565
      Begin VB.Frame FrameDescontosAcrescimos 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1680
         Left            =   8265
         TabIndex        =   82
         Top             =   375
         Width           =   4590
         Begin MSMask.MaskEdBox DescontoValor1 
            Height          =   345
            Left            =   1710
            TabIndex        =   83
            Top             =   795
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DescontoPerc1 
            Height          =   345
            Left            =   3420
            TabIndex        =   84
            Top             =   780
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AcrescimoPerc 
            Height          =   345
            Left            =   3435
            TabIndex        =   85
            Top             =   1305
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DescontoValor 
            Height          =   345
            Left            =   1710
            TabIndex        =   86
            Top             =   270
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DescontoPerc 
            Height          =   345
            Left            =   3420
            TabIndex        =   87
            Top             =   270
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4305
            TabIndex        =   92
            Top             =   1335
            Width           =   240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4290
            TabIndex        =   91
            Top             =   825
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4290
            TabIndex        =   90
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "De&sconto 2:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   210
            TabIndex        =   89
            Top             =   795
            Width           =   1470
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "De&sconto 1:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   210
            TabIndex        =   88
            Top             =   300
            Width           =   1470
         End
      End
      Begin VB.CommandButton Desmembrar 
         Caption         =   "Desmembrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2895
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   11160
         Top             =   -360
      End
      Begin MSMask.MaskEdBox ProdutoNomeRed 
         Height          =   330
         Left            =   3480
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox AcrescimoValor 
         Height          =   345
         Left            =   1815
         TabIndex        =   93
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipo 
         Caption         =   "NFCe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   570
         TabIndex        =   97
         Top             =   300
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A Pagar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   705
         TabIndex        =   96
         Top             =   1905
         Width           =   1065
      End
      Begin VB.Label APagar 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   95
         Top             =   1905
         Width           =   2535
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tx. Entrega:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   255
         TabIndex        =   94
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   570
         TabIndex        =   50
         Top             =   765
         Width           =   1110
      End
      Begin VB.Label Total 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1815
         TabIndex        =   49
         Top             =   735
         Width           =   2520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Troco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   48
         Top             =   1905
         Width           =   780
      End
      Begin VB.Label Troco 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5595
         TabIndex        =   47
         Top             =   1905
         Width           =   2520
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4845
         TabIndex        =   46
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Pago 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5580
         TabIndex        =   45
         Top             =   810
         Width           =   2520
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Falta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4860
         TabIndex        =   44
         Top             =   1365
         Width           =   720
      End
      Begin VB.Label Falta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5580
         TabIndex        =   43
         Top             =   1350
         Width           =   2520
      End
      Begin VB.Label DataHora 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   7800
         TabIndex        =   42
         Top             =   -960
         Width           =   2715
      End
      Begin VB.Label Apresentacao 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   0
         TabIndex        =   41
         Top             =   -960
         Width           =   6555
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3225
      Left            =   9150
      ScaleHeight     =   3165
      ScaleWidth      =   2355
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   915
      Width           =   2415
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   13
         Left            =   360
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   "(Ctrl+F4)"
         Top             =   1935
         Width           =   690
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   360
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   "(Ctrl+F3)"
         Top             =   1485
         Width           =   690
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   11
         Left            =   465
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "(F3)"
         Top             =   1050
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   6
         Left            =   420
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   "(Esc)"
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   465
         TabIndex        =   79
         TabStop         =   0   'False
         Text            =   "(F11)"
         Top             =   165
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   405
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "(F2)"
         Top             =   2820
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   7
         Left            =   390
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "(F12)"
         Top             =   2370
         Width           =   540
      End
      Begin VB.CommandButton BotaoCaptura 
         Caption         =   "                 TEF Captura"
         Enabled         =   0   'False
         Height          =   360
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1920
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   270
         Picture         =   "Pagamento.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Fechar"
         Top             =   2745
         Width           =   1920
      End
      Begin VB.CommandButton BotaoAbrirGaveta 
         Caption         =   " A.G."
         Height          =   360
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Fechar"
         Top             =   2295
         Width           =   1920
      End
      Begin VB.CommandButton BotaoTEF 
         Caption         =   "TEF"
         Height          =   360
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   954
         Width           =   1920
      End
      Begin VB.CommandButton BotaoCancelar 
         Height          =   360
         Left            =   270
         Picture         =   "Pagamento.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Cancelar"
         Top             =   507
         Width           =   1920
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   270
         Picture         =   "Pagamento.ctx":04C0
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   1920
      End
      Begin VB.CommandButton BotaoTEFMultiplo 
         Caption         =   "                TEF Multiplo"
         Height          =   360
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1401
         Width           =   1920
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10920
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Meios de Pagamento"
      Height          =   4215
      Left            =   375
      TabIndex        =   19
      Top             =   4350
      Width           =   11175
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   5310
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "(F9)"
         Top             =   2880
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   10
         Left            =   5295
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "(F10)"
         Top             =   3270
         Width           =   480
      End
      Begin VB.CommandButton BotaoTicket 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2790
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   9
         Left            =   5295
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "(F8)"
         Top             =   2430
         Width           =   330
      End
      Begin VB.CommandButton BotaoTroca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2355
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   5310
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "(F7)"
         Top             =   2010
         Width           =   345
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   5295
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "(F6)"
         Top             =   1575
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   5310
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "(F5)"
         Top             =   1155
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   5295
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "(F4)"
         Top             =   735
         Width           =   330
      End
      Begin VB.CommandButton BotaoCartaoDebito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1500
         Width           =   2400
      End
      Begin VB.CommandButton BotaoCarne 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1935
         Width           =   2400
      End
      Begin VB.CommandButton BotaoCartaoCredito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1065
         Width           =   2400
      End
      Begin VB.CommandButton BotaoCheques 
         Caption         =   "   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   645
         Width           =   2400
      End
      Begin VB.CommandButton BotaoOutros 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3210
         Width           =   2400
      End
      Begin MSMask.MaskEdBox MaskCheques 
         Height          =   345
         Left            =   2250
         TabIndex        =   3
         Top             =   705
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskCartaoCredito 
         Height          =   345
         Left            =   2250
         TabIndex        =   5
         Top             =   1125
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskCartaoDebito 
         Height          =   345
         Left            =   2250
         TabIndex        =   7
         Top             =   1530
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskOutros 
         Height          =   375
         Left            =   2250
         TabIndex        =   11
         Top             =   3180
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskTicket 
         Height          =   375
         Left            =   2250
         TabIndex        =   9
         Top             =   2760
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDinheiro 
         Height          =   345
         Left            =   2250
         TabIndex        =   1
         Top             =   270
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   8130
         TabIndex        =   69
         Top             =   1545
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   8130
         TabIndex        =   68
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   16
         Left            =   8130
         TabIndex        =   67
         Top             =   3255
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   8130
         TabIndex        =   66
         Top             =   1110
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   18
         Left            =   8130
         TabIndex        =   65
         Top             =   690
         Width           =   165
      End
      Begin VB.Label Outros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   64
         Top             =   3240
         Width           =   2010
      End
      Begin VB.Label CartaoDebito 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   63
         Top             =   1530
         Width           =   2010
      End
      Begin VB.Label Carne 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   62
         Top             =   1950
         Width           =   2010
      End
      Begin VB.Label CartaoCredito 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   61
         Top             =   1095
         Width           =   2010
      End
      Begin VB.Label ChequeVista 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   60
         Top             =   660
         Width           =   2010
      End
      Begin VB.Label Dinheiro 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   59
         Top             =   255
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   8130
         TabIndex        =   58
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Troca 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   57
         Top             =   2400
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   8130
         TabIndex        =   56
         Top             =   2400
         Width           =   165
      End
      Begin VB.Label Ticket 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   55
         Top             =   2790
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   8130
         TabIndex        =   54
         Top             =   2805
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   20
         Left            =   4650
         TabIndex        =   33
         Top             =   1575
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   19
         Left            =   4665
         TabIndex        =   32
         Top             =   2835
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   4665
         TabIndex        =   31
         Top             =   3270
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Ticket :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   1320
         TabIndex        =   8
         Top             =   2805
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Troca :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   1365
         TabIndex        =   27
         Top             =   2400
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   4635
         TabIndex        =   22
         Top             =   690
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   4635
         TabIndex        =   21
         Top             =   1110
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Outros :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   1230
         TabIndex        =   10
         Top             =   3225
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Carnê :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   1320
         TabIndex        =   20
         Top             =   1980
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cartão Dé&bito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   1575
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ca&rtão Crédito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   285
         TabIndex        =   4
         Top             =   1110
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Cheques :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   975
         TabIndex        =   2
         Top             =   690
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Dinheiro :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1035
         TabIndex        =   0
         Top             =   285
         Width           =   1170
      End
   End
   Begin MSMask.MaskEdBox ProdNomeRed 
      Height          =   450
      Left            =   1440
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1920
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   794
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
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   -90
      Top             =   2370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   105
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Pagamento.ctx":09F2
   End
   Begin VB.Label Apresentacao1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   375
      TabIndex        =   36
      Top             =   150
      Width           =   6555
   End
   Begin VB.Label DataHora1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   8160
      TabIndex        =   35
      Top             =   105
      Width           =   2715
   End
End
Attribute VB_Name = "Pagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Guarda as informações da Venda que se deseja pagar.
Dim gobjVenda As ClassVenda
Dim gdDescontoAnterior As Double
Dim gdDescontoAnterior1 As Double
Dim gdAcrescimoAnterior As Double
Dim gdPercDescontoAnterior As Double
Dim gdPercDescontoAnterior1 As Double
Dim gdPercAcrescimoAnterior As Double
Dim giSaida As Integer
Dim gobjGenerico As AdmGenerico

'Variáveis para controlar a alteração dos valores da tela
'para evitar recálculo.

'??? 24/08/2016 Dim gdSaldoDinheiroAnterior As Double
Dim gdSaldoChequesAnterior As Double
Dim gdSaldoCartaoDebitoAnterior As Double
Dim gdSaldoCartaoCreditoAnterior As Double
Dim gdSaldoOutrosAnterior As Double
Dim gdSaldoTicketAnterior As Double
Dim gdSaldoTEFAnterior As Double

Public Sub Form_Load()
        
    Call Timer1_Timer
    
    giSaida = 0
    
    Apresentacao1.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
    
    UserControl.Parent.WindowState = 2
    
    If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
        LabelTipo.Visible = True
    End If
    
    'Inicializa os campos de Valores da Tela
    Call Inicializa_Valores
    
    If giTEF = NAO_TEM_TEF Then
        BotaoTEF.Enabled = False
        BotaoTEFMultiplo.Enabled = False
    End If
    
    If giDinheiroAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskDinheiro.Enabled = False
        Label1(1).Enabled = False
    End If
        
    If giChequeAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskCheques.Enabled = False
        Label1(2).Enabled = False
        BotaoCheques.Enabled = False
    End If
        
    If giCartaoCreditoAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskCartaoCredito.Enabled = False
        Label1(3).Enabled = False
        BotaoCartaoCredito.Enabled = False
    End If
        
    If giCartaoDebitoAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskCartaoDebito.Enabled = False
        Label1(4).Enabled = False
        BotaoCartaoDebito.Enabled = False
    End If
        
    If giCarneAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        Label1(5).Enabled = False
        BotaoCarne.Enabled = False
    End If
    
    If giTrocaAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        Label1(7).Enabled = False
        BotaoTroca.Enabled = False
    End If
    
    If giTicketAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskTicket.Enabled = False
        Label1(10).Enabled = False
        BotaoTicket.Enabled = False
    End If
    
    If giOutrosAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskOutros.Enabled = False
        Label1(6).Enabled = False
        BotaoOutros.Enabled = False
    End If
    
    If giCodModeloECF = IMPRESSORA_NFCE Or giCodModeloECF = IMPRESSORA_NFE Then
        FrameDescontosAcrescimos.Visible = False
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
End Sub

Private Sub Inicializa_Valores()
    
    APagar.Caption = Format(0, "Standard")
    Pago.Caption = Format(0, "Standard")
    Falta.Caption = Format(0, "Standard")
    Troco.Caption = Format(0, "Standard")
    Dinheiro.Caption = Format(0, "Standard")
    ChequeVista.Caption = Format(0, "Standard")
    CartaoCredito.Caption = Format(0, "Standard")
    CartaoDebito.Caption = Format(0, "Standard")
    Carne.Caption = Format(0, "Standard")
    Troca.Caption = Format(0, "Standard")
    Outros.Caption = Format(0, "Standard")
    Ticket.Caption = Format(0, "Standard")
    BotaoCheques.Caption = Format(0, "Standard")
    BotaoCartaoCredito.Caption = Format(0, "Standard")
    BotaoCartaoDebito.Caption = Format(0, "Standard")
    BotaoCarne.Caption = Format(0, "Standard")
    BotaoTroca.Caption = Format(0, "Standard")
    BotaoOutros.Caption = Format(0, "Standard")
    BotaoTicket.Caption = Format(0, "Standard")

    
    Exit Sub
    
End Sub

Function Trata_Parametros(objVenda As ClassVenda, objGenerico As AdmGenerico) As Long
        
    'Deixa a informação da Venda Passada disponível globalmente.
    Set gobjVenda = objVenda
    
    'Deixa a referencia para o objeto que vai retornar o status da tela
    Set gobjGenerico = objGenerico
    
    gobjGenerico.vVariavel = vbAbort
    
    'Joga o Valor Total do Cupom Fiscal (Formatado)
    Total.Caption = Format(gobjVenda.objCupomFiscal.dValorProdutos, "Standard")
    
    'Coloca na Tela os dados da Venda.
    Call Traz_Dados_Tela
        
    'Calcula os valores de totais e subtotais
    Call Recalcula_Valores1
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Sub Traz_Dados_Tela()

Dim objCheque As ClassChequePre
Dim objMovCaixa As ClassMovimentoCaixa
Dim objCarneParc As ClassCarneParcelas
Dim objTroca As ClassTroca
Dim dTotal As Double
Dim dTotal1 As Double
Dim dTotal2 As Double
Dim dTotal3 As Double
    
    'Calcula o somatório dos cheques
    For Each objCheque In gobjVenda.colCheques
        If objCheque.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
            dTotal = dTotal + objCheque.dValor
        Else
            dTotal1 = dTotal1 + objCheque.dValor
        End If
    Next
    
    'Exibe os valores de cheques na tela
    MaskCheques.Text = IIf(dTotal1 <> 0, Format(dTotal1, "Standard"), "")
    BotaoCheques.Caption = Format(dTotal, "Standard")
    ChequeVista.Caption = Format(dTotal + dTotal1, "Standard")
    
    'Zera os Totalizadores utilizados
    dTotal = 0:    dTotal1 = 0
    
    'Calcula o somatório das trocas
    For Each objTroca In gobjVenda.colTroca
        dTotal = dTotal + objTroca.dValor
    Next
    
    'Exibe os valores de Troca
    BotaoTroca.Caption = Format(dTotal, "Standard")
    Troca.Caption = Format(dTotal, "Standard")
    
    'Zera o totalizador utilizado
    dTotal = 0
    
    'Calcula o somatório do carnê
    For Each objCarneParc In gobjVenda.objCarne.colParcelas
        dTotal = dTotal + objCarneParc.dValor
    Next
    
    'Exibe os Valores de Carnê
    BotaoCarne.Caption = Format(dTotal, "Standard")
    Carne.Caption = Format(dTotal, "Standard")
    
    'Zera o totalizador utilizado
    dTotal = 0
    
    'Calcula o somatório do dinheiro e outros
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
        
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then
            'Acumula total em dinheiro
            dTotal = dTotal + objMovCaixa.dValor
        ElseIf objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS And objMovCaixa.iAdmMeioPagto <> 0 Then
            'Acumula total em outros especificado
            dTotal1 = dTotal1 + objMovCaixa.dValor
        ElseIf objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS And objMovCaixa.iAdmMeioPagto = 0 Then
            'Acumula total em outros não especificado
            dTotal2 = dTotal2 + objMovCaixa.dValor
        End If
    Next
    
    'Exibe os totais de dinheiro e Outros
    MaskDinheiro.Text = IIf(dTotal <> 0, Format(dTotal, "Standard"), "")
    Dinheiro.Caption = Format(dTotal, "Standard")
    Call MaskDinheiro_Validate(False)
    
    MaskOutros.Text = IIf(dTotal2 <> 0, Format(dTotal2, "Standard"), "")
    BotaoOutros.Caption = Format(dTotal1, "Standard")
    Outros.Caption = Format(dTotal1 + dTotal2, "Standard")
    
    'Zera os acumuladores utilizados
    dTotal = 0:  dTotal1 = 0:  dTotal2 = 0
    
    'Calcula o somatório dos Ticket
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then
            
            If objMovCaixa.iAdmMeioPagto <> 0 Then
                'Acumula Ticket especificados
                dTotal = dTotal + objMovCaixa.dValor
            Else
                'Acumula Ticket não especificados
                dTotal1 = dTotal1 + objMovCaixa.dValor
            End If
        End If
    Next
    
    'Exibe totais de Tickets
    MaskTicket.Text = IIf(dTotal1 <> 0, Format(dTotal1, "Standard"), "")
    BotaoTicket.Caption = Format(dTotal, "Standard")
    Ticket.Caption = Format(dTotal + dTotal1, "Standard")
    
    'Zera os acumuladores utilizados
    dTotal = 0: dTotal1 = 0
    
    'Calcula o somatório de Cartões
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Then
            If objMovCaixa.iAdmMeioPagto <> 0 Then
                'Acumula cartões de crédito especificados
                dTotal = dTotal + objMovCaixa.dValor
            Else
                'Acumula cartões de crédito não especificados
                dTotal1 = dTotal1 + objMovCaixa.dValor
            End If
        
        ElseIf objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO Then
            
            If objMovCaixa.iAdmMeioPagto <> 0 Then
                   'Acumula cartões de Débito especificados
                dTotal2 = dTotal2 + objMovCaixa.dValor
            Else
                'Acumula cartões de Débito não especificados
                dTotal3 = dTotal3 + objMovCaixa.dValor
            End If
        End If
    
    Next
    
    'Exibe os totais de cartões de crédito
    MaskCartaoCredito.Text = IIf(dTotal1 <> 0, Format(dTotal1, "Standard"), "")
    BotaoCartaoCredito.Caption = Format(dTotal, "Standard")
    CartaoCredito.Caption = Format(dTotal + dTotal1, "Standard")
    
    'Exibe os totais de cartões de Débito
    MaskCartaoDebito.Text = IIf(dTotal3 <> 0, Format(dTotal3, "Standard"), "")
    BotaoCartaoDebito.Caption = Format(dTotal2, "Standard")
    CartaoDebito.Caption = Format(dTotal3 + dTotal2, "Standard")
    
    'Exibe o Total do Cupom
    Total.Caption = Format(gobjVenda.objCupomFiscal.dValorProdutos, "Standard")
    
    'Exibe o Acrescimo e o Desconto
    AcrescimoValor.Text = IIf(gobjVenda.objCupomFiscal.dValorAcrescimo <> 0, Format(gobjVenda.objCupomFiscal.dValorAcrescimo, "standard"), "")
    DescontoValor.Text = IIf(gobjVenda.objCupomFiscal.dValorDesconto <> 0, Format(gobjVenda.objCupomFiscal.dValorDesconto, "Standard"), "")
    DescontoValor1.Text = IIf(gobjVenda.objCupomFiscal.dValorDesconto1 <> 0, Format(gobjVenda.objCupomFiscal.dValorDesconto1, "Standard"), "")
    
    Call AcrescimoValor_Validate(False)
    Call DescontoValor_Validate(False)
    Call DescontoValor1_Validate(False)
    
    'Recalcula os totalizadores da tela
    Call Recalcula_Valores1
    Call Recalcula_Valores2
    
    Exit Sub
    
End Sub

Sub Recalcula_Valores1()

    'Calcula quanto falta pagar
    APagar.Caption = Format(StrParaDbl(Total.Caption) - StrParaDbl(DescontoValor.Text) - StrParaDbl(DescontoValor1.Text) + StrParaDbl(AcrescimoValor.Text), "Standard")
        
    'Calcula o possível troco
    Call Calcula_Faltatroco
    
    Exit Sub
    
End Sub

Sub Calcula_Faltatroco()
    
    'Se tiver troco para ser dado
    If StrParaDbl(APagar.Caption) > StrParaDbl(Pago.Caption) Then
        'Informa o trcoo calculado
        Falta.Caption = Format(StrParaDbl(APagar.Caption) - StrParaDbl(Pago.Caption), "Standard")
        Troco.Caption = Format(0, "Standard")
    Else
        Troco.Caption = Format(StrParaDbl(Pago.Caption) - StrParaDbl(APagar.Caption), "Standard")
        Falta.Caption = Format(0, "Standard")
    End If
    
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)

Exit Sub
    
End Sub

Sub Recalcula_Valores2()

    'Recalcula quanto já foi pago
    Pago.Caption = Format(StrParaDbl(Dinheiro.Caption) + StrParaDbl(ChequeVista.Caption) + StrParaDbl(CartaoDebito.Caption) + StrParaDbl(CartaoCredito.Caption) + StrParaDbl(Carne.Caption) + StrParaDbl(Troca.Caption) + StrParaDbl(Outros.Caption) + StrParaDbl(Ticket.Caption), "Standard")
        
    'Verifica e calcula o troco
    Call Calcula_Faltatroco
    
    Exit Sub
        
End Sub

Private Sub Inclui_Movimento(dValor As Double, iTipo As Integer, Optional iTipoCartao As Integer = 0)

Dim objMovimento As New ClassMovimentoCaixa
Dim bAchou As Boolean
Dim iIndice As Integer
    
    bAchou = False
    
    'Para cada movimento da tela
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o Movimento
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        'Se o movimento for do tipo que foi passado e não especificado
        If objMovimento.iTipo = iTipo And (objMovimento.iAdmMeioPagto = 0 Or iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO) Then
        
            'Se o valor a atribuir do movimento for positivo
            If dValor > 0 Then
                'Atribui o novo valor ao movimento
                objMovimento.dValor = dValor
            
            'Senão
            Else
                'remove o movimento
                gobjVenda.colMovimentosCaixa.Remove (iIndice)
            End If
            
            bAchou = True
            Exit For
        End If
    Next

    'Se tiver valor a atribuir e o movimento não foi encontrado
    If Not (bAchou) And dValor > 0 Then
        
        'Cria um novo movimento
        Set objMovimento = New ClassMovimentoCaixa
        
        'Preenche o novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iTipo = iTipo
        objMovimento.iParcelamento = COD_A_VISTA
        objMovimento.dHora = CDbl(Time)
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = dValor
        objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
        objMovimento.iTipoCartao = iTipoCartao
        
        'caso seja do tipo de Dinheiro autualiza o saldo em dinheiro
        If iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then
'            gdSaldoDinheiro = Arredonda_Moeda(gdSaldoDinheiro + dValor)
            objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
        End If
        
        'Adiciona o novo movimento à coleção global da tela
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
        
    End If
    
    Exit Sub
    
End Sub

Private Sub BotaoCaptura_Click()

Dim lErro As Long
Dim sIndice As String
Dim sMsg As String
Dim lSequencial As Long
Dim lSequencialCaixa As Long
Dim objCheque As New ClassChequePre
Dim sRet As String

On Error GoTo Erro_BotaoTEF_Click
        
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207991
        
    lErro = CF_ECF("TEF_Gerenciador_Padrao_PAYGO")
    If lErro <> SUCESSO Then gError 133761

    lErro = CF_ECF("Testa_Limite_Desconto", gobjVenda)
    If lErro <> SUCESSO Then gError 126778

    'Se o valor q falta é zero  -->erro.
    If StrParaDbl(Falta.Caption) = 0 Then gError 112400
        
    gobjVenda.iTipo = OPTION_CF
'    gobjVenda.objCupomFiscal.dtDataEmissao = Date
'    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.dValorAcrescimo = StrParaDbl(AcrescimoValor.Text)
    gobjVenda.objCupomFiscal.dValorDesconto = StrParaDbl(DescontoValor.Text)
    gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
    'gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
'    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
 
    lErro = Informa_Meios_Pagto_TEF(StrParaDbl(Falta.Caption), TEF_PRE_AUTORIZACAO_CAPTURA)
    If lErro <> SUCESSO Then gError 112401
        
    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
    
    gobjGenerico.vVariavel = vbOK
    
    Unload Me
    
    giSaida = 1
    
    Exit Sub
        
Erro_BotaoTEF_Click:
    
    Select Case gErr
    
        Case 105797
            Call Rotina_ErroECF(vbOKOnly, ERRO_TEF_NAO_ATIVO, gErr)
    
        Case 112400
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_JA_PAGO, gErr)
        
        Case 112401, 112402, 126778, 133761, 207991
        
        Case 133782
            Call Rotina_ErroECF(vbOKOnly, ERRO_BOTAO_TEF_SEM_FALTA, gErr)
            
        Case 133784
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_TEF, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164140)

    End Select
        
    Exit Sub

End Sub

Private Sub BotaoTEFMultiplo_Click()

Dim lErro As Long
Dim sIndice As String
Dim sMsg As String
Dim lSequencial As Long
Dim lSequencialCaixa As Long
Dim objCheque As New ClassChequePre
Dim sRet As String
Dim dValorFalta As Double
Dim objMovCaixa As ClassMovimentoCaixa
Dim iIndice As Integer

On Error GoTo Erro_BotaoTEFMultiplo_Click
        
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207990
        
    lErro = CF_ECF("TEF_Gerenciador_Padrao_PAYGO")
    If lErro <> SUCESSO Then gError 214566
        
    lErro = CF_ECF("Testa_Limite_Desconto", gobjVenda)
    If lErro <> SUCESSO Then gError 133785

    'Se o valor q falta é zero  -->erro.
    If StrParaDbl(Falta.Caption) = 0 Then gError 133786
        
    gobjVenda.iTipo = OPTION_CF
'    gobjVenda.objCupomFiscal.dtDataEmissao = Date
'    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.dValorAcrescimo = StrParaDbl(AcrescimoValor.Text)
    gobjVenda.objCupomFiscal.dValorDesconto = StrParaDbl(DescontoValor.Text)
    gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
    'gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
    gobjVenda.dValorTEF = StrParaDbl(Falta.Caption)
'    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
    gobjVenda.objCupomFiscal.iECF = giCodECF

    Call Venda_AjustaTrib
    
    Call Chama_TelaECF_Modal("TEFMultiplo", gobjVenda)
        
    If giRetornoTela <> vbOK Then
    
        'cancelar cartoes
    
        For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
            Set objMovCaixa = gobjVenda.colMovimentosCaixa.Item(iIndice)
            If objMovCaixa.iTipoCartao = TIPO_TEF Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
        Next
    
'        'Limpa os dados da Venda
'        Set gobjVenda = New ClassVenda
'
'        'Retorna para a tela de venda a informação de cancelamento da venda
'        gobjGenerico.vVariavel = vbCancel
'
'        'Fecha a tela
'        Unload Me
'
'        giSaida = 1
    
    
    Else
        
        lErro = Informa_Meios_Pagto_TEF1(StrParaDbl(Falta.Caption))
        If lErro <> SUCESSO Then gError 133793
        
        Set gobjVenda = New ClassVenda
        gobjVenda.iCodModeloECF = giCodModeloECF
        
        gobjGenerico.vVariavel = vbOK
        
        Unload Me
        
        giSaida = 1
    
    End If
    
    Exit Sub
        
Erro_BotaoTEFMultiplo_Click:
    
    Select Case gErr
    
        Case 133785, 133793, 207990, 214566
    
        Case 133786
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_JA_PAGO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164141)

    End Select
        
    Exit Sub

End Sub

Private Sub DescontoPerc_GotFocus()
    
    gdPercDescontoAnterior = StrParaDbl(DescontoPerc.Text)
    
End Sub

Private Sub DescontoPerc1_GotFocus()
    
    gdPercDescontoAnterior1 = StrParaDbl(DescontoPerc1.Text)
    
End Sub

Private Sub DescontoValor_GotFocus()
    
    'Posiciona o cursor na frente do campo
    Call MaskEdBox_TrataGotFocus(DescontoValor)
    
    gdDescontoAnterior = StrParaDbl(DescontoValor.Text)
    
End Sub

Private Sub DescontoValor1_GotFocus()
    
    'Posiciona o cursor na frente do campo
    Call MaskEdBox_TrataGotFocus(DescontoValor1)
    
    gdDescontoAnterior1 = StrParaDbl(DescontoValor1.Text)
    
End Sub

Private Sub DescontoValor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DescontoValor_Validate
    
    'Se o valor foi preenchido
    If Len(Trim(DescontoValor.Text)) > 0 Then
        'Verifica se é um valor aceito
        lErro = Valor_NaoNegativo_Critica(DescontoValor.Text)
        If lErro <> SUCESSO Then gError 99608
    End If
        
    'Não permite desconto maior que o total a pagar menos a troca.
    If (StrParaDbl(DescontoValor.Text) + StrParaDbl(DescontoValor1.Text)) - (StrParaDbl(Total.Caption) - StrParaDbl(Troca.Caption)) > 0.0001 Then
        gError 99931
    End If
            
    If gdDescontoAnterior <> StrParaDbl(DescontoValor.Text) Then
            
        'Exibe o desconto na tela
        If StrParaDbl(DescontoValor.Text) > 0 Then
            DescontoValor.Text = Round(StrParaDbl(DescontoValor.Text), 2)
            If (StrParaDbl(Total.Caption) - StrParaDbl(Troca.Caption)) > 0 Then DescontoPerc.Text = Round(StrParaDbl(DescontoValor.Text) / (StrParaDbl(Total.Caption) - StrParaDbl(Troca.Caption)) * 100, 2)
        Else
            DescontoPerc.Text = ""
        End If
        
        If StrParaDbl(DescontoValor1.Text) > 0 And (StrParaDbl(Total.Caption) - (StrParaDbl(Troca.Caption) + StrParaDbl(DescontoValor.Text)) > 0) Then DescontoPerc1.Text = Round(StrParaDbl(DescontoValor1.Text) / (StrParaDbl(Total.Caption) - (StrParaDbl(Troca.Caption) + StrParaDbl(DescontoValor.Text))) * 100, 2)
        
        gobjVenda.objCupomFiscal.dValorDesconto = StrParaDbl(DescontoValor.Text)
        
        'Recalcula os totais de pagamento com o novo valor de desconto
        Call Recalcula_Valores1
        
    End If
    
    Exit Sub
    
Erro_DescontoValor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99608
        
        Case 99931
            Call Rotina_ErroECF(vbOKOnly, ERRO_DESCONTO_MAIOR, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164142)

    End Select

    Exit Sub
    
End Sub

Private Sub DescontoValor1_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DescontoValor1_Validate
    
    'Se o valor foi preenchido
    If Len(Trim(DescontoValor1.Text)) > 0 Then
        'Verifica se é um valor aceito
        lErro = Valor_NaoNegativo_Critica(DescontoValor1.Text)
        If lErro <> SUCESSO Then gError 126744
    End If
        
    'Não permite desconto maior que o total a pagar menos a troca.
    If (StrParaDbl(DescontoValor.Text) + StrParaDbl(DescontoValor1.Text)) - (StrParaDbl(Total.Caption) - StrParaDbl(Troca.Caption)) > 0.0001 Then
        gError 126745
    End If
            
    If gdDescontoAnterior1 <> StrParaDbl(DescontoValor1.Text) Then
            
        'Exibe o desconto na tela
        If StrParaDbl(DescontoValor1.Text) > 0 Then
            DescontoValor1.Text = Round(StrParaDbl(DescontoValor1.Text), 2)
            DescontoPerc1.Text = Round(StrParaDbl(DescontoValor1.Text) / (StrParaDbl(Total.Caption) - (StrParaDbl(Troca.Caption) + StrParaDbl(DescontoValor.Text))) * 100, 2)
        Else
            DescontoPerc1.Text = ""
        End If
        
        gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
        
        'Recalcula os totais de pagamento com o novo valor de desconto
        Call Recalcula_Valores1
        
    End If
    
    Exit Sub
    
Erro_DescontoValor1_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 126744
        
        Case 126745
            Call Rotina_ErroECF(vbOKOnly, ERRO_DESCONTO_MAIOR, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164143)

    End Select

    Exit Sub
    
End Sub

Private Sub DescontoPerc_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim dPercentDesc  As Double
Dim dPercentDescAnterior  As Double
Dim lTamanho As Long

On Error GoTo Erro_DescontoPerc_Validate
    
    If gdPercDescontoAnterior <> StrParaDbl(DescontoPerc.Text) Then
        'Se o percentual de desconto está preenchid
        If Len(Trim(DescontoPerc.Text)) > 0 Then
            'Critica se é um percentual válido
            lErro = Porcentagem_Critica(DescontoPerc.Text)
            If lErro <> SUCESSO Then gError 99609
        End If
        
        DescontoValor.Text = ""
        
        'Exibe o novo valor formatado
        If StrParaDbl(DescontoPerc.Text) > 0 Then
            DescontoPerc.Text = Round(StrParaDbl(DescontoPerc.Text), 2)
            dPercentDesc = StrParaDbl(DescontoPerc.Text)
            If (StrParaDbl(Total.Caption) - StrParaDbl(BotaoTroca.Caption)) > 0 Then DescontoValor.Text = Round((dPercentDesc / 100) * (StrParaDbl(Total.Caption) - StrParaDbl(BotaoTroca.Caption)), 2)
        End If
        
        If StrParaDbl(DescontoPerc1.Text) > 0 And (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption) + StrParaDbl(DescontoValor.Text))) > 0 Then DescontoValor1.Text = Round((StrParaDbl(DescontoPerc1.Text) / 100) * (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption) + StrParaDbl(DescontoValor.Text))), 2)
        
        gobjVenda.objCupomFiscal.dValorDesconto = StrParaDbl(DescontoValor.Text)
        gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
        
        'Recalcula os totais de pagamento com o novo valor de desconto
        Call Recalcula_Valores1
    End If
    
    Exit Sub
    
Erro_DescontoPerc_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99609
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164144)

    End Select

    Exit Sub
    
End Sub

Private Sub DescontoPerc1_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim dPercentDesc  As Double
Dim dPercentDescAnterior  As Double
Dim lTamanho As Long

On Error GoTo Erro_DescontoPerc1_Validate
    
    If gdPercDescontoAnterior1 <> StrParaDbl(DescontoPerc1.Text) Then
        'Se o percentual de desconto está preenchid
        If Len(Trim(DescontoPerc1.Text)) > 0 Then
            'Critica se é um percentual válido
            lErro = Porcentagem_Critica(DescontoPerc1.Text)
            If lErro <> SUCESSO Then gError 126746
        End If
        
        DescontoValor1.Text = ""
        
        'Exibe o novo valor formatado
        If StrParaDbl(DescontoPerc1.Text) > 0 Then
            DescontoPerc1.Text = Round(StrParaDbl(DescontoPerc1.Text), 2)
            dPercentDesc = StrParaDbl(DescontoPerc1.Text)
            If (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption)) + StrParaDbl(DescontoValor.Text)) > 0 Then DescontoValor1.Text = Round((dPercentDesc / 100) * (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption) + StrParaDbl(DescontoValor.Text))), 2)
        End If
        
        gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
        
        'Recalcula os totais de pagamento com o novo valor de desconto
        Call Recalcula_Valores1
    End If
    
    Exit Sub
    
Erro_DescontoPerc1_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 126746
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164145)

    End Select

    Exit Sub
    
End Sub

Private Sub AcrescimoValor_GotFocus()
    
    'Posiciona o cursor na frente do campo
    Call MaskEdBox_TrataGotFocus(AcrescimoValor)
    
    gdAcrescimoAnterior = StrParaDbl(AcrescimoValor.Text)
    
End Sub

Private Sub AcrescimoValor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_AcrescimoValor_Validate
    
    'Se o acrescimo estiver preenchido
    If Len(Trim(AcrescimoValor.Text)) > 0 Then
        'Verifica se um valor válido
        lErro = Valor_NaoNegativo_Critica(AcrescimoValor.Text)
        If lErro <> SUCESSO Then gError 99620
    End If
        
    'Se o Acrescimo dado é maior que o total
    If StrParaDbl(AcrescimoValor.Text) - StrParaDbl(Total.Caption) > 0.0001 Then
        gError 99932
    End If
        
    If gdAcrescimoAnterior <> StrParaDbl(AcrescimoValor.Text) Then
        
        'Exibe o valor formatado na tela
        If Len(AcrescimoValor.Text) > 0 Then
            AcrescimoValor.Text = Round(StrParaDbl(AcrescimoValor.Text), 2)
            AcrescimoPerc.Text = Round(StrParaDbl(AcrescimoValor.Text) / StrParaDbl(Total.Caption) * 100, 2)
        Else
            AcrescimoPerc.Text = ""
        End If
        
        gobjVenda.objCupomFiscal.dValorAcrescimo = StrParaDbl(AcrescimoValor.Text)
        
        'Recalcula os totais com o novo acrescimo
        Call Recalcula_Valores1
        
    End If
    
    Exit Sub
    
Erro_AcrescimoValor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99620
        
        Case 99932
            Call Rotina_ErroECF(vbOKOnly, ERRO_ACRESCIMO_MAIOR, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164146)

    End Select

    Exit Sub
    
End Sub

Private Sub AcrescimoPerc_GotFocus()
    
    gdPercAcrescimoAnterior = StrParaDbl(AcrescimoPerc.Text)
    
End Sub

Private Sub AcrescimoPerc_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim dAcrescimoPerc As Double

On Error GoTo Erro_AcrescimoPerc_Validate
            
    If StrParaDbl(AcrescimoPerc.Text) <> gdPercAcrescimoAnterior Then
    
        'Se estiver preenchido
        If Len(Trim(AcrescimoPerc.Text)) > 0 Then
            'Verifica se é um valor válido
            lErro = Porcentagem_Critica(AcrescimoPerc.Text)
            If lErro <> SUCESSO Then gError 99621
        End If
        
        AcrescimoValor.Text = ""
        
        'Coloca o valor formatado na tela
        If StrParaDbl(AcrescimoPerc.Text) > 0 Then
            AcrescimoPerc.Text = Round(StrParaDbl(AcrescimoPerc.Text), 2)
            dAcrescimoPerc = StrParaDbl(AcrescimoPerc.Text)
            If StrParaDbl(Total.Caption) > 0 Then AcrescimoValor.Text = Round((dAcrescimoPerc / 100) * StrParaDbl(Total.Caption), 2)
        End If
        
        gobjVenda.objCupomFiscal.dValorAcrescimo = StrParaDbl(AcrescimoValor.Text)
        
        'recalcula os totais levando em conta o novo acrescimo
        Call Recalcula_Valores1
        
    End If
    
    Exit Sub
    
Erro_AcrescimoPerc_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99621
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164147)

    End Select

    Exit Sub
    
End Sub


Private Sub MaskDinheiro_GotFocus()
        
    'Posiciona o cursor na frente
    Call MaskEdBox_TrataGotFocus(MaskDinheiro)
    
End Sub

Private Sub MaskDinheiro_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_MaskDinheiro_Validate
    
    'Se o vaor em dinheiro estiver preenchido
    If Len(Trim(MaskDinheiro.Text)) > 0 Then
        'Verifica se é válido
        lErro = Valor_NaoNegativo_Critica(MaskDinheiro.Text)
        If lErro <> SUCESSO Then gError 99622
    End If
      
    'Se o valor informado é diferente do que estava anteriormente
'    If gdSaldoDinheiroAnterior <> StrParaDbl(MaskDinheiro.Text) Then
    
        'COloca o valor formatado na tela
        Dinheiro.Caption = Format(StrParaDbl(MaskDinheiro.Text), "Standard")
        
        'Atualiza o movimenot referente ao pagamento em dinheiro
        Call Inclui_Movimento(StrParaDbl(MaskDinheiro.Text), MOVIMENTOCAIXA_RECEB_DINHEIRO)
    
        'recalcula os totais levando em conta o novo valor de pagamento em dinheiro
        Call Recalcula_Valores2
    
        If Len(Trim(MaskDinheiro.Text)) > 0 Then MaskDinheiro.Text = Round(StrParaDbl(MaskDinheiro.Text), 2)
        
        'Guarda o valor presente no campo dinheiro
'??? 24/08/2016         gdSaldoDinheiroAnterior = StrParaDbl(MaskDinheiro.Text)
    
        
'    End If
    
    Exit Sub
    
Erro_MaskDinheiro_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99622
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164148)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskCheques_GotFocus()
        
    'Posiciona o cursor na frente
    Call MaskEdBox_TrataGotFocus(MaskCheques)
    
End Sub

Private Sub MaskCheques_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim bAchou As Boolean
Dim objCheque As New ClassChequePre
Dim objMovCaixa As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim lTamanho As Long
Dim sRetorno As String

On Error GoTo Erro_MaskCheques_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskCheques.Text)) > 0 Then
        
        'Verifica se é um valor válido
        lErro = Valor_NaoNegativo_Critica(MaskCheques.Text)
        If lErro <> SUCESSO Then gError 99623
        
    End If
        
    'Se o valor nesse campo foi alterado
    If gdSaldoChequesAnterior <> StrParaDbl(MaskCheques.Text) Then
    
        If Len(Trim(MaskCheques.Text)) > 0 Then MaskCheques.Text = Round(StrParaDbl(MaskCheques.Text), 2)
    
        bAchou = False
        
        'Exibe formatado na tela
        ChequeVista.Caption = Format(StrParaDbl(BotaoCheques.Caption) + StrParaDbl(MaskCheques.Text), "Standard")
        
        'Para cada cheque
        For iIndice = gobjVenda.colCheques.Count To 1 Step -1
            'Pega o cheque
            Set objCheque = gobjVenda.colCheques.Item(iIndice)
            'Se ele for não especificado
            If objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO Then
                'Se o valor de cheque nao especificado for positivo
                If StrParaDbl(MaskCheques.Text) > 0 Then
                    'Guarda o valor e numero do Cheque
                    objCheque.dValor = StrParaDbl(MaskCheques.Text)
                Else
                    'remove o cheque nao especificado
                    gobjVenda.colCheques.Remove (iIndice)
                End If
                bAchou = True
                Exit For
            End If
        Next
        
        'Se achou o chque nao especificado
        If bAchou Then
            'Para cada movimento
            For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
                'Pega o movimento
                Set objMovCaixa = gobjVenda.colMovimentosCaixa.Item(iIndice)
                'Se for o cheque
                If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CHEQUE And objMovCaixa.lNumRefInterna = objCheque.lSequencialCaixa Then
                    'Se há valor em cheque nao especificado
                    If StrParaDbl(MaskCheques.Text) > 0 Then
                        'Atualiza o Valor do moivmento
                        objMovCaixa.dValor = StrParaDbl(MaskCheques.Text)
                    Else
                        'Retira o movimento
                        gobjVenda.colMovimentosCaixa.Remove (iIndice)
                    End If
                    Exit For
                End If
            Next
        
        'Se não achou
        Else
            'Se há valor de cheque a incluir
            If StrParaDbl(MaskCheques.Text) > 0 Then
                'Cria um novo cheque
                Set objCheque = New ClassChequePre
                'Preenche os dados defaults do cheque
                objCheque.dtDataDeposito = Date
                objCheque.dValor = StrParaDbl(MaskCheques.Text)
                objCheque.iFilialEmpresaLoja = giFilialEmpresa
                objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO
                'em cheque não especificado, o número é para ficar em branco.
                objCheque.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
                objCheque.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
                
                lTamanho = 50
                sRetorno = String(lTamanho, 0)
        
                Call GetPrivateProfileString(APLICACAO_CAIXA, "NumProxCheque", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
                If sRetorno <> String(lTamanho, 0) Then objCheque.lSequencialCaixa = StrParaLong(sRetorno)
                
                If objCheque.lSequencialCaixa = 0 Then objCheque.lSequencialCaixa = 1
            
                'Atualiza o sequencial de arquivo
                lErro = WritePrivateProfileString(APLICACAO_CAIXA, "NumProxCheque", CStr(objCheque.lSequencialCaixa + 1), NOME_ARQUIVO_CAIXA)
                If lErro = 0 Then gError 105775
                
                'Adiciona o cheque  na coleção da venda
                gobjVenda.colCheques.Add objCheque
                        
                'criar movimento para o cheque
                Set objMovCaixa = New ClassMovimentoCaixa
            
                'Preenche o novo movcaixa
                objMovCaixa.iFilialEmpresa = giFilialEmpresa
                objMovCaixa.iCaixa = giCodCaixa
                objMovCaixa.iCodOperador = giCodOperador
                objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CHEQUE
                objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE
                objMovCaixa.iParcelamento = COD_A_VISTA
                objMovCaixa.dtDataMovimento = Date
                objMovCaixa.dValor = StrParaDbl(MaskCheques.Text)
                objMovCaixa.dHora = CDbl(Time)
                objMovCaixa.lNumRefInterna = objCheque.lSequencialCaixa
                objMovCaixa.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
                
                'Adiciona o movimento a coleção de moivmewntos da venda
                gobjVenda.colMovimentosCaixa.Add objMovCaixa
            End If
            
        End If
                
        'Recalcula os totais
        Call Recalcula_Valores2
    
        'Guarda o valor presente no campo Cheque
        gdSaldoChequesAnterior = StrParaDbl(MaskCheques.Text)
    
    End If
    
    Exit Sub
    
Erro_MaskCheques_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99623
        
        Case 105775
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_CAIXA, "NumProxCheque", NOME_ARQUIVO_CAIXA)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164149)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskCartaoCredito_GotFocus()
    
    'Posiciona o cursor no início
    Call MaskEdBox_TrataGotFocus(MaskCartaoCredito)
    
End Sub

Private Sub MaskCartaoCredito_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_MaskCartaoCredito_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskCartaoCredito.Text)) > 0 Then
        'verifica se é um valor válido
        lErro = Valor_NaoNegativo_Critica(MaskCartaoCredito.Text)
        If lErro <> SUCESSO Then gError 99625
    
    End If
    
    'Se nao houve alteracao de valor
'    If gdSaldoCartaoCreditoAnterior <> StrParaDbl(MaskCartaoCredito.Text) Then
    
        'Exibe o valor formatado na tela
        CartaoCredito.Caption = Format(StrParaDbl(BotaoCartaoCredito.Caption) + StrParaDbl(MaskCartaoCredito.Text), "Standard")
        
        'Recalcula os valores
        Call Recalcula_Valores2
            
        'Inclui o movimento
        Call Inclui_Movimento(StrParaDbl(MaskCartaoCredito.Text), MOVIMENTOCAIXA_RECEB_CARTAOCREDITO, TIPO_POS)
    
        If Len(Trim(MaskCartaoCredito.Text)) > 0 Then MaskCartaoCredito.Text = Round(StrParaDbl(MaskCartaoCredito.Text), 2)
    
        'guarda o valor atual em cartão crédito nao especificado
        gdSaldoCartaoCreditoAnterior = StrParaDbl(MaskCartaoCredito.Text)
    
'    End If
    
    If gobjNFeInfo.iFocaTipoVenda = MARCADO And gobjVenda.iForcadoF5 = DESMARCADO Then
        'Fecha a tela e abre a certa
        If Not (StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO) Then
            gobjVenda.iTipoForcado = OPTION_ORCAMENTO
            Call BotaoFechar_Click
        End If
    End If
    
    Exit Sub
    
Erro_MaskCartaoCredito_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99625
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164150)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskCartaoDebito_GotFocus()
    
    'Posiciona o cursor no início
    Call MaskEdBox_TrataGotFocus(MaskCartaoDebito)
    
End Sub

Private Sub MaskCartaoDebito_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_MaskCartaoDebito_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskCartaoDebito.Text)) > 0 Then
        'verifica se é um valor válido
        lErro = Valor_NaoNegativo_Critica(MaskCartaoDebito.Text)
        If lErro <> SUCESSO Then gError 99624
    
    End If
    
    'Se nao houve alteracao de valor
'    If gdSaldoCartaoDebitoAnterior <> StrParaDbl(MaskCartaoDebito.Text) Then
    
        'Exibe o valor formatado na tela
        CartaoDebito.Caption = Format(StrParaDbl(BotaoCartaoDebito.Caption) + StrParaDbl(MaskCartaoDebito.Text), "Standard")
        
        'Recalcula os valores
        Call Recalcula_Valores2
            
        'Inclui o movimento
        Call Inclui_Movimento(StrParaDbl(MaskCartaoDebito.Text), MOVIMENTOCAIXA_RECEB_CARTAODEBITO, TIPO_POS)
    
        If Len(Trim(MaskCartaoDebito.Text)) > 0 Then MaskCartaoDebito.Text = Round(StrParaDbl(MaskCartaoDebito.Text), 2)
    
        'guarda o valor atual em cartão Débito nao especificado
        gdSaldoCartaoDebitoAnterior = StrParaDbl(MaskCartaoDebito.Text)
    
    
'    End If
    
    If gobjNFeInfo.iFocaTipoVenda = MARCADO And gobjVenda.iForcadoF5 = DESMARCADO Then
        'Fecha a tela e abre a certa
        If Not (StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO) Then
            gobjVenda.iTipoForcado = OPTION_ORCAMENTO
            Call BotaoFechar_Click
        End If
    End If
    
    Exit Sub
    
Erro_MaskCartaoDebito_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99624
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164151)

    End Select

    Exit Sub
    
End Sub


Private Sub MaskOutros_GotFocus()
    
    'Posiciona o Cursor no Inicio
    Call MaskEdBox_TrataGotFocus(MaskOutros)
    
End Sub

Private Sub MaskOutros_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_MaskOutros_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskOutros.Text)) > 0 Then
        'Verifica se o valor é válido
        lErro = Valor_NaoNegativo_Critica(MaskOutros.Text)
        If lErro <> SUCESSO Then gError 99750
        
    End If
    
    'Se o valor não foi alterado ==> Sai
'    If gdSaldoOutrosAnterior <> StrParaDbl(MaskOutros.Text) Then
    
        'Exibe o valor formatado
        Outros.Caption = Format(StrParaDbl(BotaoOutros.Caption) + StrParaDbl(MaskOutros.Text), "Standard")
            
        'Recalcula os totais
        Call Recalcula_Valores2
            
        'Atualiza o Movimento
        Call Inclui_Movimento(StrParaDbl(MaskOutros.Text), MOVIMENTOCAIXA_RECEB_OUTROS, TIPO_POS)
    
        If Len(Trim(MaskOutros.Text)) > 0 Then MaskOutros.Text = Round(StrParaDbl(MaskOutros.Text), 2)
    
        'Guarda o valor que está em outros
        gdSaldoOutrosAnterior = StrParaDbl(MaskOutros.Text)
    
'    End If
    
    Exit Sub
    
Erro_MaskOutros_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99750
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164152)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskTicket_GotFocus()
        
    'POsiciona o cursor no inicio
    Call MaskEdBox_TrataGotFocus(MaskTicket)
    
End Sub

Private Sub MaskTicket_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_MaskTicket_Validate
    
    'Se está preenchido
    If Len(Trim(MaskTicket.Text)) > 0 Then
        
        'Se o valor não é válido
        lErro = Valor_NaoNegativo_Critica(MaskTicket.Text)
        If lErro <> SUCESSO Then gError 99751
        
    End If
        
    'Se o valor não foi alterado ==> Sai
'    If gdSaldoTicketAnterior <> StrParaDbl(MaskTicket.Text) Then
    
        'Exibe o valor formatado
        Ticket.Caption = Format(StrParaDbl(BotaoTicket.Caption) + StrParaDbl(MaskTicket.Text), "Standard")
    
        'Recalcula os totais
        Call Recalcula_Valores2
            
        'ATualiza o movimento
        Call Inclui_Movimento(StrParaDbl(MaskTicket.Text), MOVIMENTOCAIXA_RECEB_VALETICKET, TIPO_POS)
    
        If Len(Trim(MaskTicket.Text)) > 0 Then MaskTicket.Text = Round(StrParaDbl(MaskTicket.Text), 2)
    
        'Guarda o valor atual
        gdSaldoTicketAnterior = StrParaDbl(MaskTicket.Text)
    
'    End If
    
    Exit Sub
    
Erro_MaskTicket_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99751
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164153)

    End Select

    Exit Sub
    
End Sub

Private Sub Troco_Tela()
    
Dim lErro As Long
Dim objMovCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Troco_Tela
    
    'Se não há troco para especificar --> erro.
    If StrParaDbl(Troco.Caption) = 0 Then
        Exit Sub
    End If
    
    'Joga o valor do troco no obj
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    
    'Calcula o  troco
    Call Calcula_Troco
    
'    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then

        'Chama tela de troco
        Call Chama_TelaECF_Modal("Troco", gobjVenda)
        
'    End If
        
    Exit Sub
        
Erro_Troco_Tela:
    
    Select Case gErr
        
        Case 99627
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_TROCO_NAO_ESPECIFICADO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164154)

    End Select

    Exit Sub
    
End Sub

Private Sub Calcula_Troco()
'Varrer a col de movimentos procurando movimentos de troco (din,carta,c/v)
'Acumula os valores de troco encontrados e se estiver faltando incluir o que falta em um movimento de troco em dinheiro
'Se não encontrar cria um com todo o valor para troco em dinheiro

Dim dTroco As Double
Dim dTroco1 As Double
Dim bAchou As Boolean
Dim objMovimento As ClassMovimentoCaixa
Dim iIndice As Integer

    dTroco = 0
    dTroco1 = 0
    
    'Para cada movimento
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        'Se for do tipo troco em dinheiro
        If objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO Then
            'Guarda o valor do movimento em dinheiro
            dTroco1 = objMovimento.dValor
            'Acumula o troco
            dTroco = dTroco + dTroco1
            
            bAchou = True

        'Se for do tipo Ticket
        ElseIf objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_VALE Then
            'Acumula o troco
            dTroco = dTroco + objMovimento.dValor
        
        'Se for do tipo ContraVale
        ElseIf objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_CONTRAVALE Then
            'Acumula a troco
            dTroco = dTroco + objMovimento.dValor
        End If
    Next
    
    'Se o troco da tela for maior do que o até agora especificado
    If StrParaDbl(Troco.Caption) - dTroco > 0.00001 Then
        'Calcula a diferença
        dTroco = StrParaDbl(Troco.Caption) - dTroco
        
        If bAchou Then
            'Acrescenta a diferenç a o troco em dinheiro
            For Each objMovimento In gobjVenda.colMovimentosCaixa
            'Se for do tipo dinheiro
                If objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO Then objMovimento.dValor = objMovimento.dValor + dTroco
            Next
        Else
            'Cria um movimento em dinheiro para a diferença
            Set objMovimento = New ClassMovimentoCaixa
            
            objMovimento.iFilialEmpresa = giFilialEmpresa
            objMovimento.iCaixa = giCodCaixa
            objMovimento.iCodOperador = giCodOperador
            objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO
            objMovimento.iParcelamento = COD_A_VISTA
            objMovimento.dHora = CDbl(Time)
            objMovimento.dtDataMovimento = Date
            objMovimento.dValor = dTroco
            objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
            objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
            
            gobjVenda.colMovimentosCaixa.Add objMovimento
        End If
    'se o troco diminuiu
    ElseIf StrParaDbl(Troco.Caption) = 0 Then
    
        For iIndice = (gobjVenda.colMovimentosCaixa.Count) To 1 Step -1
            Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
            If objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_CONTRAVALE Or objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_VALE Or objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
        Next
    
    Else
        'Se o troco em dinheiro for maior ou igual ao novo troco
        If dTroco1 - StrParaDbl(Troco.Caption) > 0.00001 Then
            'Exclui todos os recebimentos em vale e contra-vale
            For iIndice = (gobjVenda.colMovimentosCaixa.Count) To 1 Step -1
                Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
                If objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_CONTRAVALE Or objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_VALE Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
            Next
            'Joga o valor total do troco no movimento em dinheiro
            dTroco = StrParaDbl(Troco.Caption)
        Else
            'Se for do tipo dinheiro-->update com o valor restante para completar o troco
            dTroco = StrParaDbl(Troco.Caption) - (dTroco - dTroco1)
        End If
        
        For iIndice = (gobjVenda.colMovimentosCaixa.Count) To 1 Step -1
            Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
            If objMovimento.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO Then objMovimento.dValor = dTroco
        Next
        
    End If
        
    Exit Sub
    
End Sub

Private Sub BotaoCheques_Click()
    
Dim lErro As Long
Dim objCheques As New ClassChequePre
Dim dTotal As Double
Dim dTotal1 As Double
Dim iTotalAprovado As Integer

On Error GoTo Erro_BotaoCheques_Click
    
    gobjVenda.dFalta = StrParaDbl(Falta.Caption)
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoCheque", gobjVenda)
        
    'Faz o somatório dos cheques
    For Each objCheques In gobjVenda.colCheques
        If objCheques.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
            'Acumula os especificados
            dTotal = dTotal + objCheques.dValor
        Else
            'Acumula os não especificados
            dTotal1 = dTotal1 + objCheques.dValor
        End If
                
        iTotalAprovado = iTotalAprovado + objCheques.iAprovado
        
    Next
    
    ' *** para que não acumule o valor do cheque não especificado
    
    MaskCheques.Text = ""
    
    ' ******************* 31/10/2002 Sergio
    
    'Joga o valor do somatório no botão e na MaskedBox
    BotaoCheques.Caption = Format(dTotal, "Standard")
    If dTotal1 <> 0 Then MaskCheques.Text = Format(StrParaDbl(MaskCheques.Text) + dTotal1, "Standard")
    
    'Atualiza o total
    ChequeVista.Caption = Format(StrParaDbl(MaskCheques.Text) + StrParaDbl(BotaoCheques.Caption), "Standard")
     
    'Recalcula os Totais
    Call Recalcula_Valores2
            
    'imposicao da homologacao
    If StrParaDbl(Falta.Caption) = 0 And iTotalAprovado > 0 Then
        Call BotaoAbrirGaveta_Click
    End If
            
    Exit Sub
    
Erro_BotaoCheques_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164155)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCartaoDebito_Click()
    
Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim dTotal As Double
Dim iNumCartoes As Integer

On Error GoTo Erro_BotaoCartaoDebito_Click

    iNumCartoes = 0
    
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO And objMovimento.iAdmMeioPagto <> 0 Then
            iNumCartoes = iNumCartoes + 1
        End If
    Next
    
    gobjVenda.dFalta = StrParaDbl(Falta.Caption)
    
    'Chama tela de pagamento cheque modal
    If iNumCartoes > 1 Then
        Call Chama_TelaECF_Modal("PagamentoCartao", gobjVenda, MOVIMENTOCAIXA_RECEB_CARTAODEBITO)
    Else
        Call Chama_TelaECF_Modal("PagamentoCC", gobjVenda, MOVIMENTOCAIXA_RECEB_CARTAODEBITO)
    End If
           
    'Faz o somatório dos CartaoDebito
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO And objMovimento.iAdmMeioPagto <> 0 Then dTotal = dTotal + objMovimento.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoCartaoDebito.Caption = Format(dTotal, "Standard")
        
    'Atualiza o total
    CartaoDebito.Caption = Format(StrParaDbl(MaskCartaoDebito.Text) + StrParaDbl(BotaoCartaoDebito.Caption), "Standard")
    
    'Atualiza os totais
    Call Recalcula_Valores2
    
    Exit Sub
            
Erro_BotaoCartaoDebito_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164156)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCartaoCredito_Click()
    
Dim lErro As Long
Dim objMovCaixa As New ClassMovimentoCaixa
Dim dTotal As Double
Dim iNumCartoes As Integer

On Error GoTo Erro_BotaoCartaoCredito_Click

    iNumCartoes = 0
    
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO And objMovCaixa.iAdmMeioPagto <> 0 Then
            iNumCartoes = iNumCartoes + 1
        End If
    Next
    
    gobjVenda.dFalta = StrParaDbl(Falta.Caption)
    
    'Chama tela de pagamento cheque modal
    If iNumCartoes > 1 Then
        Call Chama_TelaECF_Modal("PagamentoCartao", gobjVenda, MOVIMENTOCAIXA_RECEB_CARTAOCREDITO)
    Else
        Call Chama_TelaECF_Modal("PagamentoCC", gobjVenda, MOVIMENTOCAIXA_RECEB_CARTAOCREDITO)
    End If
        
    'Faz o somatório dos CartaoCredito
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO And objMovCaixa.iAdmMeioPagto <> 0 Then dTotal = dTotal + objMovCaixa.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoCartaoCredito.Caption = Format(dTotal, "Standard")
        
    'Atualiza o Total De Cartão De Crédito
    CartaoCredito.Caption = Format(StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(BotaoCartaoCredito.Caption), "Standard")
    
    'Atualiza os Totais Da Tela
    Call Recalcula_Valores2
            
    Exit Sub
    
Erro_BotaoCartaoCredito_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164157)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCarne_Click()
    
Dim lErro As Long
Dim objCarneParc As New ClassCarneParcelas
Dim dTotal As Double

On Error GoTo Erro_BotaoCarne_Click
    
    gobjVenda.dValorTEF = StrParaDbl(Falta.Caption)
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoPrazo", gobjVenda)
        
    'Faz o somatório dos Carne
    For Each objCarneParc In gobjVenda.objCarne.colParcelas
        dTotal = dTotal + objCarneParc.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoCarne.Caption = Format(dTotal, "Standard")
        
    'exibe o Valor do Carnê formatado
    Carne.Caption = Format(StrParaDbl(BotaoCarne.Caption), "Standard")
    
    'recalcula os totais da tela
    Call Recalcula_Valores2
            
    Exit Sub
    
Erro_BotaoCarne_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164158)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoTroca_Click()
    
Dim lErro As Long
Dim objTroca As New ClassTroca
Dim dTotal As Double

On Error GoTo Erro_BotaoTroca_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoTroca", gobjVenda)
        
    'Faz o somatório dos Troca
    For Each objTroca In gobjVenda.colTroca
        dTotal = dTotal + objTroca.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoTroca.Caption = Format(dTotal, "Standard")
        
    'Exibe o total em roca formatado
    Troca.Caption = Format(StrParaDbl(BotaoTroca.Caption), "Standard")
    
    If StrParaDbl(DescontoPerc.Text) > 0 And (StrParaDbl(Total.Caption) - StrParaDbl(BotaoTroca.Caption)) > 0 Then
        DescontoValor.Text = Round((StrParaDbl(DescontoPerc.Text) / 100) * (StrParaDbl(Total.Caption) - StrParaDbl(BotaoTroca.Caption)), 2)
        gobjVenda.objCupomFiscal.dValorDesconto = StrParaDbl(DescontoValor.Text)
    End If
    
    If StrParaDbl(DescontoPerc1.Text) > 0 And (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption) + StrParaDbl(DescontoValor.Text))) > 0 Then
        DescontoValor1.Text = Round((StrParaDbl(DescontoPerc1.Text) / 100) * (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption) + StrParaDbl(DescontoValor.Text))), 2)
        gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
    End If
    
    'Atualiza os totais
    Call Recalcula_Valores2
            
    Exit Sub
    
Erro_BotaoTroca_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164159)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoOutros_Click()
    
Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim dTotal As Double

On Error GoTo Erro_BotaoOutros_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoOutros", gobjVenda)
        
    'Faz o somatório dos Outros
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS And objMovimento.iAdmMeioPagto <> 0 Then dTotal = dTotal + objMovimento.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoOutros.Caption = Format(dTotal, "Standard")
        
    'Exibe o valor formatado
    Outros.Caption = Format(StrParaDbl(MaskOutros.Text) + StrParaDbl(BotaoOutros.Caption), "Standard")
    
    'Recalcula os totais
    Call Recalcula_Valores2
    
    Exit Sub
            
Erro_BotaoOutros_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164160)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoTicket_Click()
    
Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim dTotal As Double

On Error GoTo Erro_BotaoTicket_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoTicket", gobjVenda)
        
    'Faz o somatório dos Ticket
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET And objMovimento.iAdmMeioPagto <> 0 Then dTotal = dTotal + objMovimento.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoTicket.Caption = Format(dTotal, "Standard")
        
    'Exibe o valor formatado
    Ticket.Caption = Format(StrParaDbl(MaskTicket.Text) + StrParaDbl(BotaoTicket.Caption), "Standard")
    
    'recalcula os totais
    Call Recalcula_Valores2
    
    Exit Sub
            
Erro_BotaoTicket_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164161)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
            
    gobjGenerico.vVariavel = vbAbort
    
    Unload Me

End Sub

Private Sub BotaoCancelar_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim objOperador As New ClassOperador
Dim lErro As Long
Dim iCodGerente As Integer

On Error GoTo Erro_BotaoCancelar_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207988

    'Envia aviso perguntando se realmente deseja Cancelar a Compra
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELAR_COMPRA)
    
    'Se não quise ==> Sai
    If vbMsgRes = vbNo Then Exit Sub
        
    'Se for Necessário a autorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 102505
        
        'Se Operador for Gerente
        iCodGerente = objOperador.iCodigo

    End If
    
    'Limpa os dados da Venda
    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
    
    'Retorna para a tela de venda a informação de cancelamento da venda
    gobjGenerico.vVariavel = vbCancel
    
    'Fecha a tela
    Unload Me
    
    giSaida = 1
    
    Exit Sub
            
Erro_BotaoCancelar_Click:
    
    Select Case gErr
        
        Case 102505, 207988
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164162)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'Função que tem as chamadas para as Funções que limpam a tela

Dim vbMsgRes As VbMsgBoxResult
Dim objMovCaixa As New ClassMovimentoCaixa
Dim bAchou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Botaolimpar_Click
                    
    Set gobjVenda.colMovimentosCaixa = New Collection
    Set gobjVenda.colCheques = New Collection
    Set gobjVenda.objCarne = New ClassCarne
        
    Call Limpa_Tela_Pagamento
    
    Call Recalcula_Valores1
    
    Exit Sub
        
Erro_Botaolimpar_Click:

    Select Case gErr
    
        Case 99504
            'Erro Tratado dentro da Função Chamadora
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164163)

    End Select
    
    Exit Sub
        
End Sub

Private Sub Limpa_Tela_Pagamento()
    
    'Reinicializa os valores
    Call Inicializa_Valores
    
    MaskDinheiro.Text = ""
    MaskCheques.Text = ""
    MaskCartaoCredito.Text = ""
    MaskCartaoDebito.Text = ""
    MaskOutros.Text = ""
    MaskTicket.Text = ""
    
End Sub

Private Sub TEF_Click()

End Sub



Private Sub Timer1_Timer()

Dim sHora As String
Dim iPosHora As Integer
Dim sMinuto As String
Dim iPosMinuto As Integer
Dim ssegundo As String
Dim iPossegundo As Integer
Dim vbMsgBox As VbMsgBoxResult
Dim bAchou As Boolean
Dim lErro As Long
Dim sMinutoAnt As String
Dim dTimerTemp As Double
Dim dtData As Date
Dim dtime As Double

    dtime = Timer
    If dtime > 3600 Then
        'Coloca a hora atual do Sistema
        sHora = CStr(dtime / (60 * 60))
        iPosHora = InStr(1, sHora, ",")
        If iPosHora > 0 Then sHora = Mid(sHora, 1, iPosHora - 1)
    Else
        sHora = 0
    End If
    
    If sHora <> 0 Then
        dTimerTemp = dtime - (CLng(sHora * 3600))
    Else
        dTimerTemp = dtime
    End If
    
    If dTimerTemp > 60 Then
        sMinuto = CStr(dtime / 60) - (CInt(sHora * 60))
        iPosMinuto = InStr(1, sMinuto, ",")
        If iPosMinuto > 0 Then sMinuto = Mid(sMinuto, 1, iPosMinuto - 1)
    Else
        sMinuto = 0
    End If
    
    ssegundo = CStr(dtime) - ((CLng(sMinuto * 60)) + (CLng(sHora * 3600)))
    iPossegundo = InStr(1, ssegundo, ",")
    If iPossegundo > 0 Then ssegundo = Mid(ssegundo, 1, iPossegundo - 1)
    
    DataHora1.Caption = Format(Date, "dd/mm/yyyy") & "   " & Format(sHora, "00") & ":" & Format(sMinuto, "00") & ":" & Format(ssegundo, "00")
    
    'DataHora1.Caption = DataHora1.Caption & " R$ " & CStr(gdSaldoDinheiro)

    Exit Sub

End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If giSaida = 1 Then Exit Sub
    
    If Shift = vbCtrlMask Then
    
        Select Case KeyCode
    
            Case vbKeyF3
                If BotaoTEFMultiplo.Enabled = True Then
                    Call BotaoTEFMultiplo_Click
                End If
                
            Case vbKeyF4
                If BotaoCaptura.Enabled = True Then
                    Call BotaoCaptura_Click
                End If
    
        End Select
        
    ElseIf Shift = vbShiftMask Then
    
        Select Case KeyCode
   
            Case vbKeyF5
                If Not TrocaFoco(Me, Nothing) Then Exit Sub
                If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
                    'Fecha a tela e abre a certa
                    If Not (StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO) Then
                        gobjVenda.iTipoForcado = OPTION_ORCAMENTO
                        gobjVenda.iForcadoF5 = DESMARCADO
                        Call BotaoFechar_Click
                        Exit Sub
                    End If
                End If
    
        End Select
    
    Else
    
        Select Case KeyCode
            
            Case vbKeyReturn
                KeyCode = vbKeyTab
        
            Case vbKeyF3
                If BotaoTEF.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoTEF) Then Exit Sub
                    Call BotaoTEF_Click
                End If
                
            Case vbKeyF4
                If BotaoCheques.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoCheques) Then Exit Sub
                    Call BotaoCheques_Click
                End If
            
            Case vbKeyF5
                If BotaoCartaoCredito.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoCartaoCredito) Then Exit Sub
                    Call BotaoCartaoCredito_Click
                End If
                
            Case vbKeyF6
                If BotaoCartaoDebito.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoCartaoDebito) Then Exit Sub
                    Call BotaoCartaoDebito_Click
                End If
            
            Case vbKeyF7
                If BotaoCarne.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoCarne) Then Exit Sub
                    Call BotaoCarne_Click
                End If
            
            Case vbKeyF8
                If BotaoTroca.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoTroca) Then Exit Sub
                    Call BotaoTroca_Click
                End If
            
            Case vbKeyF9
                If BotaoTicket.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoTicket) Then Exit Sub
                    Call BotaoTicket_Click
                End If
                    
            Case vbKeyF10
                If BotaoOutros.Enabled = True Then
                    If Not TrocaFoco(Me, BotaoOutros) Then Exit Sub
                    Call BotaoOutros_Click
                End If
                
            Case vbKeyF11
                If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
                Call BotaoLimpar_Click
            
            Case vbKeyEscape
                If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
                Call BotaoCancelar_Click
                
            Case vbKeyF12
                If Not TrocaFoco(Me, BotaoAbrirGaveta) Then Exit Sub
                Call BotaoAbrirGaveta_Click
               
            Case vbKeyF2
                If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
                Call BotaoFechar_Click
            
        End Select
    
    
    End If
   
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
 Dim lTamanho As Long
 Dim sRetorno As String
    
    If gobjGenerico.vVariavel <> vbCancel Then
        lTamanho = 10
        sRetorno = String(lTamanho, 0)
        Call GetPrivateProfileString(APLICACAO_ECF, "CupomAberto", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        'se o cupom não foi encerrado
        'If CInt(sRetorno) <> 0 Then gobjGenerico.vVariavel = vbAbort
    End If
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
        
    Set gobjVenda = Nothing
        
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
    
    Caption = Formata_Campo(ALINHAMENTO_DIREITA, 20, " ", "Pagamento") & "Filial : " & giFilialEmpresa & "    Caixa : " & giCodCaixa & "    Operador : " & sOper
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Pagamento"

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

Private Sub BotaoTEF_Click()

Dim lErro As Long
Dim sIndice As String
Dim sMsg As String
Dim lSequencial As Long
Dim lSequencialCaixa As Long
Dim objCheque As New ClassChequePre
Dim sRet As String

On Error GoTo Erro_BotaoTEF_Click
        
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207989
        
    lErro = CF_ECF("TEF_Gerenciador_Padrao_PAYGO")
    If lErro <> SUCESSO Then gError 133761

    lErro = CF_ECF("Testa_Limite_Desconto", gobjVenda)
    If lErro <> SUCESSO Then gError 126778

    'Se o valor q falta é zero  -->erro.
    If StrParaDbl(Falta.Caption) = 0 Then gError 112400
        
    gobjVenda.iTipo = OPTION_CF
'    gobjVenda.objCupomFiscal.dtDataEmissao = Date
'    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.dValorAcrescimo = StrParaDbl(AcrescimoValor.Text)
    gobjVenda.objCupomFiscal.dValorDesconto = StrParaDbl(DescontoValor.Text)
    gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
    'gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
'    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
    gobjVenda.objCupomFiscal.iECF = giCodECF
    
    Call Venda_AjustaTrib

    lErro = Informa_Meios_Pagto_TEF(StrParaDbl(Falta.Caption))
    If lErro <> SUCESSO Then gError 112401
        
    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
    
    gobjGenerico.vVariavel = vbOK
    
    Unload Me
    
    giSaida = 1
    
    Exit Sub
        
Erro_BotaoTEF_Click:
    
    Select Case gErr
    
        Case 105797
            Call Rotina_ErroECF(vbOKOnly, ERRO_TEF_NAO_ATIVO, gErr)
    
        Case 112400
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_JA_PAGO, gErr)
        
        Case 112401, 112402, 126778, 133761, 207989
        
        Case 133782
            Call Rotina_ErroECF(vbOKOnly, ERRO_BOTAO_TEF_SEM_FALTA, gErr)
            
        Case 133784
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_TEF, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164164)

    End Select
        
    Exit Sub
    
End Sub

Private Function Informa_Meios_Pagto_TEF(ByVal dValorTEF As Double, Optional iCaptura As Integer = 0) As Long

Dim lErro As Long
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim iTipo As Integer
Dim iIndice As Integer
Dim objFormMsg As Object
Dim objTela As Object

On Error GoTo Erro_Informa_Meios_Pagto_TEF
    
    'Atualiza o arquivo(aberto e com TEF)
    Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "2", NOME_ARQUIVO_CAIXA)
        
    Set objTela = Me
    Set objFormMsg = MsgTEF
        
    'Atualiza o arquivo(aberto e com Multiplo TEF)
    Call WritePrivateProfileString(APLICACAO_ECF, "COO", CStr(gobjVenda.objCupomFiscal.lNumero), NOME_ARQUIVO_CAIXA)
        
    'Executa o processo de Tranferencia eletrônica
    lErro = CF_ECF("TEF_Venda", dValorTEF, gobjVenda.objCupomFiscal.lNumero, gobjVenda, objFormMsg, objTela, iCaptura)
    If lErro <> SUCESSO Then gError 112404
    
    lErro = Informa_Meios_Pagto_TEF1(StrParaDbl(Falta.Caption))
    If lErro <> SUCESSO Then gError 133794
    
    Informa_Meios_Pagto_TEF = SUCESSO
    
    Exit Function
        
Erro_Informa_Meios_Pagto_TEF:
    
    Informa_Meios_Pagto_TEF = gErr
    
    Select Case gErr
    
        Case 133794
            
        Case 112404
            'Atualiza o arquivo(aberto e sem TEF)
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "1", NOME_ARQUIVO_CAIXA)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164165)

    End Select
        
    Exit Function
        
End Function

Private Function Informa_Meios_Pagto_TEF1(ByVal dValorTEF As Double) As Long

Dim lErro As Long
Dim sMsg As String
Dim sIndice As String
Dim objMovCaixa As New ClassMovimentoCaixa
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim iTipo As Integer
Dim iIndice As Integer
Dim colMeiosPag As New Collection
Dim objMovCaixa1 As New ClassMovimentoCaixa
Dim objMovCaixa2 As New ClassMovimentoCaixa
Dim iNovoIndice As Integer
Dim iIndice2 As Integer
Dim objAux As New ClassMovimentoCaixa
Dim dMenor As Double
Dim dTotal As Double
Dim lNumero As Long
Dim sDescricao As String
Dim objFormMsg As Object
Dim objTela As Object

On Error GoTo Erro_Informa_Meios_Pagto_TEF1

    'iNFORMAR PARA IMPRESSORA AS FORMAS PAGTO
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
        
        lErro = CF_ECF("Trata_MovCaixa", objMovCaixa, colMeiosPag)
        If lErro <> SUCESSO Then gError 133733
        
    Next
    
    'ordenar por valores...
    For iIndice = 1 To colMeiosPag.Count - 1
        Set objMovCaixa = colMeiosPag.Item(iIndice)
        dMenor = objMovCaixa.dValor
        iNovoIndice = iIndice
        For iIndice2 = iIndice To colMeiosPag.Count
            Set objMovCaixa1 = colMeiosPag.Item(iIndice2)
            If objMovCaixa1.dValor < dMenor Then
                dMenor = objMovCaixa1.dValor
                iNovoIndice = iIndice2
            End If
        Next
        Set objMovCaixa1 = colMeiosPag.Item(iNovoIndice)
        Call Inverte_Col(objAux, objMovCaixa)
        Call Inverte_Col(objMovCaixa, objMovCaixa1)
        Call Inverte_Col(objMovCaixa1, objAux)
    Next
    
    lErro = Executa_Fechamento_Cupom(dValorTEF, gobjVenda.objCupomFiscal.lNumero, colMeiosPag)
    If lErro <> SUCESSO Then gError 133712

    Informa_Meios_Pagto_TEF1 = SUCESSO
    
    Exit Function
        
Erro_Informa_Meios_Pagto_TEF1:
    
    Informa_Meios_Pagto_TEF1 = gErr
    
    'iNFORMAR PARA IMPRESSORA AS FORMAS PAGTO
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        
        Set objMovCaixa = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovCaixa.iTipoCartao = TIPO_TEF Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
        
    Next
    
    Select Case gErr
    
        Case 133712, 133733
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164166)

    End Select
        
    Exit Function

End Function

Private Sub BotaoAbrirGaveta_Click()

Dim lErro As Long
Dim sIndice As String
Dim sMsg As String
Dim lSequencial As Long
Dim objCheque As New ClassChequePre
Dim objOperador As New ClassOperador

On Error GoTo Erro_BotaoAbrirGaveta_Click
        
'    If gobjNFeInfo.iFocaTipoVenda = MARCADO And gobjVenda.iForcadoF5 = DESMARCADO Then
'        'Fecha a tela e abre a certa
'        If Not (StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO) Then
'            gobjVenda.iTipoForcado = OPTION_ORCAMENTO
'            Call BotaoFechar_Click
'            Exit Sub
'        End If
'
'    ElseIf gobjNFeInfo.iFocaTipoVenda = MARCADO And gobjVenda.iForcadoF5 = MARCADO Then
'        'Se forçou manualmente pede autorização do gerente
'        If Not (StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO) Then
'            'Chama a Tela de Senha
'            Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
'
'            'Sai de Função se a Tela de Login não Retornar ok
'            If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
'        End If
'    End If
    
    If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
        'Se forçou manualmente pede autorização do gerente
        If Not (StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO) Then
            'Chama a Tela de Senha
            Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
    
            'Sai de Função se a Tela de Login não Retornar ok
            If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
        End If
    End If
        
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207992
        
    'Se estiver preenchido
'    If StrParaDbl(MaskDinheiro.Text) > 0 Then Call MaskDinheiro_Validate(False)
'    If StrParaDbl(MaskCheques.Text) > 0 Then Call MaskCheques_Validate(False)
'    If StrParaDbl(MaskCartaoCredito.Text) > 0 Then Call MaskCartaoCredito_Validate(False)
'    If StrParaDbl(MaskCartaoDebito.Text) > 0 Then Call MaskCartaoDebito_Validate(False)
'    If StrParaDbl(MaskTicket.Text) > 0 Then Call MaskTicket_Validate(False)
'    If StrParaDbl(MaskOutros.Text) > 0 Then Call MaskOutros_Validate(False)

    lErro = CF_ECF("Testa_Limite_Desconto", gobjVenda)
    If lErro <> SUCESSO Then gError 126777
        
    'Se o valor é insuficiente para pagar
    If StrParaDbl(Falta.Caption) > 0 Then gError 99746
        
    If StrParaDbl(Pago.Caption) = 0 Then gError 133759
    
    'Calcula o troca da tela
    Call Calcula_Troco
        
    gobjVenda.iTipo = OPTION_CF
'    gobjVenda.objCupomFiscal.dtDataEmissao = Date
'    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iECF = giCodECF
    gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
'    gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
'    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
                                    
    Call Venda_AjustaTrib
    
    'se nao possui impressora fiscal
    If giCodModeloECF = 4 Then gobjVenda.objCupomFiscal.iStatus = STATUS_BAIXADO
    
'    'If gobjlojaecf.iImprimeItemAItem = DESMARCADO Then
'        lErro = Transforma_Cupom
'        If lErro <> SUCESSO Then gError 112075
'    'End If
'
    
    lErro = Informa_Meios_Pagto
    If lErro <> SUCESSO Then gError 109564

    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
'
    gobjGenerico.vVariavel = vbOK
    
    Unload Me
        
    giSaida = 1
        
    Exit Sub
        
Erro_BotaoAbrirGaveta_Click:
    
    Select Case gErr
    
        Case 99746
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_INSUFICIENTE, gErr)
        
        Case 99824, 109564, 109692, 112075, 126777, 207992
        
        Case 133759
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_JA_PAGO1, gErr)
        
        Case 133783
            Call Rotina_ErroECF(vbOKOnly, ERRO_BOTAO_AG_NAO_TEF, gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164167)

    End Select
        
    Exit Sub
     
End Sub

Private Function Informa_Meios_Pagto() As Long

Dim lErro As Long
Dim sIndice As String
Dim objMovCaixa As New ClassMovimentoCaixa
Dim sDescForma As String
Dim iTipo As Integer
Dim iIndice As Integer
Dim dValorTEF As Double
Dim colMeiosPag As New Collection
Dim dValorCC As Double
Dim dValorCD As Double
Dim dValorCarne As Double
Dim dValorCheque As Double
Dim dValorDin As Double
Dim dValorOutros As Double
Dim dValorTroca As Double
Dim dValorVT As Double
Dim dValor As String
Dim objMovCaixa1 As New ClassMovimentoCaixa
Dim objMovCaixa2 As New ClassMovimentoCaixa
Dim iNovoIndice As Integer
Dim iIndice2 As Integer
Dim objAux As New ClassMovimentoCaixa
Dim dMenor As Double
Dim dTotal As Double
Dim dDinheiro As Double

On Error GoTo Erro_Informa_Meios_Pagto
        
    
    'iNFORMAR PARA IMPRESSORA AS FORMAS PAGTO
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
        
        lErro = CF_ECF("Trata_MovCaixa", objMovCaixa, colMeiosPag)
        If lErro <> SUCESSO Then gError 133732
        
    Next
    
    'ordenar por valores...
    For iIndice = 1 To colMeiosPag.Count - 1
        Set objMovCaixa = colMeiosPag.Item(iIndice)
        dMenor = objMovCaixa.dValor
        iNovoIndice = iIndice
        For iIndice2 = iIndice + 1 To colMeiosPag.Count
            Set objMovCaixa1 = colMeiosPag.Item(iIndice2)
            If objMovCaixa1.dValor < dMenor Then
                dMenor = objMovCaixa1.dValor
                iNovoIndice = iIndice2
            End If
        Next
        Set objMovCaixa1 = colMeiosPag.Item(iNovoIndice)
        Call Inverte_Col(objAux, objMovCaixa)
        Call Inverte_Col(objMovCaixa, objMovCaixa1)
        Call Inverte_Col(objMovCaixa1, objAux)
    Next
    
    dTotal = 0
    
    For iIndice = 1 To colMeiosPag.Count
        Set objMovCaixa = colMeiosPag.Item(iIndice)
        dTotal = dTotal + objMovCaixa.dValor
        If dTotal >= StrParaDbl(APagar.Caption) And iIndice <> colMeiosPag.Count Then gError 112394
        If objMovCaixa.iTipo = TIPOMEIOPAGTOLOJA_DINHEIRO Then dDinheiro = objMovCaixa.dValor
    Next
        
    'se o troco que tem que ser dado ultrapassar a quantidade paga em dinheiro ==> erro
    'se deixar passar o ECF barra
    If dTotal > (StrParaDbl(APagar.Caption) + 0.0001) Then
    
        If (dDinheiro + 0.0001) < dTotal - StrParaDbl(APagar.Caption) Then gError 126594
        
    End If
        
    lErro = Executa_Fechamento_Cupom(dValorTEF, gobjVenda.objCupomFiscal.lNumero, colMeiosPag)
    If lErro <> SUCESSO Then gError 133710
        
    Informa_Meios_Pagto = SUCESSO
    
    Exit Function
        
Erro_Informa_Meios_Pagto:
    
    Informa_Meios_Pagto = gErr
    
    Select Case gErr
    
        Case 112394
            Call Rotina_ErroECF(vbOKOnly, ERRO_MEIOSPAG_ULTRAPASSAM, gErr)
            
        Case 126594
            Call Rotina_ErroECF(vbOKOnly, ERRO_TROCO_MAIOR_DINHEIRO, gErr)
            
        Case 133710, 133732
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164168)

    End Select
        
    Exit Function
        
End Function
    
'Private Function Transforma_Cupom() As Long
'
'Dim objItens As ClassItemCupomFiscal
'Dim lErro As Long
'Dim bAchou As Boolean
'Dim objProduto As ClassProduto
'Dim sProduto As String
'Dim lNum As Long
'Dim lNumero As Long
'Dim objAliquota As ClassAliquotaICMS
'Dim sAliquota As String
'
'On Error GoTo Erro_Transforma_Cupom
'
'    For Each objItens In gobjVenda.objCupomFiscal.colItens
'
'        ProdutoNomeRed.Text = objItens.sProduto
'
'        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, ProdutoNomeRed, objProduto)
'
'        'caso o produto não seja encontrado
'        If objProduto Is Nothing Then gError 99884
'
'        For Each objAliquota In gcolAliquotasTotal
'            If objAliquota.sSigla = objProduto.sICMSAliquota Then objItens.dAliquotaICMS = objAliquota.dAliquota
'        Next
'
'        If objItens.dAliquotaICMS > 0 Then
'            If objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
'                sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL & Format(objItens.dAliquotaICMS * 10000, "0000")
'            Else
'                sAliquota = Format(objItens.dAliquotaICMS * 10000, "0000")
'            End If
'        Else
'           sAliquota = objProduto.sSituacaoTribECF
'        End If
'
'        lErro = AFRAC_VenderItem(objProduto.sCodigo, objProduto.sDescricao, objItens.dQuantidade, CStr(Format(objItens.dPrecoUnitario, "standard")), 0, 0, 0, StrParaDbl(objItens.dPrecoUnitario * objItens.dQuantidade), sAliquota, objProduto.sSiglaUMVenda, False)
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Vender Item")
'        If lErro <> SUCESSO Then gError 99912
'
'    Next
'
'    Transforma_Cupom = SUCESSO
'
'    Exit Function
'
'Erro_Transforma_Cupom:
'
'    Transforma_Cupom = gErr
'
'    Select Case gErr
'
'        Case 99818, 99884, 99912
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164169)
'
'    End Select
'
'    Exit Function
'
'
'End Function
    
Private Sub Inverte_Col(objMovCaixa As ClassMovimentoCaixa, objMovCaixa1 As ClassMovimentoCaixa)
    
    objMovCaixa.dHora = objMovCaixa1.dHora
    objMovCaixa.dtDataMovimento = objMovCaixa1.dtDataMovimento
    objMovCaixa.dValor = objMovCaixa1.dValor
    objMovCaixa.iAdmMeioPagto = objMovCaixa1.iAdmMeioPagto
    objMovCaixa.iCaixa = objMovCaixa1.iCaixa
    objMovCaixa.iCodConta = objMovCaixa1.iCodConta
    objMovCaixa.iCodOperador = objMovCaixa1.iCodOperador
    objMovCaixa.iExcluiu = objMovCaixa1.iExcluiu
    objMovCaixa.iFilialEmpresa = objMovCaixa1.iFilialEmpresa
    objMovCaixa.iGerente = objMovCaixa1.iGerente
    objMovCaixa.iParcelamento = objMovCaixa1.iParcelamento
    objMovCaixa.iQuantLog = objMovCaixa1.iQuantLog
    objMovCaixa.iTipo = objMovCaixa1.iTipo
    objMovCaixa.iTipoCartao = objMovCaixa1.iTipoCartao
    objMovCaixa.lCupomFiscal = objMovCaixa1.lCupomFiscal
    objMovCaixa.lMovtoEstorno = objMovCaixa1.lMovtoEstorno
    objMovCaixa.lMovtoTransf = objMovCaixa1.lMovtoTransf
    objMovCaixa.lNumero = objMovCaixa1.lNumero
    objMovCaixa.lNumIntDocLog = objMovCaixa1.lNumIntDocLog
    objMovCaixa.lNumIntExt = objMovCaixa1.lNumIntExt
    objMovCaixa.lNumMovto = objMovCaixa1.lNumMovto
    objMovCaixa.lNumRefInterna = objMovCaixa1.lNumRefInterna
    objMovCaixa.lSequencial = objMovCaixa1.lSequencial
    objMovCaixa.lSequencialConta = objMovCaixa1.lSequencialConta
    objMovCaixa.lTransferencia = objMovCaixa1.lTransferencia
    objMovCaixa.sFavorecido = objMovCaixa1.sFavorecido
    objMovCaixa.sHistorico = objMovCaixa1.sHistorico
    objMovCaixa.iIndiceImpChq = objMovCaixa1.iIndiceImpChq
    
End Sub

'Private Sub Desmembrar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Desmembrar_Click
'
'    lErro = CF_ECF("Requisito_XXII")
'    If lErro <> SUCESSO Then gError 207993
'
'   lErro = CF_ECF("Desmembrar_ECF", gobjVenda)
'    If lErro <> SUCESSO Then gError 109811
'
'    gcolVendas.Add gobjVenda
'
'    Call Traz_Dados_Tela
'
'    Exit Sub
'
'Erro_Desmembrar_Click:
'
'    Select Case gErr
'
'        Case 109811, 207993
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164170)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Function Executa_Fechamento_Cupom(ByVal dValorTEF As Double, ByVal lCOO As Long, ByVal colMeiosPag As Collection) As Long

Dim lErro As Long
Dim objMovCaixa As ClassMovimentoCaixa
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim sMsg As String
Dim sMsgVendedor As String
Dim objTroca As ClassTroca
Dim objVendedor As ClassVendedor
Dim iRetorno As Integer
Dim objCheque As New ClassChequePre
Dim objCliente As ClassCliente
Dim sCPF As String
Dim lNumero As Long
Dim objFormMsg As Object
Dim lTamanho As Long
Dim sRetorno As String
Dim objTela As Object
Dim sDescricao As String
Dim sNomeArq1 As String
Dim sNomeArq2 As String
Dim iIndiceChq As Integer
Dim sOrcamento As String



On Error GoTo Erro_Executa_Fechamento_Cupom

    Set objTela = Me

    iIndiceChq = 0

    Parent.MousePointer = vbHourglass

    lErro = CF_ECF("ECF_AcrescimoDescontoCupom", CStr(Format(gobjVenda.objCupomFiscal.dValorAcrescimo, "standard")), CStr(Format(gobjVenda.objCupomFiscal.dValorDesconto + gobjVenda.objCupomFiscal.dValorDesconto1 + gobjVenda.objCupomFiscal.dValorDescontoTEF, "standard")), "", "")
    If lErro <> SUCESSO Then gError 109691

    'iNFORMAR PARA IMPRESSORA AS FORMAS PAGTO
    For Each objMovCaixa In colMeiosPag
    
        'se for diferente de TEF ==> imprime primeiro
        If objMovCaixa.iTipo <> TIPOMEIOPAGTOLOJA_TEF And objMovCaixa.iTipo <> TIPOMEIOPAGTOLOJA_TEFCHQ Then
        
            For Each objTiposMeiosPagtos In gcolTiposMeiosPagtos
                'Se for do tipo procurado
                If objTiposMeiosPagtos.iTipo = objMovCaixa.iTipo Then
                
                    If objMovCaixa.iTipo = TIPOMEIOPAGTOLOJA_TROCA Then
                    
                        If gobjVenda.colTroca.Count > 0 Then
                            sMsg = "Produtos:"
                        
                            For Each objTroca In gobjVenda.colTroca
                            
                                sMsg = sMsg & " " & objTroca.dQuantidade & " " & objTroca.sProduto & "/"
                            
                            Next
                
                            sMsg = left(sMsg, Len(sMsg) - 1)
                
                        End If
                
                    End If
                    
                    lErro = CF_ECF("ECF_FormaPagamento", objTiposMeiosPagtos.sDescricao, objTiposMeiosPagtos.iTipo, Format(objMovCaixa.dValor, "standard"), sMsg)
                    If lErro <> SUCESSO Then gError 99821
                    
                    sDescricao = objTiposMeiosPagtos.sDescricao
                
                    sMsg = ""
                
                End If
                
            Next
            
        End If
        
    Next
        
    'iNFORMAR PARA IMPRESSORA AS FORMAS PAGTO
    For Each objMovCaixa In colMeiosPag
            
        If objMovCaixa.iTipo = TIPOMEIOPAGTOLOJA_TEF Then
        
            '**** ALTERADO DE TEF PARA Cartão Crédito pois nao havia mas espaco para programar TEF  *** MARIO ****
            'informa a forma de pagamento passando os seus dados
'            lErro = CF_ECF("ECF_FormaPagamento", "Cartão Crédito", TIPOMEIOPAGTOLOJA_TEF, Format(objMovCaixa.dValor + gobjVenda.objCupomFiscal.dValorTrocoTEF - gobjVenda.objCupomFiscal.dValorDescontoTEF, "standard"), sMsg)
            lErro = CF_ECF("ECF_FormaPagamento", "TEF", TIPOMEIOPAGTOLOJA_TEF, Format(objMovCaixa.dValor + gobjVenda.objCupomFiscal.dValorTrocoTEF - gobjVenda.objCupomFiscal.dValorDescontoTEF, "standard"), sMsg)
            If lErro <> SUCESSO Then gError 133734
                
            objMovCaixa.sVinculado = "TEF"
                
            gobjVenda.colVinculado.Add objMovCaixa
            
        End If
                
        If objMovCaixa.iTipo = TIPOMEIOPAGTOLOJA_TEFCHQ Then
                
            iIndiceChq = iIndiceChq + 1
    
            'informa a forma de pagamento passando os seus dados
            lErro = CF_ECF("ECF_FormaPagamento", "TEFCHQ" & CStr(iIndiceChq), TIPOMEIOPAGTOLOJA_TEF + iIndiceChq, Format(objMovCaixa.dValor, "standard"), sMsg)
            If lErro <> SUCESSO Then gError 126593
                
            objMovCaixa.sVinculado = "TEFCHQ" & CStr(iIndiceChq)
                
            gobjVenda.colVinculado.Add objMovCaixa
                
        End If
            
    Next
    
    'se for a Elgin
    If giCodModeloECF = 7 Then

        'Recolhe a descrição
        For Each objTiposMeiosPagtos In gcolTiposMeiosPagtos
            'Se o tipo for dinheiro
            If objTiposMeiosPagtos.iTipo = StrParaInt(TIPOMEIOPAGTOLOJA_DINHEIRO) Then
                Exit For
            End If
        Next


        'informa a forma de pagamento passando os seus dados
        lErro = CF_ECF("ECF_FormaPagamento", objTiposMeiosPagtos.sDescricao, objTiposMeiosPagtos.iTipo, Format(0, "standard"), sMsg)
        If lErro <> SUCESSO Then gError 126593
    
    End If
    
    'Verifica se existe carnê
    If gobjVenda.objCarne.colParcelas.Count > 0 Then
        'Gera o Código do Carnê
        gobjVenda.objCarne.sCodBarrasCarne = FormataCpoNum(giFilialEmpresa, 5) & FormataCpoNum(giCodCaixa, 5) & FormataCpoNum(gobjVenda.objCupomFiscal.lNumero, 10)
    End If
    
    'Abri a Gaveta
    lErro = CF_ECF("ECF_AbrirGaveta")
    If lErro <> SUCESSO Then gError 99823
            
    If gobjVenda.objCupomFiscal.iVendedor <> 0 Then

        Set objVendedor = New ClassVendedor
                
        lErro = CF_ECF("Vendedores_Le_Codigo", gobjVenda.objCupomFiscal.iVendedor, objVendedor)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214949
    
        sMsgVendedor = TRACO_CAB
        sMsgVendedor = sMsgVendedor & VENDEDOR_ECF_MSG & Formata_Campo(ALINHAMENTO_DIREITA, 38, " ", gobjVenda.objCupomFiscal.iVendedor & " - " & objVendedor.sNomeReduzido)
        If Len(Trim(gobjVenda.objCupomFiscal.sNomeCliente)) > 0 Then sMsgVendedor = sMsgVendedor & "Cliente: " & Formata_Campo(ALINHAMENTO_DIREITA, 39, " ", gobjVenda.objCupomFiscal.sNomeCliente)
        sMsgVendedor = sMsgVendedor & TRACO_CAB

    End If
            
    If gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then
            
        If gobjVenda.objCupomFiscal.lNumeroDAV = 0 Then
            sOrcamento = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "PV""" & Format(gobjVenda.objCupomFiscal.lNumOrcamento, "0000000000") & """ ")
        Else
            sOrcamento = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "DAV""" & Format(gobjVenda.objCupomFiscal.lNumeroDAV, "0000000000") & """ ")
        End If
            
    End If
                
    If Not (gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF)) Then
    
        'Fecha cupom
        Timer1.Enabled = False
        lErro = CF_ECF("ECF_FecharCupom", objTela, gobjVenda, False, gobjVenda.objCupomFiscal.sCPFCGC1, gobjVenda.objCupomFiscal.sNomeCliente, gobjVenda.objCupomFiscal.sEndereco, False, sOrcamento, sMsgVendedor)
        Timer1.Enabled = True
        If lErro <> SUCESSO Then gError 99822
        
    End If
        
    'imprimie cartado de credito e debito POS. Foi colocado nesta posicao pois teoricamente
    'poderia ter venda POS e TEF no mesmo cupom.
    lErro = CF_ECF("CCD_Imprime", colMeiosPag, gobjVenda)
    If lErro <> SUCESSO Then gError 214140
        
    If Not (gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF)) Then
    
        lTamanho = 10
        sRetorno = String(lTamanho, 0)
        
        'Indica o status do TEF
        Call GetPrivateProfileString(APLICACAO_ECF, "StatusTEF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
            
        sRetorno = StringZ(sRetorno)
        
        If sRetorno = TEF_STATUS_VENDA Then
            
            Set objFormMsg = MsgTEF
            
                '**** ALTERADO DE TEF PARA Outros pois nao havia mas espaco para programar TEF  *** MARIO ****
            sDescricao = "TEF"
    '        sDescricao = "Cartão Crédito"
            
            'se dValorTEF for zero significa que só tem cancelamentos ==> portanto vai imprimir no gerencial ja que nao tem forma de pagamento TEF
            If dValorTEF > 0 Then
            
    '        lErro = CF_ECF("TEF_Imprime", sDescricao, dValorTEF, lCOO, objFormMsg, objTela, gobjVenda)
                lErro = CF_ECF("TEF_Imprime_PAYGO", sDescricao, dValorTEF + gobjVenda.objCupomFiscal.dValorTrocoTEF - gobjVenda.objCupomFiscal.dValorDescontoTEF, lCOO, objFormMsg, objTela, gobjVenda)
                If lErro <> SUCESSO Then gError 133719
                
            Else
            
                lErro = CF_ECF("TEF_Imprime_CNC_PAYGO", objFormMsg, objTela, gobjVenda)
                If lErro <> SUCESSO Then gError 214564
            
            End If
            
        End If
        
    End If
        
    'incluido para tratar saque de TEF (quando o valor tirado do cartao é maior que o valor a pagar, o residuo vai sair como troco
    Troco.Caption = Format(StrParaDbl(Troco.Caption) + gobjVenda.objCupomFiscal.dValorTrocoTEF, "Standard")
    
    Call Troco_Tela
    
    If gobjVenda.iTipo = OPTION_CF And AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        'Fecha cupom
        Timer1.Enabled = False
        lErro = CF_ECF("ECF_FecharCupom", objTela, gobjVenda, False, gobjVenda.objCupomFiscal.sCPFCGC1, gobjVenda.objCupomFiscal.sNomeCliente, gobjVenda.objCupomFiscal.sEndereco, False, sOrcamento, sMsgVendedor)
        Timer1.Enabled = True
        If lErro <> SUCESSO Then gError 99822
    
    Else
        
        'Realiza as operações necessárias para gravar
        lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
        If lErro <> SUCESSO Then gError 99824
    
    End If
    
    'Atualiza os Movimentos nas coleções globais
    Call CF_ECF("Atualiza_Movimentos_Memoria", gobjVenda)
    
    'Atribui para a coleção global o objvenda
    gcolVendas.Add gobjVenda
        
    'Jogo todos os cheques na col global
    For Each objCheque In gobjVenda.colCheques
    
        'Atualiza o saldos de cheques
        gdSaldocheques = gdSaldocheques + objCheque.dValor
        'Adiciona os cheques na coleção global
        gcolCheque.Add objCheque
    
    Next
    
    'Para cada movimento da venda
'??? 24/08/2016     For Each objMovCaixa In gobjVenda.colMovimentosCaixa
'??? 24/08/2016         'Se for de cartao de crédito ou débito especificado
'??? 24/08/2016         If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro + objMovCaixa.dValor
'??? 24/08/2016     Next

    'Atualiza o arquivo
    Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)
        
    If gobjLojaECF.iAbreAposFechamento = MARCADO Then
    
        sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
        lErro = CF_ECF("Abre_Cupom", gobjVenda)
        If lErro <> SUCESSO Then gError 99818
        
'        gobjVenda.objCupomFiscal.lNumero = lNumero
    End If

    Parent.MousePointer = vbDefault

    Executa_Fechamento_Cupom = SUCESSO
    
    Exit Function
    
Erro_Executa_Fechamento_Cupom:

    Executa_Fechamento_Cupom = gErr

    Parent.MousePointer = vbDefault

    Select Case gErr
    
        
        Case 99821 To 99824, 99818, 109691, 126593, 133711, 133719, 133734, 133805, 133855, 204348, 204349, 214140, 214564, 214611
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 654321)
   
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164171)

    End Select
        
'    lTamanho = 10
'    sRetorno = String(lTamanho, 0)
'
'    'Indica o status do TEF
'    Call GetPrivateProfileString(APLICACAO_ECF, "StatusTEF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'
'    sRetorno = StringZ(sRetorno)
'
'    If sRetorno = TEF_STATUS_VENDA Then
'
'        Call CF_ECF("TEF_NaoConfirma_Transacao1", objTela, gobjVenda)
'
'        'cancela os cartoes ja confirmados e nao confirma o ultimo
'        Call CF_ECF("TEF_CNC", gobjVenda, objFormMsg, objTela)
'
''        Call CF_ECF("TEF_Imprime_CNC", gobjVenda, objFormMsg, objTela)
'
'    End If
'
'    lTamanho = 10
'    sRetorno = String(lTamanho, 0)
'
'    'Indica o status do TEF
'    Call GetPrivateProfileString(APLICACAO_ECF, "CupomAberto", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'
'    sRetorno = StringZ(sRetorno)
'
'    If sRetorno <> "0" Then
'
'        'Retorna para a tela de venda a informação de cancelamento da venda
'        gobjGenerico.vVariavel = vbCancel
'
'    Else
'
'        gobjGenerico.vVariavel = vbOK
'
'    End If
'
'    'Limpa os dados da Venda
'    Set gobjVenda = New ClassVenda
'
'
'    'Fecha a tela
'    Unload Me
'
'    giSaida = 1
        
    Exit Function

End Function

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

