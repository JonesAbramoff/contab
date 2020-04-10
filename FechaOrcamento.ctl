VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl FechaOrcamento 
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   ScaleHeight     =   8745
   ScaleMode       =   0  'User
   ScaleWidth      =   12593.21
   Begin MSComDlg.CommonDialog CD1 
      Left            =   105
      Top             =   1125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   300
      TabIndex        =   107
      Top             =   210
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FechaOrcamento.ctx":0000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pagamento"
      Height          =   2505
      Index           =   0
      Left            =   360
      TabIndex        =   77
      Top             =   1680
      Width           =   8565
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   11160
         Top             =   -360
      End
      Begin VB.CommandButton BotaoTroco 
         Caption         =   "(F11)  Troco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6180
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1920
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
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   1710
      End
      Begin MSMask.MaskEdBox DescontoValor 
         Height          =   345
         Left            =   1590
         TabIndex        =   17
         Top             =   645
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
         Left            =   3285
         TabIndex        =   18
         Top             =   645
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
      Begin MSMask.MaskEdBox AcrescimoValor 
         Height          =   345
         Left            =   1575
         TabIndex        =   23
         Top             =   1515
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
      Begin MSMask.MaskEdBox AcrescimoPerc 
         Height          =   345
         Left            =   3285
         TabIndex        =   24
         Top             =   1515
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
      Begin MSMask.MaskEdBox ProdutoNomeRed 
         Height          =   330
         Left            =   3480
         TabIndex        =   80
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
      Begin MSMask.MaskEdBox DescontoValor1 
         Height          =   345
         Left            =   1575
         TabIndex        =   20
         Top             =   1080
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
         Left            =   3285
         TabIndex        =   21
         Top             =   1080
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
         Left            =   4125
         TabIndex        =   106
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Left            =   75
         TabIndex        =   19
         Top             =   1095
         Width           =   1470
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
         TabIndex        =   94
         Top             =   -960
         Width           =   6555
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
         TabIndex        =   93
         Top             =   -960
         Width           =   2715
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "&Acrésci&mo:"
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
         TabIndex        =   22
         Top             =   1530
         Width           =   1335
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
         Left            =   4140
         TabIndex        =   92
         Top             =   1530
         Width           =   240
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
         TabIndex        =   91
         Top             =   870
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
         TabIndex        =   90
         Top             =   885
         Width           =   720
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
         TabIndex        =   89
         Top             =   345
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
         TabIndex        =   88
         Top             =   360
         Width           =   705
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
         Left            =   4125
         TabIndex        =   87
         Top             =   645
         Width           =   240
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
         Left            =   1560
         TabIndex        =   86
         Top             =   1950
         Width           =   2535
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
         Left            =   480
         TabIndex        =   85
         Top             =   1965
         Width           =   1065
      End
      Begin VB.Label Label5 
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
         Left            =   75
         TabIndex        =   16
         Top             =   645
         Width           =   1470
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
         Left            =   5580
         TabIndex        =   84
         Top             =   1410
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
         TabIndex        =   83
         Top             =   1425
         Width           =   780
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
         Left            =   1575
         TabIndex        =   82
         Top             =   225
         Width           =   2520
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   855
         TabIndex        =   81
         Top             =   225
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Meios de Pagamento"
      Height          =   4065
      Left            =   360
      TabIndex        =   39
      Top             =   4320
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
         Index           =   16
         Left            =   5295
         TabIndex        =   101
         Text            =   "(F4)"
         Top             =   960
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
         Index           =   15
         Left            =   5310
         TabIndex        =   100
         Text            =   "(F5)"
         Top             =   1380
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
         Index           =   14
         Left            =   5295
         TabIndex        =   99
         Text            =   "(F6)"
         Top             =   1800
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
         Index           =   13
         Left            =   5310
         TabIndex        =   98
         Text            =   "(F7)"
         Top             =   2235
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
         Index           =   12
         Left            =   5295
         TabIndex        =   97
         Text            =   "(F8)"
         Top             =   2655
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
         Index           =   11
         Left            =   5295
         TabIndex        =   96
         Text            =   "(F10)"
         Top             =   3495
         Width           =   480
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
         Left            =   5310
         TabIndex        =   95
         Text            =   "(F9)"
         Top             =   3105
         Width           =   330
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
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3390
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   825
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
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1245
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
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2115
         Width           =   2400
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1680
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
         Index           =   1
         Left            =   5295
         TabIndex        =   48
         Text            =   "(F4)"
         Top             =   870
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
         Index           =   2
         Left            =   5310
         TabIndex        =   47
         Text            =   "(F5)"
         Top             =   1290
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
         Index           =   3
         Left            =   5295
         TabIndex        =   46
         Text            =   "(F6)"
         Top             =   1710
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
         Index           =   0
         Left            =   5310
         TabIndex        =   45
         Text            =   "(F7)"
         Top             =   2145
         Width           =   345
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
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2535
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
         TabIndex        =   43
         Text            =   "(F8)"
         Top             =   2565
         Width           =   330
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
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2970
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
         Index           =   10
         Left            =   5295
         TabIndex        =   41
         Text            =   "(F10)"
         Top             =   3405
         Width           =   480
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
         Index           =   5
         Left            =   5310
         TabIndex        =   40
         Text            =   "(F9)"
         Top             =   3015
         Width           =   330
      End
      Begin MSMask.MaskEdBox MaskCheques 
         Height          =   345
         Left            =   2235
         TabIndex        =   3
         Top             =   885
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
         Left            =   2235
         TabIndex        =   5
         Top             =   1305
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
         Left            =   2235
         TabIndex        =   7
         Top             =   1710
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
         Left            =   2235
         TabIndex        =   11
         Top             =   3390
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
         Left            =   2235
         TabIndex        =   9
         Top             =   2940
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
         Left            =   2235
         TabIndex        =   1
         Top             =   480
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
         Index           =   24
         Left            =   1020
         TabIndex        =   0
         Top             =   450
         Width           =   1170
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
         Index           =   23
         Left            =   960
         TabIndex        =   2
         Top             =   870
         Width           =   1230
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
         Index           =   22
         Left            =   270
         TabIndex        =   4
         Top             =   1290
         Width           =   1920
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
         Index           =   21
         Left            =   345
         TabIndex        =   6
         Top             =   1725
         Width           =   1845
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
         Left            =   1305
         TabIndex        =   76
         Top             =   2160
         Width           =   885
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
         Left            =   1215
         TabIndex        =   10
         Top             =   3435
         Width           =   975
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
         TabIndex        =   75
         Top             =   1290
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
         Index           =   12
         Left            =   4635
         TabIndex        =   74
         Top             =   870
         Width           =   165
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
         Left            =   1350
         TabIndex        =   73
         Top             =   2580
         Width           =   840
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
         Left            =   1305
         TabIndex        =   8
         Top             =   2985
         Width           =   885
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
         TabIndex        =   72
         Top             =   3450
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
         TabIndex        =   71
         Top             =   3015
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
         TabIndex        =   70
         Top             =   1755
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
         Index           =   9
         Left            =   8130
         TabIndex        =   69
         Top             =   2985
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
         TabIndex        =   68
         Top             =   2970
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
         Index           =   4
         Left            =   8130
         TabIndex        =   67
         Top             =   2580
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
         TabIndex        =   66
         Top             =   2580
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
         TabIndex        =   65
         Top             =   450
         Width           =   165
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
         TabIndex        =   64
         Top             =   435
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
         TabIndex        =   63
         Top             =   840
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
         TabIndex        =   62
         Top             =   1275
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
         TabIndex        =   61
         Top             =   2130
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
         TabIndex        =   60
         Top             =   1710
         Width           =   2010
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
         TabIndex        =   59
         Top             =   3420
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
         Index           =   18
         Left            =   8130
         TabIndex        =   58
         Top             =   870
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
         TabIndex        =   57
         Top             =   1290
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
         TabIndex        =   56
         Top             =   3435
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
         TabIndex        =   55
         Top             =   2160
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
         Index           =   14
         Left            =   8130
         TabIndex        =   54
         Top             =   1725
         Width           =   165
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10800
      Top             =   600
   End
   Begin VB.PictureBox Picture2 
      Height          =   3240
      Left            =   9120
      ScaleHeight     =   3180
      ScaleWidth      =   2355
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   960
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
         Height          =   210
         Index           =   4
         Left            =   345
         TabIndex        =   108
         TabStop         =   0   'False
         Text            =   "(F11)"
         Top             =   285
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
         Index           =   19
         Left            =   345
         TabIndex        =   104
         TabStop         =   0   'False
         Text            =   "(F12)"
         Top             =   1785
         Width           =   615
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
         Index           =   18
         Left            =   75
         TabIndex        =   103
         TabStop         =   0   'False
         Text            =   "(Esc)"
         Top             =   2880
         Visible         =   0   'False
         Width           =   540
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
         Index           =   17
         Left            =   375
         TabIndex        =   102
         TabStop         =   0   'False
         Text            =   "(F3)"
         Top             =   1035
         Width           =   540
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   450
         Left            =   510
         Picture         =   "FechaOrcamento.ctx":008B
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Excluir"
         Top             =   2925
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   450
         Left            =   240
         Picture         =   "FechaOrcamento.ctx":0215
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Gravar"
         Top             =   900
         Width           =   1920
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   450
         Left            =   240
         Picture         =   "FechaOrcamento.ctx":036F
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir"
         Top             =   1650
         Width           =   1920
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   450
         Left            =   255
         Picture         =   "FechaOrcamento.ctx":0471
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Limpar"
         Top             =   165
         Width           =   1920
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
         Left            =   360
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "(F2)"
         Top             =   2490
         Width           =   360
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   435
         Left            =   240
         Picture         =   "FechaOrcamento.ctx":09A3
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Fechar"
         Top             =   2370
         Width           =   1920
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Validade"
      Height          =   735
      Left            =   360
      TabIndex        =   31
      Top             =   840
      Width           =   8565
      Begin VB.CommandButton BotaoLimparData 
         Height          =   315
         Left            =   3165
         Picture         =   "FechaOrcamento.ctx":0B21
         Style           =   1  'Graphical
         TabIndex        =   105
         TabStop         =   0   'False
         ToolTipText     =   "Limpar"
         Top             =   300
         Width           =   330
      End
      Begin MSComCtl2.UpDown UpDownDataFinal 
         Height          =   315
         Left            =   5010
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   4035
         TabIndex        =   13
         Top             =   270
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Duracao 
         Height          =   300
         Left            =   6585
         TabIndex        =   15
         Top             =   270
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipo 
         Caption         =   "DAV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   330
         TabIndex        =   109
         Top             =   330
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   7065
         TabIndex        =   34
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D&uração:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   5745
         TabIndex        =   14
         Top             =   315
         Width           =   795
      End
      Begin VB.Label LabelEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2160
         TabIndex        =   33
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   1365
         TabIndex        =   32
         Top             =   330
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   12
         Top             =   315
         Width           =   360
      End
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
      TabIndex        =   38
      Top             =   120
      Width           =   2715
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
      TabIndex        =   37
      Top             =   135
      Width           =   6555
   End
End
Attribute VB_Name = "FechaOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'Property Variables:
Dim m_Caption As String
Event Unload()

'Global
Dim gobjVenda As ClassVenda
Dim gdDescontoAnterior As Double
Dim gdDescontoAnterior1 As Double
Dim gdAcrescimoAnterior As Double
Dim gdPercDescontoAnterior As Double
Dim gdPercDescontoAnterior1 As Double
Dim gdPercAcrescimoAnterior As Double
Dim giSaida As Integer
Dim gobjGenerico As AdmGenerico

'??? 24/08/2016 Dim gdSaldoDinheiroAnterior As Double
Dim gdSaldoChequesAnterior As Double
Dim gdSaldoCartaoDebitoAnterior As Double
Dim gdSaldoCartaoCreditoAnterior As Double
Dim gdSaldoOutrosAnterior As Double
Dim gdSaldoTicketAnterior As Double

Public Property Get objFalta() As Object
     Set objFalta = Falta
End Property

Public Property Get objPago() As Object
     Set objPago = Pago
End Property

Public Property Get objLabelEmissao() As Object
     Set objLabelEmissao = LabelEmissao
End Property

Public Property Get objTroco() As Object
     Set objTroco = Troco
End Property

Public Property Get objDuracao() As Object
     Set objDuracao = Duracao
End Property

Public Property Get objAPagar() As Object
     Set objAPagar = APagar
End Property

Public Property Get objDataFinal() As Object
     Set objDataFinal = DataFinal
End Property


Function Trata_Parametros(objVenda As ClassVenda, objGenerico As AdmGenerico) As Long
        
    Set gobjVenda = objVenda
    
    Set gobjGenerico = objGenerico
    
    gobjGenerico.vVariavel = vbAbort
    
    'Joga o Valor Total do Cupom Fiscal (Formatado)
    Total.Caption = Format(gobjVenda.objCupomFiscal.dValorProdutos, "standard")
    
    Call Traz_Dados_Tela
    
    'Chama a Calcula Valores
    Call Recalcula_Valores1
    
    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()
        
Dim objTela As Object
        
        
    Call Timer2_Timer
    
    giSaida = 0
    
    Apresentacao1.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
        
    UserControl.Parent.WindowState = 2
    
    If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
        LabelTipo.Visible = True
    End If
    
    Call Inicializa_Valores
    
    If giDinheiroAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskDinheiro.Enabled = False
        Label1(24).Enabled = False
    End If
        
    If giChequeAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskCheques.Enabled = False
        Label1(23).Enabled = False
        BotaoCheques.Enabled = False
    End If
        
    If giCartaoCreditoAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskCartaoCredito.Enabled = False
        Label1(22).Enabled = False
        BotaoCartaoCredito.Enabled = False
    End If
        
    If giCartaoDebitoAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        MaskCartaoDebito.Enabled = False
        Label1(21).Enabled = False
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
    
    Set objTela = Me
    
    Call CF_ECF("Inicializa_FechaOrcamento", objTela)
    
    lErro_Chama_Tela = SUCESSO

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
    Outros.Caption = Format(0, "Standard")
    Carne.Caption = Format(0, "Standard")
    Troca.Caption = Format(0, "Standard")
    Outros.Caption = Format(0, "Standard")
    BotaoCheques.Caption = Format(0, "Standard")
    BotaoCarne.Caption = Format(0, "Standard")
    BotaoTroca.Caption = Format(0, "Standard")
    BotaoOutros.Caption = Format(0, "Standard")
    BotaoTicket.Caption = Format(0, "Standard")
    BotaoCartaoCredito.Caption = Format(0, "Standard")
    BotaoCartaoDebito.Caption = Format(0, "Standard")
    LabelEmissao.Caption = Format(Date, "dd/mm/yyyy")
    
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
        If (objMovimento.iTipo = iTipo And (objMovimento.iAdmMeioPagto = 0 Or iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO)) Then
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

    'Se tiver valor a tribuir e o movimento não foi encontrado
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
            objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
        End If
        
        'Adiciona o novo movimento à coleção global da tela
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
        
    End If
    
    Exit Sub
    
End Sub

Private Sub BotaoLimparData_Click()
    LabelEmissao.Caption = Format(Date, "dd/mm/yyyy")
    Duracao.Text = StrParaDate(DataFinal.Text) - StrParaDate(LabelEmissao.Caption)
    gobjVenda.objCupomFiscal.dHoraEmissao = 0
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

Private Function ConfirmaDesconto() As Long

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_ConfirmaDesconto
    
    For Each objOperador In gcolOperadores

        If objOperador.iCodigo = giCodOperador Then

            objOperador.iLimiteDesconto = objOperador.iLimiteDesconto

            Exit For

        End If

    Next
    
    'Se for necessária a autorização do Gerente
    If objOperador.iLimiteDesconto <> 100 Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    ConfirmaDesconto = SUCESSO
    
    Exit Function
    
Erro_ConfirmaDesconto:
    
    ConfirmaDesconto = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175707)

    End Select
        
    Exit Function
    
End Function

Private Sub DescontoValor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DescontoValor_Validate

    'Se o valor foi preenchido
    If Len(Trim(DescontoValor.Text)) > 0 Then
        
        'Verifica se é um valor aceito
        lErro = Valor_NaoNegativo_Critica(DescontoValor.Text)
        If lErro <> SUCESSO Then gError 105755
    
        If ConfirmaDesconto <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    'Não permite desconto maior que o total a pagar.
    If (StrParaDbl(DescontoValor.Text) + StrParaDbl(DescontoValor1.Text)) - (StrParaDbl(Total.Caption) - StrParaDbl(Troca.Caption)) > 0.0001 Then
        gError 105756
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
        
        Case 105755, ERRO_SEM_MENSAGEM
        
        Case 105756
            Call Rotina_ErroECF(vbOKOnly, ERRO_DESCONTO_MAIOR, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160191)

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
        If lErro <> SUCESSO Then gError 126747
        
        If ConfirmaDesconto <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
        
    'Não permite desconto maior que o total a pagar menos a troca.
    If (StrParaDbl(DescontoValor.Text) + StrParaDbl(DescontoValor1.Text)) - (StrParaDbl(Total.Caption) - StrParaDbl(Troca.Caption)) > 0.0001 Then
        gError 126748
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
        
        Case 126747, ERRO_SEM_MENSAGEM
        
        Case 126748
            Call Rotina_ErroECF(vbOKOnly, ERRO_DESCONTO_MAIOR, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160192)

    End Select

    Exit Sub
    
End Sub

Private Sub DescontoPerc_GotFocus()
    
    gdPercDescontoAnterior = StrParaDbl(DescontoPerc.Text)
    
End Sub

Private Sub DescontoPerc1_GotFocus()
    
    gdPercDescontoAnterior1 = StrParaDbl(DescontoPerc1.Text)
    
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
            If lErro <> SUCESSO Then gError 105759
        
            If ConfirmaDesconto <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If
        
        DescontoValor.Text = ""
        
        'Exibe o novo valor formatado
        If StrParaDbl(DescontoPerc.Text) > 0 Then
            DescontoPerc.Text = Round(StrParaDbl(DescontoPerc.Text), 2)
            dPercentDesc = CDbl(DescontoPerc.Text)
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
        
        Case 105759, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160193)

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
            If lErro <> SUCESSO Then gError 126749
            
            If ConfirmaDesconto() <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
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
        
        Case 126749, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160194)

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
        If lErro <> SUCESSO Then gError 105757
    End If
        
    'Se o Acrescimo dado é maior que o total
    If StrParaDbl(AcrescimoValor.Text) - StrParaDbl(Total.Caption) > 0.0001 Then
        gError 105758
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
        
        Case 105757
        
        Case 105758
            Call Rotina_ErroECF(vbOKOnly, ERRO_ACRESCIMO_MAIOR, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160195)

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
        
        'Coloca o valor formatado na tela
        If StrParaDbl(AcrescimoPerc.Text) > 0 Then
            AcrescimoPerc.Text = Format(AcrescimoPerc.Text, "0.000")
            dAcrescimoPerc = CDbl(AcrescimoPerc.Text)
            AcrescimoValor.Text = (dAcrescimoPerc / 100) * StrParaDbl(Total.Caption)
        Else
            AcrescimoValor.Text = Format(0, "standard")
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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160196)

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
        If lErro <> SUCESSO Then gError 99722
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
        
        Case 99722
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160197)

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
        If lErro <> SUCESSO Then gError 99723
        
    End If
        
    'Se o valor nesse campo nao foi alterado ==> Sai
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
                If lErro = 0 Then gError 105779
                
                'Adiciona o cheque  na coleção da venda
                gobjVenda.colCheques.Add objCheque
                        
                'criar movimento para o cheque
                Set objMovCaixa = New ClassMovimentoCaixa
            
                'Preenche o novo movcaixa
                objMovCaixa.iFilialEmpresa = giFilialEmpresa
                objMovCaixa.iCaixa = giCodCaixa
                objMovCaixa.iCodOperador = giCodOperador
                objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CHEQUE
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
        
        Case 99723
        
        Case 105779
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_CAIXA, "NumProxCheque", NOME_ARQUIVO_CAIXA)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160198)

    End Select

    Exit Sub
    
End Sub


Private Sub MaskCartaoCredito_GotFocus()
    
    'Posiciona o cursor no início
    Call MaskEdBox_TrataGotFocus(MaskCartaoCredito)
    
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
        
        'Verifica se é um valor válido
        lErro = Valor_NaoNegativo_Critica(MaskCartaoDebito.Text)
        If lErro <> SUCESSO Then gError 99724
        
    End If
                
    
    'Se o valor nao foi alterado ==> Sai
'    If gdSaldoCartaoDebitoAnterior <> StrParaDbl(MaskCartaoDebito.Text) Then
        
        'Exibe o valor formatado na tela
        CartaoDebito.Caption = Format(StrParaDbl(BotaoCartaoDebito.Caption) + StrParaDbl(MaskCartaoDebito.Text), "Standard")
        
        'Recalcula os valores
        Call Recalcula_Valores2
        
        'Inclui o movimento
        Call Inclui_Movimento(StrParaDbl(MaskCartaoDebito.Text), MOVIMENTOCAIXA_RECEB_CARTAODEBITO, TIPO_MANUAL)
    
        If Len(Trim(MaskCartaoDebito.Text)) > 0 Then MaskCartaoDebito.Text = Round(StrParaDbl(MaskCartaoDebito.Text), 2)
    
        'guarda o valor atual em cartão Débito nao especificado
        gdSaldoCartaoDebitoAnterior = StrParaDbl(MaskCartaoDebito.Text)
    
'    End If

    If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
        'Fecha a tela e abre a certa
        If StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO Then
            gobjVenda.iTipoForcado = OPTION_CF
            Call BotaoFechar_Click
        End If
    End If
    
    Exit Sub
    
Erro_MaskCartaoDebito_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99724
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160199)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskCartaoCredito_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_MaskCartaoCredito_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskCartaoCredito.Text)) > 0 Then
        'verifica se é um valor válido
        lErro = Valor_NaoNegativo_Critica(MaskCartaoCredito.Text)
        If lErro <> SUCESSO Then gError 99725
    
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
    
    If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
        'Fecha a tela e abre a certa
        If StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO Then
            gobjVenda.iTipoForcado = OPTION_CF
            Call BotaoFechar_Click
        End If
    End If
    
    Exit Sub
    
Erro_MaskCartaoCredito_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99725
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160200)

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
        If lErro <> SUCESSO Then gError 99752
        
    End If
    
    
    'Se o valor não foi alterado ==> Sai
'    If gdSaldoOutrosAnterior <> StrParaDbl(MaskOutros.Text) Then
    
        'Exibe o valor formatado
        Outros.Caption = Format(StrParaDbl(BotaoOutros.Caption) + StrParaDbl(MaskOutros.Text), "Standard")
            
        'Recalcula os totais
        Call Recalcula_Valores2
            
        'Atualiza o Movimento
        Call Inclui_Movimento(StrParaDbl(MaskOutros.Text), MOVIMENTOCAIXA_RECEB_OUTROS)
    
        If Len(Trim(MaskOutros.Text)) > 0 Then MaskOutros.Text = Round(StrParaDbl(MaskOutros.Text), 2)
    
        'Guarda o valor que está em outros
        gdSaldoOutrosAnterior = StrParaDbl(MaskOutros.Text)
    
'    End If
    
    Exit Sub
    
Erro_MaskOutros_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99752
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160201)

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
        If lErro <> SUCESSO Then gError 99753
        
    End If
        
    'Se o valor não foi alterado ==> Sai
'    If gdSaldoTicketAnterior <> StrParaDbl(MaskTicket.Text) Then
    
        'Exibe o valor formatado
        Ticket.Caption = Format(StrParaDbl(BotaoTicket.Caption) + StrParaDbl(MaskTicket.Text), "Standard")
    
        'Recalcula os totais
        Call Recalcula_Valores2
            
        'ATualiza o movimento
        Call Inclui_Movimento(StrParaDbl(MaskTicket.Text), MOVIMENTOCAIXA_RECEB_VALETICKET)
    
        If Len(Trim(MaskTicket.Text)) > 0 Then MaskTicket.Text = Round(StrParaDbl(MaskTicket.Text), 2)
    
        'Guarda o valor atual
        gdSaldoTicketAnterior = StrParaDbl(MaskTicket.Text)
    
'    End If
    
    Exit Sub
    
Erro_MaskTicket_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99753
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160202)

    End Select

    Exit Sub
    
End Sub

Private Sub Datafinal_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Datafinal_Validate
    
    If Len(Trim(DataFinal.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 99726
        
        If StrParaDate(LabelEmissao.Caption) > StrParaDate(DataFinal.Text) Then gError 99894
        
        Duracao.Text = StrParaDate(DataFinal.Text) - StrParaDate(LabelEmissao.Caption)
    Else
        Duracao.Text = ""
    End If
        
    Exit Sub
    
Erro_Datafinal_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99726
        
        Case 99894
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAEMISSAO_MAIOR, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160203)

    End Select

    Exit Sub
    
End Sub

Private Sub Timer2_Timer()

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
    
    Exit Sub

End Sub

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 99727
    
    Call Datafinal_Validate(False)
    
    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 99727

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160204)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 99728
    
    Call Datafinal_Validate(False)
    
    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 99728

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160205)

    End Select

    Exit Sub

End Sub

Private Sub Duracao_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Duracao_Validate
    
    If Len(Trim(Duracao.Text)) > 0 Then
        lErro = Valor_NaoNegativo_Critica(Duracao.Text)
        If lErro <> SUCESSO Then gError 99729
        DataFinal.PromptInclude = False
        DataFinal.Text = Format(CDate(LabelEmissao.Caption) + CInt(Duracao.Text), "dd/mm/yy")
        DataFinal.PromptInclude = True
    Else
        DataFinal.PromptInclude = False
        DataFinal.Text = ""
        DataFinal.PromptInclude = True
    End If
        
    Exit Sub
    
Erro_Duracao_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99729
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160206)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCheques_Click()
    
Dim lErro As Long
Dim objCheques As New ClassChequePre
Dim dTotal As Double
Dim dTotal1 As Double

On Error GoTo Erro_BotaoCheques_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoCheque", gobjVenda)
        
    'Faz o somatório dos cheques
    For Each objCheques In gobjVenda.colCheques
        If objCheques.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
            dTotal = dTotal + objCheques.dValor
        Else
            dTotal1 = dTotal1 + objCheques.dValor
        End If
    Next
    
    'Joga o valor do somatório no botão
    BotaoCheques.Caption = Format(dTotal, "standard")
    If dTotal1 <> 0 Then MaskCheques.Text = Format(dTotal1, "standard")
    
    'Atualiza o total
    ChequeVista.Caption = Format(StrParaDbl(MaskCheques.Text) + StrParaDbl(BotaoCheques.Caption), "standard")
     
    Call Recalcula_Valores2
                    
    Exit Sub
    
Erro_BotaoCheques_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160207)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCartaoDebito_Click()
    
Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim dTotal As Double

On Error GoTo Erro_BotaoCartaoDebito_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoCartao", gobjVenda, MOVIMENTOCAIXA_RECEB_CARTAODEBITO)
        
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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160208)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCartaoCredito_Click()
    
Dim lErro As Long
Dim objMovCaixa As New ClassMovimentoCaixa
Dim dTotal As Double

On Error GoTo Erro_BotaoCartaoCredito_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoCartao", gobjVenda, MOVIMENTOCAIXA_RECEB_CARTAOCREDITO)
        
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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160209)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCarne_Click()
    
Dim lErro As Long
Dim objCarneParc As New ClassCarneParcelas
Dim dTotal As Double

On Error GoTo Erro_BotaoCarne_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_TelaECF_Modal("PagamentoPrazo", gobjVenda)
        
    'Faz o somatório dos Carne
    For Each objCarneParc In gobjVenda.objCarne.colParcelas
        dTotal = dTotal + objCarneParc.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoCarne.Caption = Format(dTotal, "standard")
        
    'Atualiza o total
    Carne.Caption = Format(StrParaDbl(BotaoCarne.Caption), "standard")
    
    Call Recalcula_Valores2
            
    Exit Sub
    
Erro_BotaoCarne_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160210)

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
    BotaoTroca.Caption = Format(dTotal, "standard")
        
    'Atualiza o total
    Troca.Caption = Format(StrParaDbl(BotaoTroca.Caption), "standard")
    
    If StrParaDbl(DescontoPerc.Text) > 0 And (StrParaDbl(Total.Caption) - StrParaDbl(BotaoTroca.Caption)) > 0 Then
        DescontoValor.Text = Round((StrParaDbl(DescontoPerc.Text) / 100) * (StrParaDbl(Total.Caption) - StrParaDbl(BotaoTroca.Caption)), 2)
        gobjVenda.objCupomFiscal.dValorDesconto = StrParaDbl(DescontoValor.Text)
    End If
    
    If StrParaDbl(DescontoPerc1.Text) > 0 And (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption) + StrParaDbl(DescontoValor.Text))) > 0 Then
        DescontoValor1.Text = Round((StrParaDbl(DescontoPerc1.Text) / 100) * (StrParaDbl(Total.Caption) - (StrParaDbl(BotaoTroca.Caption) + StrParaDbl(DescontoValor.Text))), 2)
        gobjVenda.objCupomFiscal.dValorDesconto1 = StrParaDbl(DescontoValor1.Text)
    End If
    
    Call Recalcula_Valores2
            
    Exit Sub
    
Erro_BotaoTroca_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160211)

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
    BotaoOutros.Caption = Format(dTotal, "standard")
        
    'Atualiza o total
    Outros.Caption = Format(StrParaDbl(BotaoOutros.Caption), "standard")
    
    Call Recalcula_Valores2
    
    Exit Sub
            
Erro_BotaoOutros_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160212)

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
    BotaoTicket.Caption = Format(dTotal, "standard")
        
    'Atualiza o total
    Ticket.Caption = Format(StrParaDbl(MaskTicket.Text) + StrParaDbl(BotaoTicket.Caption), "standard")
    
    Call Recalcula_Valores2
    
    Exit Sub
            
Erro_BotaoTicket_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160213)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    
    gobjGenerico.vVariavel = vbAbort
    Unload Me
    giSaida = 1

End Sub

Private Sub BotaoImprimir_Click()
    
Dim lErro As Long
Dim dtDataFinal As Date
Dim objProdutoNomeRed As Object
Dim objTela As Object

On Error GoTo Erro_BotaoImprimir_Click
        
'    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105896
        
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207969
        
    If gobjVenda.iTipo <> OPTION_DAV Then gError 204741
        
    'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
    If gobjVenda.iTipo = OPTION_DAV And gobjVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210506
        
        
    gobjVenda.objCupomFiscal.dtDataEmissao = Date
    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
'    gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
        
        
    dtDataFinal = StrParaDate(DataFinal.Text)
    

    Set objTela = Me
    
    lErro = CF_ECF("Gravar_Orcamento", objTela, gobjVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'le os registros do orcamento e loca o arquivo
    lErro = CF_ECF("Imprime_OrcamentoECF", dtDataFinal, gobjVenda.objCupomFiscal.lNumOrcamento, objTela, gobjVenda)
    If lErro <> SUCESSO Then gError 105886
    
    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
    
    gobjGenerico.vVariavel = vbOK
    
    Unload Me
    
    giSaida = 1
        
    Exit Sub
        
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 105886, 207969, ERRO_SEM_MENSAGEM
    
        Case 105896
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
    
        Case 204741
            Call Rotina_ErroECF(vbOKOnly, ERRO_IMPRESSAO_NAO_PERMITIDA, gErr)
    
        Case 210506
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_PODE_SER_REIMPRESSO, gErr)
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160214)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objVenda As New ClassVenda

On Error GoTo Erro_BotaoExcluir_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207970
    
    objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento

    'le os registros do orcamento e loca o arquivo
    lErro = CF_ECF("OrcamentoECF_Le", objVenda)
    If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 204742
    
    If lErro = 210447 Then gError 210451
    
    'se o orcamento ja esta cadastrado ==> nao pode excluir PAFECF
    If lErro = SUCESSO Then

        'impede a exclusao do orcamento caso ja esteja cadastrado
        lErro = CF_ECF("Caixa_Exclui_Orcamento1", gobjVenda)
        
        gError 204743
    
    Else

        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_ORCAMENTO)
        
        If vbMsgRes = vbNo Then gError 204744
        
    End If

    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF


    gobjGenerico.vVariavel = vbOK

    Unload Me

    giSaida = 1
    
    Exit Sub
        
Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 204742, 204743, 204744, 207970
                
        Case 210451
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda.objCupomFiscal.lNumOrcamento)
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160215)

    End Select
    
    Exit Sub

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

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objTela As Object


On Error GoTo Erro_BotaoGravar_Click

    Set objTela = Me

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207968
    
    'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
    If gobjVenda.iTipo = OPTION_DAV And gobjVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210505

    gobjVenda.objCupomFiscal.dtDataEmissao = Date
    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
'    gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior

    lErro = CF_ECF("Gravar_Orcamento", objTela, gobjVenda)
    If lErro <> SUCESSO Then gError 204304
        
    Set gobjVenda = New ClassVenda
    gobjVenda.iCodModeloECF = giCodModeloECF
    
    gobjGenerico.vVariavel = vbOK
    
    Unload Me
    
    giSaida = 1
        
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 204304, 207968
        
        Case 210505
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204306)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoAbrirGaveta_Click()

Dim lErro As Long
Dim lSequencial As Long
Dim objCheque As New ClassChequePre
Dim objMovCaixa As New ClassMovimentoCaixa
Dim vbMsgRes As VbMsgBoxResult
Dim objTela As Object
Dim dtDataFinal As Date

On Error GoTo Erro_BotaoAbrirGaveta_Click

    If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
        'Fecha a tela e abre a certa
        If StrParaDbl(MaskCartaoCredito.Text) + StrParaDbl(MaskCartaoDebito.Text) > DELTA_VALORMONETARIO Then
            gobjVenda.iTipoForcado = OPTION_CF
            Call BotaoFechar_Click
            Exit Sub
        End If
    End If
        
    lErro = CF_ECF("Testa_Limite_Desconto", gobjVenda)
    If lErro <> SUCESSO Then gError 126780
        
    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105897
        
    'Se o valor é insuficiente para pagar
    If StrParaDbl(Falta.Caption) > 0 Then gError 99749
    
    'gera o CodBarrasCarne do Carnê
    If gobjVenda.objCarne.colParcelas.Count > 0 Then gobjVenda.objCarne.sCodBarrasCarne = FormataCpoNum(giFilialEmpresa, 5) & FormataCpoNum(giCodCaixa, 5) & FormataCpoNum(gobjVenda.objCupomFiscal.lNumero, 10)
    
    'Calcula o troca da tela
    Call Calcula_Troco
    
    gobjVenda.iTipo = OPTION_ORCAMENTO
       
    gobjVenda.objCupomFiscal.dtDataEmissao = Date
    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(Troco.Caption)
    gobjVenda.objCupomFiscal.lDuracao = StrParaLong(Duracao.Text)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iECF = giCodECF
    gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
'    gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
    gobjVenda.objCupomFiscal.iStatus = STATUS_BAIXADO
    
    If giRemoveOrc = REMOVER_ORC Then
    
        'exclui o orcamento que está sendo transformado em cupom
        lErro = CF_ECF("Caixa_Exclui_Orcamento", gobjVenda)
        If lErro <> SUCESSO And lErro <> 105761 Then gError 105766
    
    End If
    
    'grava o orcamento.
    lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
    If lErro <> SUCESSO Then gError 105863
    
    'Atualiza os Movimentos nas coleções globais
    Call CF_ECF("Atualiza_Movimentos_Memoria1", gobjVenda)
    
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

    'Atribui para a coleção global o objvenda
    gcolVendas.Add gobjVenda
    
    'Envia aviso perguntando se deseja imprimir o orçamento
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ORCAMENTO_IMPRESSAO)
    If vbMsgRes = vbYes Then
    
        Set objTela = Me
        dtDataFinal = Date
        
        Call CF_ECF("Imprime_OrcamentoECF", dtDataFinal, gobjVenda.objCupomFiscal.lNumOrcamento, objTela, gobjVenda)
    
    End If
    
    'Abrir a Gaveta
    Call AFRAC_AbrirGaveta
    
    Set gobjVenda = New ClassVenda
    
    gobjGenerico.vVariavel = vbOK
    
    Unload Me
    
    giSaida = 1
    
    Exit Sub
        
Erro_BotaoAbrirGaveta_Click:

    Select Case gErr
    
        Case 99749
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_INSUFICIENTE, gErr)
        
        Case 105863, 126780, ERRO_SEM_MENSAGEM
        
        Case 105897
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160217)

    End Select
    
    Exit Sub
     
End Sub

'Private Function Grava_Venda_Arquivo_Orc() As Long
'
'Dim objMovCaixa As ClassMovimentoCaixa
'Dim objCheque As ClassChequePre
'Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
'Dim colRegistro As New Collection
'Dim iIndice As Integer
'Dim lNum As Long
'Dim lErro As Long
'Dim objVenda As ClassVenda
'Dim lSequencial As Long
'Dim sLog As String
'Dim sDescForma As String
'Dim objTiposMeiosPagtos As ClassTMPLoja
'Dim iTipo As Integer
'Dim dPagTEF As Double
'Dim colMeiosPag As New Collection
'Dim dValorCC As Double
'Dim dValorCD As Double
'Dim dValorCarne As Double
'Dim dValorCheque As Double
'Dim dValorDin As Double
'Dim dValorOutros As Double
'Dim dValorTroca As Double
'Dim dValorVT As Double
'Dim dValor As String
'Dim objMovCaixa1 As New ClassMovimentoCaixa
'Dim objMovCaixa2 As New ClassMovimentoCaixa
'Dim iNovoIndice As Integer
'Dim iIndice2 As Integer
'Dim objAux As New ClassMovimentoCaixa
'Dim dMaior As Double
'Dim dTotal As Double
'Dim lNumero As Long
'Dim sDescricao As String
'Dim lSequencialCaixa As Long
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Grava_Venda_Arquivo_Orc
'
'    'iNFORMAR PARA IMPRESSORA AS FORMAS PAGTO
'    For Each objMovCaixa In gobjVenda.colMovimentosCaixa
'
'        'Seleciona o tipo
'        Select Case objMovCaixa.iTipo
'
'            Case MOVIMENTOCAIXA_RECEB_CARNE:
'                If dValorCarne = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_CARNE Then dValorCarne = dValorCarne + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_CARNE
'                    objMovCaixa2.dValor = dValorCarne
'                    colMeiosPag.Add objMovCaixa2
'                End If
'            Case MOVIMENTOCAIXA_RECEB_CARTAOCREDITO:
'                If dValorCC = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Then dValorCC = dValorCC + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
'                    objMovCaixa2.dValor = dValorCC
'                    colMeiosPag.Add objMovCaixa2
'                End If
'            Case MOVIMENTOCAIXA_RECEB_CARTAODEBITO:
'                If dValorCD = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO Then dValorCD = dValorCD + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
'                    objMovCaixa2.dValor = dValorCD
'                    colMeiosPag.Add objMovCaixa2
'                End If
'            Case MOVIMENTOCAIXA_RECEB_CHEQUE:
'                If dValorCheque = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_CHEQUE Then dValorCheque = dValorCheque + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_CHEQUE
'                    objMovCaixa2.dValor = dValorCheque
'                    colMeiosPag.Add objMovCaixa2
'                End If
'            Case MOVIMENTOCAIXA_RECEB_DINHEIRO:
'                If dValorDin = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then dValorDin = dValorDin + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_DINHEIRO
'                    objMovCaixa2.dValor = dValorDin
'                    colMeiosPag.Add objMovCaixa2
'                End If
'            Case MOVIMENTOCAIXA_RECEB_OUTROS:
'                If dValorOutros = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS Then dValorOutros = dValorOutros + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_OUTROS
'                    objMovCaixa2.dValor = dValorOutros
'                    colMeiosPag.Add objMovCaixa2
'                End If
'            Case MOVIMENTOCAIXA_RECEB_TROCA:
'                If dValorTroca = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_TROCA Then dValorTroca = dValorTroca + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_TROCA
'                    objMovCaixa2.dValor = dValorTroca
'                    colMeiosPag.Add objMovCaixa2
'                End If
'            Case MOVIMENTOCAIXA_RECEB_VALETICKET:
'                If dValorVT = 0 Then
'                    For Each objMovCaixa1 In gobjVenda.colMovimentosCaixa
'                        If objMovCaixa1.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then dValorVT = dValorVT + objMovCaixa1.dValor
'                    Next
'                    Set objMovCaixa2 = New ClassMovimentoCaixa
'                    objMovCaixa2.iTipo = TIPOMEIOPAGTOLOJA_VALE_TICKET
'                    objMovCaixa2.dValor = dValorVT
'                    colMeiosPag.Add objMovCaixa2
'                End If
'        End Select
'    Next
'
'    'ordenar por valores...
'    For iIndice = 1 To colMeiosPag.Count - 1
'        Set objMovCaixa = colMeiosPag.Item(iIndice)
'        dMaior = objMovCaixa.dValor
'        iNovoIndice = iIndice
'        For iIndice2 = iIndice To colMeiosPag.Count
'            Set objMovCaixa1 = colMeiosPag.Item(iIndice2)
'            If objMovCaixa1.dValor > dMaior Then
'                dMaior = objMovCaixa1.dValor
'                iNovoIndice = iIndice2
'            End If
'        Next
'        Set objMovCaixa1 = colMeiosPag.Item(iNovoIndice)
'        Call Inverte_Col(objAux, objMovCaixa)
'        Call Inverte_Col(objMovCaixa, objMovCaixa1)
'        Call Inverte_Col(objMovCaixa1, objAux)
'    Next
'
'    'grava o orcamento.
'    lErro = Grava_Orcamento_ECF()
'    If lErro <> SUCESSO Then gError 105834
'
'    Grava_Venda_Arquivo_Orc = SUCESSO
'
'    Exit Function
'
'Erro_Grava_Venda_Arquivo_Orc:
'
'    Grava_Venda_Arquivo_Orc = gErr
'
'    Select Case gErr
'
'        Case 99939, 99902, 99943, 99952, 99953, 99901, 105834
'
'        Case 112394
'            Call Rotina_ErroECF(vbOKOnly, ERRO_MEIOSPAG_ULTRAPASSAM, gErr, Error)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160218)
'
'    End Select
'
'    lSequencial = glSeqTransacaoAberta
'
'    Call CF_ECF("Caixa_Transacao_Rollback", lSequencial)
'
'    Exit Function
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
    
End Sub


'Private Sub Desmembrar_Click()
'
'Dim iIndice As Integer
'Dim objVenda As New ClassVenda
'Dim lErro As Long
'
'On Error GoTo Erro_Desmembrar_Click
'
'    For iIndice = 1 To gcolOrcamentos.Count
'        Set objVenda = gcolOrcamentos.Item(iIndice)
'        'Se achou o Orçamento com o Número
'        If objVenda.iTipo = OPTION_ORCAMENTO And objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento Then
'            'Apaga o orçamento
'            gcolOrcamentos.Remove (iIndice)
'            Exit For
'        End If
'    Next
'
'    lErro = CF_ECF("Desmembrar_ECF", gobjVenda)
'    If lErro <> SUCESSO Then gError 109812
'
'    gcolOrcamentos.Add gobjVenda
'
'    Call Traz_Dados_Tela
'
'    Exit Sub
'
'Erro_Desmembrar_Click:
'
'    Select Case gErr
'
'        Case 109812
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160219)
'
'    End Select
'
'    Exit Sub
'
'End Sub

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
    
    Total.Caption = Format(gobjVenda.objCupomFiscal.dValorProdutos, "standard")
    
    AcrescimoValor.Text = IIf(gobjVenda.objCupomFiscal.dValorAcrescimo <> 0, Format(gobjVenda.objCupomFiscal.dValorAcrescimo, "standard"), "")
    DescontoValor.Text = IIf(gobjVenda.objCupomFiscal.dValorDesconto <> 0, Format(gobjVenda.objCupomFiscal.dValorDesconto, "Standard"), "")
    DescontoValor1.Text = IIf(gobjVenda.objCupomFiscal.dValorDesconto1 <> 0, Format(gobjVenda.objCupomFiscal.dValorDesconto1, "Standard"), "")
    
    Call AcrescimoValor_Validate(False)
    Call DescontoValor_Validate(False)
    Call DescontoValor1_Validate(False)
    
    If gobjVenda.objCupomFiscal.dtDataEmissao <> 0 Then LabelEmissao.Caption = Format(gobjVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yyyy")
    
    Duracao.Text = gobjVenda.objCupomFiscal.lDuracao
    Call Duracao_Validate(False)
    
    Call Recalcula_Valores1
    Call Recalcula_Valores2
    
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
    gobjVenda.objCupomFiscal.dHoraEmissao = 0
        
    Call Limpa_Tela_Pagamento
    
    Call Recalcula_Valores1
    
    Exit Sub
        
Erro_Botaolimpar_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160220)

    End Select
    
    Exit Sub
        
End Sub

Private Sub Limpa_Tela_Pagamento()

    Call Inicializa_Valores
    
    MaskDinheiro.Text = ""
    MaskCheques.Text = ""
    MaskCartaoCredito.Text = ""
    MaskCartaoDebito.Text = ""
    MaskOutros.Text = ""
    MaskTicket.Text = ""
    DataFinal.PromptInclude = False
    DataFinal.Text = ""
    DataFinal.PromptInclude = True
    Duracao.Text = ""
End Sub

Sub Recalcula_Valores1()

    APagar.Caption = Format(StrParaDbl(Total.Caption) - StrParaDbl(DescontoValor.Text) - StrParaDbl(DescontoValor1.Text) + StrParaDbl(AcrescimoValor.Text), "standard")
        
    Call Calcula_Faltatroco
    
End Sub

Sub Recalcula_Valores2()

    Pago.Caption = Format(StrParaDbl(Dinheiro.Caption) + StrParaDbl(ChequeVista.Caption) + StrParaDbl(CartaoDebito.Caption) + StrParaDbl(CartaoCredito.Caption) + StrParaDbl(Carne.Caption) + StrParaDbl(Troca.Caption) + StrParaDbl(Outros.Caption) + StrParaDbl(Ticket.Caption), "standard")
        
    Call Calcula_Faltatroco
    
End Sub

Sub Calcula_Faltatroco()
    
    If StrParaDbl(APagar.Caption) > StrParaDbl(Pago.Caption) Then
        Falta.Caption = Format(StrParaDbl(APagar.Caption) - StrParaDbl(Pago.Caption), "standard")
        Troco.Caption = Format(0, "standard")
    Else
        Troco.Caption = Format(StrParaDbl(Pago.Caption) - StrParaDbl(APagar.Caption), "standard")
        Falta.Caption = Format(0, "standard")
    End If
    
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(APagar.Caption)
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
Dim ShiftKey As Integer
    
    ShiftKey = Shift And 7
    
    If giSaida = 1 Then Exit Sub
    
    If ShiftKey = 0 Then
    
        Select Case KeyCode
        
            Case vbKeyReturn
                KeyCode = vbKeyTab
        
            Case vbKeyEscape
                If Not TrocaFoco(Me, BotaoExcluir) Then Exit Sub
                Call BotaoExcluir_Click
        
            Case vbKeyF2
                If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
                Call BotaoFechar_Click
    
            Case vbKeyF3
                If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
                Call BotaoGravar_Click
            
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
        
            Case vbKeyF12
                If Not TrocaFoco(Me, BotaoImprimir) Then Exit Sub
                Call BotaoImprimir_Click
                KeyCode = vbKeyF2
            
        End Select
        
    ElseIf ShiftKey = vbShiftMask Then
    
        Select Case KeyCode
            
            Case vbKeyF3
                If Not TrocaFoco(Me, Nothing) Then Exit Sub
                Call BotaoAbrirGaveta_Click
    
            Case vbKeyF5
                If Not TrocaFoco(Me, Nothing) Then Exit Sub
                If gobjNFeInfo.iFocaTipoVenda = MARCADO Then
                    'Fecha a tela e abre a certa
                    gobjVenda.iTipoForcado = OPTION_CF
                    gobjVenda.iForcadoF5 = MARCADO
                    Call BotaoFechar_Click
                    Exit Sub
                End If
    
        End Select
    
    End If
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
        
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
    
    Caption = Formata_Campo(ALINHAMENTO_DIREITA, 20, " ", "Orçamento") & "Filial : " & giFilialEmpresa & "    Caixa : " & giCodCaixa & "    Operador : " & sOper
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FechaOrcamento"
    
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'***** fim do trecho a ser copiado ******
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

