VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TRPVoucher 
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   5205
      Index           =   4
      Left            =   225
      TabIndex        =   98
      Top             =   870
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame FrameC 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4710
         Index           =   2
         Left            =   30
         TabIndex        =   129
         Top             =   360
         Visible         =   0   'False
         Width           =   9030
         Begin VB.Frame Frame15 
            Caption         =   "Comissão Interna"
            Height          =   2325
            Left            =   30
            TabIndex        =   139
            Top             =   2370
            Width           =   8985
            Begin MSMask.MaskEdBox PercPromoComiss 
               Height          =   225
               Left            =   6675
               TabIndex        =   268
               Top             =   840
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox VendedorPromo 
               Height          =   225
               Left            =   1995
               TabIndex        =   157
               Top             =   1170
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataPromo 
               Height          =   225
               Left            =   225
               TabIndex        =   140
               Top             =   1380
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorPromoComiss 
               Height          =   225
               Left            =   6300
               TabIndex        =   141
               Top             =   1200
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox ValorPromoBase 
               Height          =   225
               Left            =   4950
               TabIndex        =   142
               Top             =   1185
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
            Begin MSFlexGridLib.MSFlexGrid GridPromotor 
               Height          =   1695
               Left            =   45
               TabIndex        =   143
               Top             =   210
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   2990
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
            Begin VB.Label TotalComiPro 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   7350
               TabIndex        =   145
               Top             =   1935
               Width           =   1560
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   56
               Left            =   6645
               TabIndex        =   144
               Top             =   1995
               Width           =   630
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Comissão do Emissor (Over)"
            Height          =   2325
            Left            =   30
            TabIndex        =   130
            Top             =   0
            Width           =   9000
            Begin MSMask.MaskEdBox Emi 
               Height          =   225
               Left            =   810
               TabIndex        =   164
               Top             =   1020
               Width           =   2800
               _ExtentX        =   4948
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox StatusEmissor 
               Height          =   225
               Left            =   4245
               TabIndex        =   151
               Top             =   1140
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin VB.CommandButton BotaoAbrirEmi 
               Caption         =   "Abrir"
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
               Left            =   5385
               TabIndex        =   148
               Top             =   210
               Width           =   660
            End
            Begin VB.CommandButton BotaoAbrirFatEmi 
               Caption         =   "Abrir Fatura"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   60
               TabIndex        =   132
               Top             =   1935
               Width           =   2080
            End
            Begin VB.TextBox HistoricoEmissor 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   5295
               MaxLength       =   250
               TabIndex        =   131
               Top             =   1440
               Width           =   6000
            End
            Begin MSMask.MaskEdBox NumTitEmissor 
               Height          =   225
               Left            =   3165
               TabIndex        =   133
               Top             =   1320
               Width           =   700
               _ExtentX        =   1244
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataEmissor 
               Height          =   225
               Left            =   390
               TabIndex        =   134
               Top             =   1350
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorEmissor 
               Height          =   225
               Left            =   1710
               TabIndex        =   135
               Top             =   1365
               Width           =   800
               _ExtentX        =   1402
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
            Begin MSFlexGridLib.MSFlexGrid GridEmissor 
               Height          =   1215
               Left            =   60
               TabIndex        =   136
               Top             =   585
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   2143
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
            Begin VB.Label Label1 
               Caption         =   "Forn. Emissor:"
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
               Index           =   50
               Left            =   480
               TabIndex        =   150
               Top             =   255
               Width           =   1275
            End
            Begin VB.Label Emissor 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   1785
               TabIndex        =   149
               Top             =   210
               Width           =   3600
            End
            Begin VB.Label PercComiEmi 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   7830
               TabIndex        =   147
               Top             =   210
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "% Comissão:"
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
               Index           =   61
               Left            =   6750
               TabIndex        =   146
               Top             =   255
               Width           =   1140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   54
               Left            =   6630
               TabIndex        =   138
               Top             =   1995
               Width           =   630
            End
            Begin VB.Label TotalComiEmi 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   7350
               TabIndex        =   137
               Top             =   1935
               Width           =   1560
            End
         End
      End
      Begin VB.Frame FrameC 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4710
         Index           =   1
         Left            =   30
         TabIndex        =   100
         Top             =   360
         Width           =   9030
         Begin VB.Frame Frame13 
            Caption         =   "Comissão do Correntista"
            Height          =   2325
            Left            =   30
            TabIndex        =   115
            Top             =   2370
            Width           =   9000
            Begin MSMask.MaskEdBox Corr 
               Height          =   225
               Left            =   390
               TabIndex        =   162
               Top             =   1380
               Width           =   2800
               _ExtentX        =   4948
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin VB.CommandButton BotaoAbrirCor 
               Caption         =   "Abrir"
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
               Left            =   5385
               TabIndex        =   153
               Top             =   210
               Width           =   660
            End
            Begin VB.CommandButton BotaoAbrirFatCor 
               Caption         =   "Abrir Fatura"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   60
               TabIndex        =   120
               Top             =   1950
               Width           =   2080
            End
            Begin VB.TextBox HistoricoCorr 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   5055
               MaxLength       =   250
               TabIndex        =   117
               Top             =   1350
               Width           =   6000
            End
            Begin MSMask.MaskEdBox NumTitCorr 
               Height          =   225
               Left            =   3105
               TabIndex        =   116
               Top             =   735
               Width           =   700
               _ExtentX        =   1244
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataCorr 
               Height          =   225
               Left            =   585
               TabIndex        =   118
               Top             =   750
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorCorr 
               Height          =   225
               Left            =   1725
               TabIndex        =   119
               Top             =   750
               Width           =   800
               _ExtentX        =   1402
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox StatusCorr 
               Height          =   225
               Left            =   4260
               TabIndex        =   128
               Top             =   735
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridCorr 
               Height          =   1215
               Left            =   60
               TabIndex        =   121
               Top             =   585
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   2143
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
            Begin VB.Label Correntista 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   1785
               TabIndex        =   127
               Top             =   210
               Width           =   3600
            End
            Begin VB.Label Label1 
               Caption         =   "Correntista:"
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
               Index           =   60
               Left            =   690
               TabIndex        =   126
               Top             =   270
               Width           =   1080
            End
            Begin VB.Label PercComiCor 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   7860
               TabIndex        =   125
               Top             =   210
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "% Comissão:"
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
               Index           =   59
               Left            =   6810
               TabIndex        =   124
               Top             =   240
               Width           =   1140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   55
               Left            =   6690
               TabIndex        =   123
               Top             =   1995
               Width           =   630
            End
            Begin VB.Label TotalComiCor 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   7365
               TabIndex        =   122
               Top             =   1935
               Width           =   1560
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Comissão do Representante"
            Height          =   2325
            Left            =   30
            TabIndex        =   101
            Top             =   0
            Width           =   9000
            Begin MSMask.MaskEdBox Rep 
               Height          =   225
               Left            =   435
               TabIndex        =   163
               Top             =   1230
               Width           =   2800
               _ExtentX        =   4948
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin VB.CommandButton BotaoAbrirRep 
               Caption         =   "Abrir"
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
               Left            =   5385
               TabIndex        =   152
               Top             =   210
               Width           =   660
            End
            Begin VB.CommandButton BotaoAbrirFatRep 
               Caption         =   "Abrir Fatura"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   60
               TabIndex        =   107
               Top             =   1950
               Width           =   2080
            End
            Begin VB.TextBox HistoricoRepr 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   5535
               MaxLength       =   250
               TabIndex        =   104
               Top             =   1455
               Width           =   6000
            End
            Begin MSMask.MaskEdBox StatusRepr 
               Height          =   225
               Left            =   4650
               TabIndex        =   102
               Top             =   1470
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumTitRepr 
               Height          =   225
               Left            =   3375
               TabIndex        =   103
               Top             =   1455
               Width           =   700
               _ExtentX        =   1244
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataRepr 
               Height          =   225
               Left            =   810
               TabIndex        =   105
               Top             =   1455
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorRepr 
               Height          =   225
               Left            =   2025
               TabIndex        =   106
               Top             =   1455
               Width           =   800
               _ExtentX        =   1402
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
            Begin MSFlexGridLib.MSFlexGrid GridRepr 
               Height          =   1215
               Left            =   60
               TabIndex        =   108
               Top             =   585
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   2143
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   53
               Left            =   6705
               TabIndex        =   114
               Top             =   1995
               Width           =   630
            End
            Begin VB.Label TotalComiRep 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   7380
               TabIndex        =   113
               Top             =   1935
               Width           =   1560
            End
            Begin VB.Label Representante 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   1785
               TabIndex        =   112
               Top             =   210
               Width           =   3600
            End
            Begin VB.Label Label1 
               Caption         =   "Representante:"
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
               Index           =   57
               Left            =   360
               TabIndex        =   111
               Top             =   255
               Width           =   1440
            End
            Begin VB.Label PercComiRep 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   7860
               TabIndex        =   110
               Top             =   210
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "% Comissão:"
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
               Index           =   58
               Left            =   6795
               TabIndex        =   109
               Top             =   240
               Width           =   1140
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   5160
         Left            =   0
         TabIndex        =   99
         Top             =   0
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   9102
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Representante e Correntista"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Emissor e Interna"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5310
      Index           =   1
      Left            =   165
      TabIndex        =   19
      Top             =   825
      Width           =   9195
      Begin VB.CommandButton BotaoHistAlt 
         Caption         =   "Histórico de Alterações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5250
         TabIndex        =   12
         ToolTipText     =   "Mostra o histórico das alterações feitas no voucher"
         Top             =   4740
         Width           =   1260
      End
      Begin VB.CommandButton BotaoVendedores 
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
         Height          =   525
         Left            =   7845
         TabIndex        =   14
         ToolTipText     =   "Mostra os vendedores que receberão a comissão interna"
         Top             =   4740
         Width           =   1245
      End
      Begin VB.CommandButton BotaoManutencao 
         Caption         =   "Manutenção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6555
         TabIndex        =   13
         ToolTipText     =   "Abre a tela de manutenção do voucher"
         Top             =   4740
         Width           =   1260
      End
      Begin VB.CommandButton BotaoComissao 
         Caption         =   "Comissão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1290
         TabIndex        =   9
         ToolTipText     =   "Permite alterações no comissionamento"
         Top             =   4740
         Width           =   1260
      End
      Begin VB.CommandButton BotaoHist 
         Caption         =   "Detalhamento dos Valores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2580
         TabIndex        =   10
         ToolTipText     =   "Mostra o detalhamento dos valores do voucher"
         Top             =   4740
         Width           =   1365
      End
      Begin VB.CommandButton BotaoHistOcor 
         Caption         =   "Histórico de Ocorrências"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3960
         TabIndex        =   11
         ToolTipText     =   "Mostra as ocorrências do voucher"
         Top             =   4740
         Width           =   1260
      End
      Begin VB.Frame Frame7 
         Caption         =   "Complemento"
         Height          =   1500
         Left            =   30
         TabIndex        =   77
         Top             =   3210
         Width           =   9150
         Begin VB.Frame Frame19 
            Caption         =   "Em caso de emergência ligar para"
            Height          =   1335
            Left            =   5250
            TabIndex        =   255
            Top             =   105
            Width           =   3825
            Begin MSMask.MaskEdBox ContatoTelefone 
               Height          =   225
               Left            =   2100
               TabIndex        =   256
               Top             =   795
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ContatoNome 
               Height          =   225
               Left            =   300
               TabIndex        =   257
               Top             =   795
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
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
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridContatos 
               Height          =   480
               Left            =   45
               TabIndex        =   258
               Top             =   180
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   847
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Pago com Cartão"
            Height          =   615
            Left            =   75
            TabIndex        =   166
            Top             =   195
            Width           =   1515
            Begin VB.OptionButton OptSim 
               Caption         =   "Sim"
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
               Left            =   105
               TabIndex        =   168
               Top             =   225
               Width           =   645
            End
            Begin VB.OptionButton OptNao 
               Caption         =   "Não"
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
               Left            =   765
               TabIndex        =   167
               Top             =   240
               Width           =   705
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Dias Antecipados"
            Height          =   630
            Left            =   90
            TabIndex        =   91
            Top             =   825
            Width           =   1500
            Begin VB.OptionButton OptAntcNao 
               Caption         =   "Não"
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
               Left            =   750
               TabIndex        =   93
               Top             =   300
               Width           =   690
            End
            Begin VB.OptionButton OptAntcSim 
               Caption         =   "Sim"
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
               Left            =   90
               TabIndex        =   92
               Top             =   300
               Width           =   675
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Destino"
            Height          =   270
            Left            =   5070
            TabIndex        =   86
            Top             =   -75
            Visible         =   0   'False
            Width           =   3825
            Begin VB.Label DestinoVou 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   945
               TabIndex        =   88
               Top             =   675
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.Label Label1 
               Caption         =   "Dest. Vou:"
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
               Index           =   47
               Left            =   45
               TabIndex        =   87
               Top             =   720
               Visible         =   0   'False
               Width           =   1245
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Vigência"
            Height          =   630
            Left            =   1740
            TabIndex        =   81
            Top             =   195
            Width           =   3480
            Begin VB.Label Label1 
               Caption         =   "Até:"
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
               Index           =   46
               Left            =   1935
               TabIndex        =   85
               Top             =   255
               Width           =   360
            End
            Begin VB.Label VigenciaAte 
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   2310
               TabIndex        =   84
               Top             =   195
               Width           =   1110
            End
            Begin VB.Label Label1 
               Caption         =   "De:"
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
               Index           =   45
               Left            =   210
               TabIndex        =   83
               Top             =   255
               Width           =   345
            End
            Begin VB.Label VigenciaDe 
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   555
               TabIndex        =   82
               Top             =   195
               Width           =   1110
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Inf. Ext. Sigav"
            Height          =   630
            Left            =   90
            TabIndex        =   78
            Top             =   195
            Visible         =   0   'False
            Width           =   1500
            Begin VB.OptionButton OptSigSim 
               Caption         =   "Sim"
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
               Left            =   90
               TabIndex        =   80
               Top             =   300
               Width           =   675
            End
            Begin VB.OptionButton OptSigNao 
               Caption         =   "Não"
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
               Left            =   750
               TabIndex        =   79
               Top             =   300
               Width           =   675
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Controle:"
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
            Index           =   7
            Left            =   2775
            TabIndex        =   262
            Top             =   1125
            Width           =   795
         End
         Begin VB.Label Controle 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3570
            TabIndex        =   261
            Top             =   1065
            Width           =   1590
         End
         Begin VB.Label Idioma 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2310
            TabIndex        =   90
            Top             =   1065
            Width           =   405
         End
         Begin VB.Label Label1 
            Caption         =   "Idioma:"
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
            Index           =   48
            Left            =   1665
            TabIndex        =   89
            Top             =   1110
            Width           =   750
         End
      End
      Begin VB.CommandButton BotaoCancelar 
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
         Height          =   525
         Left            =   0
         TabIndex        =   8
         ToolTipText     =   "Cancela\Reativa o voucher"
         Top             =   4740
         Width           =   1260
      End
      Begin VB.Frame Frame3 
         Caption         =   "Detalhes"
         Height          =   1965
         Left            =   30
         TabIndex        =   36
         Top             =   1260
         Width           =   9150
         Begin VB.CommandButton BotaoAbrirProd 
            Caption         =   "Abrir"
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
            Left            =   8400
            TabIndex        =   7
            Top             =   855
            Width           =   660
         End
         Begin VB.CommandButton BotaoAbrirFat 
            Caption         =   "Abrir"
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
            Left            =   6885
            TabIndex        =   6
            Top             =   510
            Width           =   660
         End
         Begin VB.Label Observacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   645
            Left            =   1125
            TabIndex        =   271
            Top             =   1245
            Width           =   7950
         End
         Begin VB.Label Label1 
            Caption         =   "OBS:"
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
            Index           =   63
            Left            =   585
            TabIndex        =   270
            Top             =   1275
            Width           =   420
         End
         Begin VB.Label Destino 
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   1140
            TabIndex        =   254
            Top             =   870
            Width           =   2940
         End
         Begin VB.Label Label1 
            Caption         =   "Destino:"
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
            Index           =   44
            Left            =   315
            TabIndex        =   253
            Top             =   930
            Width           =   735
         End
         Begin VB.Label Moeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1140
            TabIndex        =   176
            Top             =   165
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Moeda:"
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
            Index           =   25
            Left            =   405
            TabIndex        =   175
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Cambio 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1140
            TabIndex        =   174
            Top             =   510
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Câmbio:"
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
            Index           =   27
            Left            =   345
            TabIndex        =   173
            Top             =   555
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Taf. Unit. Folheto:"
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
            Index           =   70
            Left            =   3810
            TabIndex        =   172
            Top             =   210
            Width           =   1590
         End
         Begin VB.Label TarifaFolheto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5430
            TabIndex        =   171
            Top             =   165
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Taf. Unit.:"
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
            Index           =   69
            Left            =   1830
            TabIndex        =   170
            Top             =   210
            Width           =   915
         End
         Begin VB.Label TarifaUnitaria 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2760
            TabIndex        =   169
            Top             =   165
            Width           =   915
         End
         Begin VB.Label NumeroFat 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5445
            TabIndex        =   43
            Top             =   525
            Width           =   1440
         End
         Begin VB.Label Pax 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   8325
            TabIndex        =   95
            Top             =   165
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Qtde Passageiros:"
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
            Index           =   51
            Left            =   6735
            TabIndex        =   94
            Top             =   210
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Fatura:"
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
            Index           =   10
            Left            =   4770
            TabIndex        =   44
            Top             =   555
            Width           =   750
         End
         Begin VB.Label Produto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5445
            TabIndex        =   42
            Top             =   870
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Produto:"
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
            Index           =   9
            Left            =   4635
            TabIndex        =   41
            Top             =   930
            Width           =   1020
         End
         Begin VB.Label CondPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2745
            TabIndex        =   40
            Top             =   510
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "C.Pagto:"
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
            Index           =   8
            Left            =   1965
            TabIndex        =   39
            Top             =   585
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Básicos"
         Height          =   1290
         Left            =   15
         TabIndex        =   23
         Top             =   -15
         Width           =   9165
         Begin VB.CommandButton BotaoAbrirCli 
            Caption         =   "Abrir"
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
            Left            =   5730
            TabIndex        =   5
            Top             =   525
            Width           =   660
         End
         Begin VB.CommandButton BotaoVou 
            Caption         =   "Vouchers"
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
            Left            =   165
            TabIndex        =   4
            Top             =   345
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto R$:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   71
            Left            =   2160
            TabIndex        =   180
            Top             =   960
            Width           =   855
         End
         Begin VB.Label TarifaMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   179
            Top             =   885
            Width           =   930
         End
         Begin VB.Label ValorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7800
            TabIndex        =   178
            Top             =   885
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Comi.AG:"
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
            Index           =   3
            Left            =   6960
            TabIndex        =   177
            Top             =   930
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "OCR:"
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
            Index           =   52
            Left            =   7305
            TabIndex        =   97
            Top             =   600
            Width           =   495
         End
         Begin VB.Label ValorOcr 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7800
            TabIndex        =   96
            Top             =   540
            Width           =   1305
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   7785
            TabIndex        =   37
            Top             =   180
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Status:"
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
            Index           =   6
            Left            =   7140
            TabIndex        =   38
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   0
            Left            =   690
            TabIndex        =   35
            Top             =   225
            Width           =   435
         End
         Begin VB.Label Label1 
            Caption         =   "Série:"
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
            Index           =   2
            Left            =   1725
            TabIndex        =   34
            Top             =   225
            Width           =   480
         End
         Begin VB.Label LabelNumVou2 
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
            Height          =   315
            Left            =   2910
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   33
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
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
            Index           =   1
            Left            =   495
            TabIndex        =   32
            Top             =   570
            Width           =   630
         End
         Begin VB.Label TipoVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1170
            TabIndex        =   30
            Top             =   180
            Width           =   480
         End
         Begin VB.Label SerieVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2280
            TabIndex        =   29
            Top             =   180
            Width           =   480
         End
         Begin VB.Label NumeroVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3660
            TabIndex        =   28
            Top             =   180
            Width           =   1050
         End
         Begin VB.Label ClienteVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   27
            Top             =   540
            Width           =   4560
         End
         Begin VB.Label DataEmissaoVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   5715
            TabIndex        =   26
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label ValorVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5730
            TabIndex        =   25
            Top             =   885
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "Valor Faturável R$:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   4035
            TabIndex        =   24
            Top             =   960
            Width           =   1860
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   5
            Left            =   4920
            TabIndex        =   31
            Top             =   255
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   66
            Left            =   615
            TabIndex        =   159
            Top             =   960
            Width           =   540
         End
         Begin VB.Label ValorBruto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3030
            TabIndex        =   158
            Top             =   885
            Width           =   930
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   2
      Left            =   165
      TabIndex        =   21
      Top             =   855
      Visible         =   0   'False
      Width           =   9210
      Begin VB.Frame Frame18 
         Caption         =   "Passageiros"
         Height          =   2310
         Left            =   15
         TabIndex        =   193
         Top             =   1125
         Width           =   9165
         Begin MSMask.MaskEdBox PaxValorEmi 
            Height          =   225
            Left            =   4065
            TabIndex        =   274
            Top             =   1335
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PaxValor 
            Height          =   225
            Left            =   4980
            TabIndex        =   275
            Top             =   1320
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin VB.CheckBox PaxTitular 
            Enabled         =   0   'False
            Height          =   225
            Left            =   2655
            TabIndex        =   273
            Top             =   1155
            Width           =   480
         End
         Begin VB.CheckBox PaxCancelado 
            Enabled         =   0   'False
            Height          =   225
            Left            =   2610
            TabIndex        =   272
            Top             =   1470
            Width           =   510
         End
         Begin MSMask.MaskEdBox PaxSexo 
            Height          =   225
            Left            =   8130
            TabIndex        =   199
            Top             =   990
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PaxNumDoc 
            Height          =   225
            Left            =   6345
            TabIndex        =   198
            Top             =   1005
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PaxTipoDoc 
            Height          =   225
            Left            =   5220
            TabIndex        =   197
            Top             =   975
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PaxDataNasc 
            Height          =   225
            Left            =   4095
            TabIndex        =   196
            Top             =   975
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PaxNome 
            Height          =   225
            Left            =   405
            TabIndex        =   195
            Top             =   960
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridPax 
            Height          =   1455
            Left            =   15
            TabIndex        =   194
            Top             =   240
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   2566
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Passageiro Principal"
         Height          =   1065
         Left            =   15
         TabIndex        =   181
         Top             =   45
         Width           =   9180
         Begin VB.CommandButton BotaoAbrirPax 
            Caption         =   "Abrir"
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
            Left            =   8460
            TabIndex        =   182
            ToolTipText     =   "Abre a tela de cliente com o passageiro"
            Top             =   210
            Width           =   660
         End
         Begin VB.Label Passageiro 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1140
            TabIndex        =   192
            Top             =   255
            Width           =   4710
         End
         Begin VB.Label Label1 
            Caption         =   "Passageiro:"
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
            Index           =   11
            Left            =   60
            TabIndex        =   191
            Top             =   300
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
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
            Index           =   12
            Left            =   6015
            TabIndex        =   190
            Top             =   270
            Width           =   690
         End
         Begin VB.Label CliPassageiro 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6795
            TabIndex        =   189
            Top             =   225
            Width           =   1680
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   13
            Left            =   615
            TabIndex        =   188
            Top             =   690
            Width           =   495
         End
         Begin VB.Label CGC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1140
            TabIndex        =   187
            Top             =   645
            Width           =   2145
         End
         Begin VB.Label Label1 
            Caption         =   "Nascimento:"
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
            Index           =   14
            Left            =   3570
            TabIndex        =   186
            Top             =   675
            Width           =   1095
         End
         Begin VB.Label DataNasc 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4695
            TabIndex        =   185
            Top             =   630
            Width           =   1155
         End
         Begin VB.Label CartaoFid 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1080
            TabIndex        =   184
            Top             =   1095
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label Label1 
            Caption         =   "Cartão Fid:"
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
            Index           =   15
            Left            =   60
            TabIndex        =   183
            Top             =   1140
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Endereço Passageiro Principal"
         Height          =   1800
         Left            =   0
         TabIndex        =   45
         Top             =   3450
         Width           =   9180
         Begin VB.Label Label1 
            Caption         =   "Tel2:"
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
            Index           =   24
            Left            =   6315
            TabIndex        =   63
            Top             =   1095
            Width           =   525
         End
         Begin VB.Label Telefone2 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6855
            TabIndex        =   62
            Top             =   1035
            Width           =   2280
         End
         Begin VB.Label Telefone1 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4035
            TabIndex        =   60
            Top             =   1020
            Width           =   2220
         End
         Begin VB.Label Contato 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   59
            Top             =   1800
            Visible         =   0   'False
            Width           =   7785
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   22
            Left            =   330
            TabIndex        =   58
            Top             =   1845
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Email 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   57
            Top             =   1410
            Width           =   7950
         End
         Begin VB.Label Label1 
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
            Height          =   270
            Index           =   21
            Left            =   615
            TabIndex        =   56
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "UF:"
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
            Index           =   20
            Left            =   6450
            TabIndex        =   55
            Top             =   675
            Width           =   330
         End
         Begin VB.Label UF 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6855
            TabIndex        =   54
            Top             =   600
            Width           =   645
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   19
            Left            =   690
            TabIndex        =   53
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label CEP 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   52
            Top             =   1005
            Width           =   2190
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   18
            Left            =   3390
            TabIndex        =   51
            Top             =   630
            Width           =   645
         End
         Begin VB.Label Cidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4035
            TabIndex        =   50
            Top             =   600
            Width           =   2220
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   17
            Left            =   570
            TabIndex        =   49
            Top             =   645
            Width           =   540
         End
         Begin VB.Label Bairro 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   48
            Top             =   600
            Width           =   2190
         End
         Begin VB.Label Endereco 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   47
            Top             =   210
            Width           =   7950
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   16
            Left            =   255
            TabIndex        =   46
            Top             =   255
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Tel1:"
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
            Index           =   23
            Left            =   3600
            TabIndex        =   61
            Top             =   1110
            Width           =   540
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5325
      Index           =   5
      Left            =   225
      TabIndex        =   22
      Top             =   810
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame Frame17 
         BorderStyle     =   0  'None
         Caption         =   "Resumo do Valores"
         Height          =   5445
         Left            =   390
         TabIndex        =   200
         Top             =   -60
         Width           =   8175
         Begin VB.Label Label1 
            Caption         =   "- CMA com OCRs:"
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
            Index           =   79
            Left            =   180
            TabIndex        =   267
            Top             =   2370
            Width           =   2085
         End
         Begin VB.Label OCRCMAReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   266
            Top             =   2340
            Width           =   1245
         End
         Begin VB.Label OCRCMAMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   265
            Top             =   2340
            Width           =   1245
         End
         Begin VB.Line Line12 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Line Line11 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label OCRMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   247
            Top             =   1965
            Width           =   1245
         End
         Begin VB.Label OCRReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   246
            Top             =   1965
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "+ BRUTO com OCRs:"
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
            Index           =   77
            Left            =   180
            TabIndex        =   245
            Top             =   1995
            Width           =   2085
         End
         Begin VB.Label CambioRV 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6600
            TabIndex        =   244
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Câmbio:"
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
            Index           =   76
            Left            =   5895
            TabIndex        =   243
            Top             =   165
            Width           =   975
         End
         Begin VB.Label MoedaRV 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   885
            TabIndex        =   242
            Top             =   135
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Moeda:"
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
            Index           =   75
            Left            =   225
            TabIndex        =   241
            Top             =   180
            Width           =   630
         End
         Begin VB.Label LiqFinalMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   240
            Top             =   5025
            Width           =   1245
         End
         Begin VB.Label LiqFinalReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   239
            Top             =   5025
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "= Líquido Final:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   74
            Left            =   165
            TabIndex        =   238
            Top             =   5040
            Width           =   2205
         End
         Begin VB.Line Line10 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   4980
            Y2              =   4980
         End
         Begin VB.Line Line9 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Label Label1 
            Caption         =   "- CMI:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   73
            Left            =   195
            TabIndex        =   237
            Top             =   4665
            Width           =   975
         End
         Begin VB.Label CMIReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   236
            Top             =   4635
            Width           =   1245
         End
         Begin VB.Label CMIMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   235
            Top             =   4635
            Width           =   1245
         End
         Begin VB.Label TarifaUN 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   234
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "Tarifa UN:"
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
            Index           =   26
            Left            =   2040
            TabIndex        =   233
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Vlr na Moeda"
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
            Index           =   35
            Left            =   4875
            TabIndex        =   232
            Top             =   510
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Percentual"
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
            Index           =   34
            Left            =   3045
            TabIndex        =   231
            Top             =   510
            Width           =   990
         End
         Begin VB.Label CMAPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   230
            Top             =   1185
            Width           =   1245
         End
         Begin VB.Label CMAMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   229
            Top             =   1185
            Width           =   1245
         End
         Begin VB.Label CMAReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   228
            Top             =   1185
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "- CMA sem OCRs:"
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
            Index           =   28
            Left            =   180
            TabIndex        =   227
            Top             =   1215
            Width           =   1680
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Label CMCCPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   226
            Top             =   2715
            Width           =   1245
         End
         Begin VB.Label CMCCMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   225
            Top             =   2715
            Width           =   1245
         End
         Begin VB.Label CMCCReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   224
            Top             =   2715
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "- CMCC:"
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
            Index           =   29
            Left            =   180
            TabIndex        =   223
            Top             =   2745
            Width           =   900
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   1530
            Y2              =   1530
         End
         Begin VB.Label FATMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   222
            Top             =   1575
            Width           =   1245
         End
         Begin VB.Label FATReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   221
            Top             =   1575
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "= Faturável:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   30
            Left            =   180
            TabIndex        =   220
            Top             =   1605
            Width           =   2100
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   2670
            Y2              =   2670
         End
         Begin VB.Label OverPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   219
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label OverMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   218
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label OverReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   217
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "- CME:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   31
            Left            =   180
            TabIndex        =   216
            Top             =   3120
            Width           =   900
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Label CMRPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   215
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Label CMRMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   214
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Label CMRReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   213
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "- CMR:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   32
            Left            =   180
            TabIndex        =   212
            Top             =   3495
            Width           =   915
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   3435
            Y2              =   3435
         End
         Begin VB.Label BrutoMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   211
            Top             =   795
            Width           =   1245
         End
         Begin VB.Label BrutoReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   210
            Top             =   795
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "+ BRUTO sem OCRs:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   33
            Left            =   180
            TabIndex        =   209
            Top             =   825
            Width           =   1995
         End
         Begin VB.Line Line6 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label LIQCMIMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   208
            Top             =   4260
            Width           =   1245
         End
         Begin VB.Label LIQCMIReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   207
            Top             =   4260
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "= Líquido Comissão Interna:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   36
            Left            =   165
            TabIndex        =   206
            Top             =   4275
            Width           =   2475
         End
         Begin VB.Line Line7 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   3825
            Y2              =   3825
         End
         Begin VB.Label Label1 
            Caption         =   "Vlr em Real"
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
            Index           =   37
            Left            =   6720
            TabIndex        =   205
            Top             =   510
            Width           =   1245
         End
         Begin VB.Line Line8 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   4215
            Y2              =   4215
         End
         Begin VB.Label Label1 
            Caption         =   "- CMC:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   72
            Left            =   195
            TabIndex        =   204
            Top             =   3900
            Width           =   975
         End
         Begin VB.Label CMCReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   203
            Top             =   3870
            Width           =   1245
         End
         Begin VB.Label CMCMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   202
            Top             =   3870
            Width           =   1245
         End
         Begin VB.Label CMCPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   201
            Top             =   3870
            Width           =   1245
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   3
      Left            =   180
      TabIndex        =   20
      Top             =   915
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame Frame6 
         Caption         =   "Dados do Pagamento com Cartão de crédito"
         Height          =   3930
         Left            =   645
         TabIndex        =   64
         Top             =   360
         Width           =   7950
         Begin VB.Label DataAutoCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6075
            TabIndex        =   264
            Top             =   2910
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Data da Autorização:"
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
            Index           =   78
            Left            =   4080
            TabIndex        =   263
            Top             =   2940
            Width           =   1920
         End
         Begin VB.Label TitularCPF 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   161
            Top             =   1170
            Width           =   2430
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   62
            Left            =   1095
            TabIndex        =   160
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   43
            Left            =   750
            TabIndex        =   76
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label NumeroCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1650
            TabIndex        =   75
            Top             =   2355
            Width           =   3045
         End
         Begin VB.Label Label1 
            Caption         =   "Parcelas:"
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
            Index           =   42
            Left            =   5070
            TabIndex        =   74
            Top             =   2370
            Width           =   855
         End
         Begin VB.Label NumParcelas 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6075
            TabIndex        =   73
            Top             =   2325
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Autorização:"
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
            Index           =   41
            Left            =   405
            TabIndex        =   72
            Top             =   2940
            Width           =   1185
         End
         Begin VB.Label NumAutorizacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   71
            Top             =   2910
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Validade do Cartão:"
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
            Index           =   40
            Left            =   4185
            TabIndex        =   70
            Top             =   1785
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label ValidadeCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6060
            TabIndex        =   69
            Top             =   1755
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Administradora:"
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
            Index           =   39
            Left            =   180
            TabIndex        =   68
            Top             =   1815
            Width           =   1425
         End
         Begin VB.Label Administradora 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   67
            Top             =   1785
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Titular:"
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
            Index           =   38
            Left            =   900
            TabIndex        =   66
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Titular 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   65
            Top             =   570
            Width           =   6000
         End
      End
   End
   Begin VB.Frame Frame20 
      Caption         =   "Frame20"
      Height          =   330
      Left            =   5235
      TabIndex        =   248
      Top             =   390
      Visible         =   0   'False
      Width           =   2700
      Begin VB.Label Convenio 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   900
         TabIndex        =   260
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Convênio:"
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
         Index           =   49
         Left            =   0
         TabIndex        =   259
         Top             =   60
         Width           =   900
      End
      Begin VB.Label CiaAerea 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1005
         TabIndex        =   252
         Top             =   0
         Width           =   2940
      End
      Begin VB.Label Label1 
         Caption         =   "Cia Aérea:"
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
         Index           =   67
         Left            =   105
         TabIndex        =   251
         Top             =   45
         Width           =   885
      End
      Begin VB.Label Aeroportos 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1005
         TabIndex        =   250
         Top             =   420
         Width           =   2940
      End
      Begin VB.Label Label1 
         Caption         =   "Aeroportos:"
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
         Index           =   68
         Left            =   0
         TabIndex        =   249
         Top             =   465
         Width           =   990
      End
   End
   Begin VB.CommandButton BotaoTrazerVou 
      Height          =   315
      Left            =   3885
      Picture         =   "TRPVouchers.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Trazer Dados"
      Top             =   120
      Width           =   435
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8295
      ScaleHeight     =   450
      ScaleWidth      =   1020
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   30
      Width           =   1080
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   60
         Picture         =   "TRPVouchers.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   540
         Picture         =   "TRPVouchers.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5685
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   10028
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Voucher"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Passageiros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cartão"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissão"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resumo dos Valores"
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
   Begin MSMask.MaskEdBox TipoVouP 
      Height          =   315
      Left            =   645
      TabIndex        =   0
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SerieVouP 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumeroVouP 
      Height          =   315
      Left            =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   315
      Left            =   4770
      TabIndex        =   165
      Top             =   60
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin VB.Label Promotor 
      Caption         =   "Label2"
      Height          =   225
      Left            =   4440
      TabIndex        =   269
      Top             =   90
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Index           =   65
      Left            =   150
      TabIndex        =   154
      Top             =   165
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Série:"
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
      Index           =   64
      Left            =   1110
      TabIndex        =   156
      Top             =   165
      Width           =   480
   End
   Begin VB.Label LabelNumVou 
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
      Height          =   330
      Left            =   2205
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   155
      Top             =   165
      Width           =   750
   End
End
Attribute VB_Name = "TRPVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFrameAtualC As Integer

Dim objGridRepr As AdmGrid
Dim objGridCorr As AdmGrid
Dim objGridEmissor As AdmGrid
Dim objGridPromotor As AdmGrid
Dim objGridPax As AdmGrid
Dim objGridContatos As AdmGrid

Public gobjVoucher As ClassTRPVouchers
Public gcolCMR As New Collection
Public gcolCMC As New Collection
Public gcolCME As New Collection
Public gcolCMP As New Collection

'Colunas do GridRepr
Dim iGrid_DataRepr_Col As Integer
Dim iGrid_Rep_Col As Integer
Dim iGrid_ValorRepr_Col As Integer
Dim iGrid_NumTitRepr_Col As Integer
Dim iGrid_StatusRepr_Col As Integer
Dim iGrid_HistoricoRepr_Col As Integer

Dim iGrid_Corr_Col As Integer
Dim iGrid_DataCorr_Col As Integer
Dim iGrid_ValorCorr_Col As Integer
Dim iGrid_NumTitCorr_Col As Integer
Dim iGrid_StatusCorr_Col As Integer
Dim iGrid_HistoricoCorr_Col As Integer

Dim iGrid_Emi_Col As Integer
Dim iGrid_DataEmissor_Col As Integer
Dim iGrid_ValorEmissor_Col As Integer
Dim iGrid_NumTitEmissor_Col As Integer
Dim iGrid_StatusEmissor_Col As Integer
Dim iGrid_HistoricoEmissor_Col As Integer

Dim iGrid_VendedorPromo_Col As Integer
Dim iGrid_DataPromo_Col As Integer
Dim iGrid_ValorPromoBase_Col As Integer
Dim iGrid_ValorPromoComiss_Col As Integer
Dim iGrid_PercPromoComiss_Col As Integer

Dim iGrid_PaxNome_Col As Integer
Dim iGrid_PaxDataNasc_Col As Integer
Dim iGrid_PaxTipoDoc_Col As Integer
Dim iGrid_PaxNumDoc_Col As Integer
Dim iGrid_PaxSexo_Col As Integer
Dim iGrid_PaxValor_Col As Integer
Dim iGrid_PaxValorEmi_Col As Integer
Dim iGrid_PaxTitular_Col As Integer
Dim iGrid_PaxCancelado_Col As Integer

Dim iGrid_ContatoNome_Col As Integer
Dim iGrid_ContatoTelefone_Col As Integer

Private WithEvents objEventoVoucher As AdmEvento
Attribute objEventoVoucher.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Vouchers"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPVoucher"

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

Private Sub BotaoManutencao_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_BotaoManutencao_Click

    objVou.lNumVou = StrParaLong(NumeroVouP.Text)
    objVou.sSerie = SerieVouP.Text
    objVou.sTipVou = TipoVouP.Text
    
    Call Chama_Tela("TRPVouManu", objVou)

    Exit Sub

Erro_BotaoManutencao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelNumVou_Click()
    Call BotaoVou_Click
End Sub

Private Sub LabelNumVou2_Click()
    Call BotaoVou_Click
End Sub

Private Sub TipoVouP_Change()
    If Len(Trim(TipoVouP.ClipText)) > 0 Then
        If SerieVouP.Visible Then SerieVouP.SetFocus
    End If
End Sub

Private Sub SerieVouP_Change()
    If Len(Trim(SerieVouP.ClipText)) > 0 Then
        If NumeroVouP.Visible Then NumeroVouP.SetFocus
    End If
End Sub

Private Sub TipoVouP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SerieVouP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
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

    'Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoVoucher = Nothing
    
    Set objGridRepr = Nothing
    Set objGridCorr = Nothing
    Set objGridEmissor = Nothing
    Set objGridPromotor = Nothing
    Set objGridPax = Nothing
    Set objGridContatos = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
    
    Set gobjVoucher = Nothing
    Set gcolCMR = Nothing
    Set gcolCMC = Nothing
    Set gcolCME = Nothing
    Set gcolCMP = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190863)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoVoucher = New AdmEvento

    Set objGridRepr = New AdmGrid
    Set objGridCorr = New AdmGrid
    Set objGridEmissor = New AdmGrid
    Set objGridPromotor = New AdmGrid
    Set objGridPax = New AdmGrid
    Set objGridContatos = New AdmGrid

    lErro = Inicializa_Grid_Representante(objGridRepr)
    If lErro <> SUCESSO Then gError 197198

    lErro = Inicializa_Grid_Correntista(objGridCorr)
    If lErro <> SUCESSO Then gError 197215

    lErro = Inicializa_Grid_Emissor(objGridEmissor)
    If lErro <> SUCESSO Then gError 197216

    lErro = Inicializa_Grid_Promotor(objGridPromotor)
    If lErro <> SUCESSO Then gError 197223

    lErro = Inicializa_Grid_Pax(objGridPax)
    If lErro <> SUCESSO Then gError 197223

    lErro = Inicializa_Grid_Contatos(objGridContatos)
    If lErro <> SUCESSO Then gError 197223

    iAlterado = 0
    
    iFrameAtual = 1
    iFrameAtualC = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 197198, 197215, 197216

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190864)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim objVouAux As ClassTRPVouchers

On Error GoTo Erro_Trata_Parametros

    If Not (objVou Is Nothing) Then
    
        Set objVouAux = New ClassTRPVouchers
    
        objVouAux.sSerie = objVou.sSerie
        objVouAux.sTipVou = objVou.sTipVou
        objVouAux.lNumVou = objVou.lNumVou
    
        lErro = Traz_TRPVouchers_Tela(objVouAux)
        If lErro <> SUCESSO Then gError 190865

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 190865

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190866)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objVou As ClassTRPVouchers, Optional bValida As Boolean = True) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    If bValida Then
    
        If Len(Trim(NumeroVou.Caption)) = 0 Then gError 190889
        If Len(Trim(SerieVou.Caption)) = 0 Then gError 190891
        If Len(Trim(TipoVou.Caption)) = 0 Then gError 190892
        
    End If

    objVou.lNumVou = StrParaLong(NumeroVou.Caption)
    objVou.sTipVou = TipoVou.Caption
    objVou.sSerie = SerieVou.Caption
    objVou.dtData = StrParaDate(DataEmissaoVou.Caption)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 190889 To 190892 'ERRO_TRP_NUMVOU_NAO_PREENCHIDO
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_NUMVOU_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190867)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRPVouchers"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objVou, False)
    If lErro <> SUCESSO Then gError 190868

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumVou", objVou.lNumVou, 0, "NumVou"
    colCampoValor.Add "Serie", objVou.sSerie, STRING_TRP_OCR_SERIE, "Serie"
    colCampoValor.Add "TipVou", objVou.sTipVou, STRING_TRP_OCR_TIPOVOU, "TipVou"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 190868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190869)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_Tela_Preenche

    objVou.lNumVou = colCampoValor.Item("NumVou").vValor
    objVou.sSerie = colCampoValor.Item("Serie").vValor
    objVou.sTipVou = colCampoValor.Item("TipVou").vValor

    If objVou.lNumVou <> 0 And Len(Trim(objVou.sSerie)) > 0 And Len(Trim(objVou.sTipVou)) > 0 Then
        
        lErro = Traz_TRPVouchers_Tela(objVou)
        If lErro <> SUCESSO Then gError 190870
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 190870

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190871)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRPVouchers() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRPVouchers

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    TipoVou.Caption = ""
    SerieVou.Caption = ""
    NumeroVou.Caption = ""
    ClienteVou.Caption = ""
    Status.Caption = ""
    BotaoCancelar.Caption = "Cancelar"
    ValorVou.Caption = ""
    ValorOcr.Caption = ""
    DataEmissaoVou.Caption = ""
    Controle.Caption = ""
    CondPagto.Caption = ""
    NumeroFat.Caption = ""
    Produto.Caption = ""
    OptNao.Value = False
    OptSim.Value = False
    DataAutoCC.Caption = ""
    Observacao.Caption = ""
    ValorBruto.Caption = ""
    TarifaUnitaria.Caption = ""
    TarifaFolheto.Caption = ""
    OCRCMAMoeda.Caption = ""
    OCRCMAReal.Caption = ""
    
    Passageiro.Caption = ""
    CliPassageiro.Caption = ""
    CGC.Caption = ""
    DataNasc.Caption = ""
    CartaoFid.Caption = ""
    Endereco.Caption = ""
    Bairro.Caption = ""
    Cidade.Caption = ""
    UF.Caption = ""
    CEP.Caption = ""
    Telefone1.Caption = ""
    Telefone2.Caption = ""
    Email.Caption = ""
    Contato.Caption = ""
    DataAutoCC.Caption = ""
    
    Moeda.Caption = ""
    TarifaUN.Caption = ""
    Cambio.Caption = ""
    MoedaRV.Caption = ""
    CambioRV.Caption = ""

    BrutoMoeda.Caption = ""
    BrutoReal.Caption = ""

    CMAPerc.Caption = ""
    CMAMoeda.Caption = ""
    CMAReal.Caption = ""
    
    CMCCPerc.Caption = ""
    CMCCMoeda.Caption = ""
    CMCCReal.Caption = ""
    
    FATMoeda.Caption = ""
    FATReal.Caption = ""
    
    OverPerc.Caption = ""
    OverMoeda.Caption = ""
    OverReal.Caption = ""
    
    CMRPerc.Caption = ""
    CMRMoeda.Caption = ""
    CMRReal.Caption = ""
    
    CMCPerc.Caption = ""
    CMCMoeda.Caption = ""
    CMCReal.Caption = ""
    
    CMIMoeda.Caption = ""
    CMIReal.Caption = ""
    
    OCRMoeda.Caption = ""
    OCRReal.Caption = ""
    
    TarifaMoeda.Caption = ""
    ValorComissao.Caption = ""
    
    LIQCMIMoeda.Caption = ""
    LIQCMIReal.Caption = ""
    
    LiqFinalMoeda.Caption = ""
    LiqFinalReal.Caption = ""
    
    Titular.Caption = ""
    Administradora.Caption = ""
    ValidadeCC.Caption = ""
    NumAutorizacao.Caption = ""
    NumParcelas.Caption = ""
    NumeroCC.Caption = ""
    
    VigenciaAte.Caption = ""
    VigenciaDe.Caption = ""
    
    OptSigNao.Value = False
    OptSigSim.Value = False

    OptAntcNao.Value = False
    OptAntcSim.Value = False

    Pax.Caption = ""
    Emissor.Caption = ""
    Destino.Caption = ""
    DestinoVou.Caption = ""
    Idioma.Caption = ""
    Convenio.Caption = ""
    
    Representante.Caption = ""
    Promotor.Caption = ""
    Emissor.Caption = ""
    Correntista.Caption = ""
    
    TotalComiPro.Caption = ""
    TotalComiEmi.Caption = ""
    TotalComiRep.Caption = ""
    TotalComiCor.Caption = ""
    
    PercComiEmi.Caption = ""
    PercComiRep.Caption = ""
    PercComiCor.Caption = ""
    
    Set gobjVoucher = Nothing
    Set gcolCMR = New Collection
    Set gcolCMC = New Collection
    Set gcolCME = New Collection
    Set gcolCMP = New Collection
    
    Call Grid_Limpa(objGridRepr)
    Call Grid_Limpa(objGridPromotor)
    Call Grid_Limpa(objGridEmissor)
    Call Grid_Limpa(objGridCorr)
    Call Grid_Limpa(objGridContatos)
    Call Grid_Limpa(objGridPax)

    iAlterado = 0

    Limpa_Tela_TRPVouchers = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRPVouchers:

    Limpa_Tela_TRPVouchers = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190881)

    End Select

    Exit Function

End Function

Function Traz_TRPVouchers_Tela(objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objVendedor As New ClassVendedor
Dim sSiglaDest As String
Dim sDescricaoDest As String

On Error GoTo Erro_Traz_TRPVouchers_Tela
    
    Call Limpa_Tela_TRPVouchers

    NumeroVouP.PromptInclude = False
    NumeroVouP.Text = CStr(objVou.lNumVou)
    NumeroVouP.PromptInclude = True
    SerieVouP.Text = objVou.sSerie
    TipoVouP.Text = objVou.sTipVou
    
    'Lê o TRPVouchers que está sendo Passado
    lErro = CF("TRPVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192646

    If lErro = SUCESSO Then
        
        TipoVou.Caption = objVou.sTipVou
        SerieVou.Caption = objVou.sSerie
        NumeroVou.Caption = objVou.lNumVou
        
        If objVou.iPax <> 0 Then
            Pax.Caption = objVou.iPax
        Else
            Pax.Caption = ""
        End If
        
        objCliente.lCodigo = objVou.lClienteVou
       
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 192649
        
        ClienteVou.Caption = objVou.lClienteVou & SEPARADOR & objCliente.sNomeReduzido
               
        Select Case objVou.iStatus
        
            Case STATUS_TRP_VOU_ABERTO
                Status.Caption = STATUS_TRP_VOU_ABERTO_TEXTO
                BotaoCancelar.Caption = "Cancelar"
            
            Case STATUS_TRP_VOU_CANCELADO
                Status.Caption = STATUS_TRP_VOU_CANCELADO_TEXTO
                BotaoCancelar.Caption = "Reativar"
            
        End Select
        
        ValorVou.Caption = Format(objVou.dValor, "STANDARD")
        ValorOcr.Caption = Format(objVou.dValorOcr, "STANDARD")
        DataEmissaoVou.Caption = Format(objVou.dtData, "dd/mm/yyyy")
        Controle.Caption = objVou.sControle
        
        objCondicaoPagto.iCodigo = objCliente.iCondicaoPagto

        'Lê Condição Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 192650
        
        CondPagto.Caption = objCliente.iCondicaoPagto & SEPARADOR & objCondicaoPagto.sDescReduzida
        
        If objVou.lNumFat <> 0 Then
            NumeroFat.Caption = objVou.lNumFat
        Else
            NumeroFat.Caption = ""
        End If
        
        lErro = CF("Produto_Formata", objVou.sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 192650
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 192650
        
        Produto.Caption = objVou.sProduto & SEPARADOR & objProduto.sDescricao
        
        If objVou.iCartao = MARCADO Then
            OptSim.Value = True
        Else
            OptNao.Value = True
        End If
        
        If objVou.iIdioma = 0 Then
            Idioma.Caption = "P"
        Else
            Idioma.Caption = "I"
        End If
        
        Titular.Caption = objVou.sTitular
        Administradora.Caption = objVou.sCiaCart
        'ValidadeCC.Caption = objVou.sValidade
        
        If objVou.lNumAuto > 0 Then
            NumAutorizacao.Caption = objVou.lNumAuto
        Else
            NumAutorizacao.Caption = ""
        End If
        
        If objVou.iQuantParc > 0 Then
            NumParcelas.Caption = objVou.iQuantParc
        Else
            NumParcelas.Caption = ""
        End If
        
        If Len(Trim(objVou.sNumCCred)) > 6 Then
            NumeroCC.Caption = String(Len(Trim(objVou.sNumCCred)) - 6, "X") & right(Trim(objVou.sNumCCred), 6)
        Else
            NumeroCC.Caption = ""
        End If
        
        'Lê o TRPVouchers que está sendo Passado
        lErro = CF("TRPVoucherInfo_Le", objVou)
        If lErro <> SUCESSO Then gError 197212
        
        If objVou.lCorrentista <> 0 Then
        
            Set objCliente = New ClassCliente
             
            objCliente.lCodigo = objVou.lCorrentista
            
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 192649
        
            Correntista.Caption = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
            PercComiCor.Caption = Format(objVou.dComissaoCorr, "PERCENT")
            
        End If
        
        If objVou.lRepresentante <> 0 Then
        
            Set objCliente = New ClassCliente
             
            objCliente.lCodigo = objVou.lRepresentante
            
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 192649
        
            Representante.Caption = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
            PercComiRep.Caption = Format(objVou.dComissaoRep, "PERCENT")
            
        End If
        
        If objVou.lEmissor <> 0 Then
                     
            objFornecedor.lCodigo = objVou.lEmissor
            
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 192649
        
            Emissor.Caption = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
            PercComiEmi.Caption = Format(objVou.dComissaoEmissor, "PERCENT")
            
        End If
        
        'Busca o Cliente no BD
        If objVou.iPromotor <> 0 Then
            Cliente.Text = objVou.iPromotor
            lErro = TP_Vendedor_Le2(Cliente, objVendedor)
            If lErro <> SUCESSO Then gError 192649
            Promotor.Caption = Cliente.Text
        End If
        
'        If objVou.dValorBruto <> 0 Then
            ValorBruto.Caption = Format(objVou.dValorBruto, "STANDARD")
'        End If
        
'        If objVou.dValorCambio <> 0 Then
            TarifaMoeda.Caption = Format(objVou.dValorCambio, "STANDARD")
'        End If
        
'        If objVou.dValorComissao <> 0 Then
            ValorComissao.Caption = Format(objVou.dValorComissao, "STANDARD")
'        End If
        
        If objVou.dtDataAutoCC <> DATA_NULA Then
            DataAutoCC.Caption = Format(objVou.dtDataAutoCC, "dd/mm/yyyy")
        End If
               
        Select Case Len(Trim(objVou.sTitularCPF))
    
            Case STRING_CPF 'CPF
                TitularCPF.Caption = Format(objVou.sTitularCPF, "000\.000\.000-00; ; ; ")
    
            Case STRING_CGC 'CGC
                TitularCPF.Caption = Format(objVou.sTitularCPF, "00\.000\.000\/0000-00; ; ; ")
                
            Case Else
                TitularCPF.Caption = objVou.sTitularCPF
            
        End Select
        
        Select Case Len(Trim(objVou.sPassageiroCGC))
    
            Case STRING_CPF 'CPF
                CGC.Caption = Format(objVou.sPassageiroCGC, "000\.000\.000-00; ; ; ")
    
            Case STRING_CGC 'CGC
                CGC.Caption = Format(objVou.sPassageiroCGC, "00\.000\.000\/0000-00; ; ; ")
                
            Case Else
                CGC.Caption = objVou.sPassageiroCGC
            
        End Select
    
        Passageiro.Caption = objVou.sPassageiroNome & " " & objVou.sPassageiroSobreNome
        
        If objVou.lCliPassageiro <> 0 Then
            CliPassageiro.Caption = CStr(objVou.lCliPassageiro)
        Else
            CliPassageiro.Caption = ""
        End If
        
        If objVou.dtPassageiroDataNasc <> DATA_NULA Then
            DataNasc.Caption = Format(objVou.dtPassageiroDataNasc, "dd/mm/yyyy")
        Else
            DataNasc.Caption = ""
        End If
        
        CiaAerea.Caption = objVou.sCiaaerea
        Aeroportos.Caption = objVou.sAeroportos
        Endereco.Caption = objVou.objEnderecoPax.sEndereco
        Bairro.Caption = objVou.objEnderecoPax.sBairro
        Cidade.Caption = objVou.objEnderecoPax.sCidade
        UF.Caption = objVou.objEnderecoPax.sSiglaEstado
        CEP.Caption = objVou.objEnderecoPax.sCEP
        Telefone1.Caption = objVou.objEnderecoPax.sTelefone1
        Telefone2.Caption = objVou.objEnderecoPax.sTelefone2
        Email.Caption = objVou.objEnderecoPax.sEmail
        Contato.Caption = objVou.objEnderecoPax.sContato
        Observacao.Caption = objVou.sObservacao
        
        If objVou.iMoeda = MOEDA_REAL Then
            Moeda.Caption = "BRL"
        ElseIf objVou.iMoeda = MOEDA_DOLAR Then
            Moeda.Caption = "USD"
        Else
            Moeda.Caption = "BRL"
        End If
        
        TarifaUnitaria.Caption = Format(objVou.dTarifaUnitaria, "STANDARD")
        TarifaFolheto.Caption = Format(objVou.dTarifaUnitariaFolheto, "STANDARD")
        
        If objVou.dCambio <> 0 Then
            Cambio.Caption = Format(objVou.dCambio, "0.0000")
        Else
            Cambio.Caption = ""
        End If
        
        If Len(Trim(objVou.sNumCCred)) > 6 Then
            NumeroCC.Caption = String(Len(Trim(objVou.sNumCCred)) - 6, "X") & right(Trim(objVou.sNumCCred), 6)
        Else
            NumeroCC.Caption = ""
        End If
        
        If objVou.dtDataVigenciaDe <> DATA_NULA Then
            VigenciaDe.Caption = Format(objVou.dtDataVigenciaDe, "dd/mm/yyyy")
        Else
            VigenciaDe.Caption = ""
        End If
        
        If objVou.dtDataVigenciaAte <> DATA_NULA Then
            VigenciaAte.Caption = Format(objVou.dtDataVigenciaAte, "dd/mm/yyyy")
        Else
            VigenciaAte.Caption = ""
        End If
        
        If objVou.iDiasAntc = MARCADO Then
            OptAntcSim.Value = True
        Else
            OptAntcNao.Value = True
        End If
        
        lErro = CF("TRPDestino_Le", objVou.iDestino, sSiglaDest, sDescricaoDest)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192646
        
        Destino.Caption = objVou.iDestino & SEPARADOR & sDescricaoDest
        
        lErro = Traz_Grids_Tela(objVou)
        If lErro <> SUCESSO Then gError 197213
        
    End If
    
    Set gobjVoucher = objVou

    iAlterado = 0

    Traz_TRPVouchers_Tela = SUCESSO

    Exit Function

Erro_Traz_TRPVouchers_Tela:

    Traz_TRPVouchers_Tela = gErr

    Select Case gErr

        Case 192646 To 192650, 197212, 197213

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190883)

    End Select

    Exit Function

End Function

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190886)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_TRPVouchers

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190888)

    End Select

    Exit Sub

End Sub

Sub BotaoCancelar_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoCancelar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 190893
    
    lErro = CF("TRPVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198200
    
    If lErro <> SUCESSO Then gError 198201
    
    If UCase(BotaoCancelar.Caption) = "CANCELAR" Then
    
        If Status.Caption = STATUS_TRP_OCR_CANCELADO_TEXTO Then gError 192650
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_CANCELAMENTO_TRPVOUCHERS")
    
        If vbMsgRes = vbYes Then
    
            'Cancela o voucher
            lErro = CF("TRPVoucher_Exclui", objVou)
            If lErro <> SUCESSO Then gError 190894
    
            'Limpa Tela
            Call Limpa_Tela_TRPVouchers
            
            Call Trata_Parametros(objVou)
            
            Call Rotina_Aviso(vbOKOnly, "AVISO_VOUCHER_CANCELADO")
    
        End If
        
    Else
    
        If Status.Caption = STATUS_TRP_VOU_ABERTO_TEXTO Then gError 200711
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_REATIVAMENTO_TRPVOUCHERS")
    
        If vbMsgRes = vbYes Then
    
            'Reativa o voucher
            lErro = CF("TRPVoucher_Reativa", objVou)
            If lErro <> SUCESSO Then gError 190894
    
            'Limpa Tela
            Call Limpa_Tela_TRPVouchers
            
            Call Trata_Parametros(objVou)
            
            Call Rotina_Aviso(vbOKOnly, "AVISO_VOUCHER_REATIVADO")
    
        End If
        
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoCancelar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 190893, 190894, 198200
        
        Case 192650
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)

        Case 198201
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)

        Case 200711
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_ATIVO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190895)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVoucher_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVou As ClassTRPVouchers

On Error GoTo Erro_objEventoVoucher_evSelecao

    Set objVou = obj1

    'Mostra os dados do TRPVouchers na tela
    lErro = Traz_TRPVouchers_Tela(objVou)
    If lErro <> SUCESSO Then gError 190909

    Me.Show

    Exit Sub

Erro_objEventoVoucher_evSelecao:

    Select Case gErr

        Case 190909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143951)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVou.Caption)
    objVoucher.sSerie = SerieVou.Caption
    objVoucher.sTipVou = TipoVou.Caption

    Call Chama_Tela("TRPVoucherRapidoLista", colSelecao, objVoucher, objEventoVoucher)

    Exit Sub

Erro_BotaoVou_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190160)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190634)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip2_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip2)
End Sub

Private Sub TabStrip2_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip2_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index = iFrameAtualC Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtualC, TabStrip2, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    FrameC(TabStrip2.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    FrameC(iFrameAtualC).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtualC = TabStrip2.SelectedItem.Index

    Exit Sub

Erro_TabStrip2_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190634)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAbrirCli_Click()

Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirCli_Click

    objCliente.lCodigo = LCodigo_Extrai(ClienteVou.Caption)

    Call Chama_Tela("Clientes", objCliente)

    Exit Sub

Erro_BotaoAbrirCli_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirFat_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_BotaoAbrirFat_Click

    If Len(Trim(NumeroFat.Caption)) <> 0 Then

        objVou.lNumVou = NumeroVou.Caption
        objVou.sTipVou = TipoVou.Caption
        objVou.sSerie = SerieVou.Caption

        lErro = CF("TRPVouchers_Le", objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192882
        
        lErro = Abrir_Doc_Destino(objVou.lNumIntDocDestino, objVou.iTipoDocDestino)
        If lErro <> SUCESSO Then gError 192883
        
    End If
    
    Exit Sub

Erro_BotaoAbrirFat_Click:

    Select Case gErr
    
        Case 192882, 192883
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192884)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirPax_Click()

Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirPax_Click

    objCliente.lCodigo = LCodigo_Extrai(CliPassageiro.Caption)

    Call Chama_Tela("Clientes", objCliente)

    Exit Sub

Erro_BotaoAbrirPax_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192885)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirProd_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoAbrirProd_Click

    lErro = CF("Produto_Formata", SCodigo_Extrai(Produto.Caption), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 177323
    
    objProduto.sCodigo = sProdutoFormatado

    Call Chama_Tela("Produto", objProduto)

    Exit Sub

Erro_BotaoAbrirProd_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192886)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirEmi_Click()

Dim objForn As New ClassFornecedor

On Error GoTo Erro_BotaoAbrirEmi_Click

    objForn.lCodigo = LCodigo_Extrai(Emissor.Caption)

    Call Chama_Tela("Fornecedores", objForn)

    Exit Sub

Erro_BotaoAbrirEmi_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHistOcor_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistOcor_Click

    colSelecao.Add StrParaLong(NumeroVou.Caption)
    colSelecao.Add TipoVou.Caption
    colSelecao.Add SerieVou.Caption

    Call Chama_Tela("OcorrenciasHistLista", colSelecao, Nothing, Nothing)

    Exit Sub

Erro_BotaoHistOcor_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Representante(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Rep.")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Título")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Histórico")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Rep.Name)
    objGridInt.colCampo.Add (DataRepr.Name)
    objGridInt.colCampo.Add (ValorRepr.Name)
    objGridInt.colCampo.Add (NumTitRepr.Name)
    objGridInt.colCampo.Add (StatusRepr.Name)
    objGridInt.colCampo.Add (HistoricoRepr.Name)

    'Colunas do GridRepr
    iGrid_Rep_Col = 1
    iGrid_DataRepr_Col = 2
    iGrid_ValorRepr_Col = 3
    iGrid_NumTitRepr_Col = 4
    iGrid_StatusRepr_Col = 5
    iGrid_HistoricoRepr_Col = 6

    'Grid do GridInterno
    objGridInt.objGrid = GridRepr

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridRepr.ColWidth(0) = 200

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Representante = SUCESSO

    Exit Function

End Function

Public Sub GridRepr_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRepr, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRepr, iAlterado)
    End If
    
End Sub

Public Sub GridRepr_EnterCell()
    Call Grid_Entrada_Celula(objGridRepr, iAlterado)
End Sub

Public Sub GridRepr_GotFocus()
    Call Grid_Recebe_Foco(objGridRepr)
End Sub

Public Sub GridRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridRepr)
    
End Sub

Public Sub GridRepr_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRepr, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRepr, iAlterado)
    End If
    
End Sub

Public Sub GridRepr_LeaveCell()
    Call Saida_Celula(objGridRepr)
End Sub

Public Sub GridRepr_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridRepr)
End Sub

Public Sub GridRepr_RowColChange()
    Call Grid_RowColChange(objGridRepr)
End Sub

Public Sub GridRepr_Scroll()
    Call Grid_Scroll(objGridRepr)
End Sub

Private Function Inicializa_Grid_Correntista(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Corr.")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Título")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Histórico")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Corr.Name)
    objGridInt.colCampo.Add (DataCorr.Name)
    objGridInt.colCampo.Add (ValorCorr.Name)
    objGridInt.colCampo.Add (NumTitCorr.Name)
    objGridInt.colCampo.Add (StatusCorr.Name)
    objGridInt.colCampo.Add (HistoricoCorr.Name)

    'Colunas do GridRepr
    iGrid_Corr_Col = 1
    iGrid_DataCorr_Col = 2
    iGrid_ValorCorr_Col = 3
    iGrid_NumTitCorr_Col = 4
    iGrid_StatusCorr_Col = 5
    iGrid_HistoricoCorr_Col = 6

    'Grid do GridInterno
    objGridInt.objGrid = GridCorr

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridCorr.ColWidth(0) = 200
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Correntista = SUCESSO

    Exit Function

End Function

Public Sub GridCorr_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCorr, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCorr, iAlterado)
    End If
    
End Sub

Public Sub GridCorr_EnterCell()
    Call Grid_Entrada_Celula(objGridCorr, iAlterado)
End Sub

Public Sub GridCorr_GotFocus()
    Call Grid_Recebe_Foco(objGridCorr)
End Sub

Public Sub GridCorr_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridCorr)
    
End Sub

Public Sub GridCorr_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCorr, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCorr, iAlterado)
    End If
    
End Sub

Public Sub GridCorr_LeaveCell()
    Call Saida_Celula(objGridCorr)
End Sub

Public Sub GridCorr_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridCorr)
End Sub

Public Sub GridCorr_RowColChange()
    Call Grid_RowColChange(objGridCorr)
End Sub

Public Sub GridCorr_Scroll()
    Call Grid_Scroll(objGridCorr)
End Sub

Private Function Inicializa_Grid_Emissor(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Emis.")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Título")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Histórico")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Emi.Name)
    objGridInt.colCampo.Add (DataEmissor.Name)
    objGridInt.colCampo.Add (ValorEmissor.Name)
    objGridInt.colCampo.Add (NumTitEmissor.Name)
    objGridInt.colCampo.Add (StatusEmissor.Name)
    objGridInt.colCampo.Add (HistoricoEmissor.Name)

    'Colunas do GridRepr
    iGrid_Emi_Col = 1
    iGrid_DataEmissor_Col = 2
    iGrid_ValorEmissor_Col = 3
    iGrid_NumTitEmissor_Col = 4
    iGrid_StatusEmissor_Col = 5
    iGrid_HistoricoEmissor_Col = 6

    'Grid do GridInterno
    objGridInt.objGrid = GridEmissor

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Largura da primeira coluna
    GridEmissor.ColWidth(0) = 200

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Emissor = SUCESSO

    Exit Function

End Function

Public Sub GridEmissor_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridEmissor, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEmissor, iAlterado)
    End If
    
End Sub

Public Sub GridEmissor_EnterCell()
    Call Grid_Entrada_Celula(objGridEmissor, iAlterado)
End Sub

Public Sub GridEmissor_GotFocus()
    Call Grid_Recebe_Foco(objGridEmissor)
End Sub

Public Sub GridEmissor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridEmissor)
    
End Sub

Public Sub GridEmissor_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridEmissor, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEmissor, iAlterado)
    End If
    
End Sub

Public Sub GridEmissor_LeaveCell()
    Call Saida_Celula(objGridEmissor)
End Sub

Public Sub GridEmissor_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridEmissor)
End Sub

Public Sub GridEmissor_RowColChange()
    Call Grid_RowColChange(objGridEmissor)
End Sub

Public Sub GridEmissor_Scroll()
    Call Grid_Scroll(objGridEmissor)
End Sub

Private Function Inicializa_Grid_Promotor(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Valor Comissão")
    objGridInt.colColuna.Add ("% Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (VendedorPromo.Name)
    objGridInt.colCampo.Add (DataPromo.Name)
    objGridInt.colCampo.Add (ValorPromoBase.Name)
    objGridInt.colCampo.Add (ValorPromoComiss.Name)
    objGridInt.colCampo.Add (PercPromoComiss.Name)

    'Colunas do GridRepr
    iGrid_VendedorPromo_Col = 1
    iGrid_DataPromo_Col = 2
    iGrid_ValorPromoBase_Col = 3
    iGrid_ValorPromoComiss_Col = 4
    iGrid_PercPromoComiss_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridPromotor

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridPromotor.ColWidth(0) = 400

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Promotor = SUCESSO

    Exit Function

End Function

Public Sub GridPromotor_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPromotor, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPromotor, iAlterado)
    End If
    
End Sub

Public Sub GridPromotor_EnterCell()
    Call Grid_Entrada_Celula(objGridPromotor, iAlterado)
End Sub

Public Sub GridPromotor_GotFocus()
    Call Grid_Recebe_Foco(objGridPromotor)
End Sub

Public Sub GridPromotor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridPromotor)
    
End Sub

Public Sub GridPromotor_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPromotor, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPromotor, iAlterado)
    End If
    
End Sub

Public Sub GridPromotor_LeaveCell()
    Call Saida_Celula(objGridPromotor)
End Sub

Public Sub GridPromotor_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridPromotor)
End Sub

Public Sub GridPromotor_RowColChange()
    Call Grid_RowColChange(objGridPromotor)
End Sub

Public Sub GridPromotor_Scroll()
    Call Grid_Scroll(objGridPromotor)
End Sub

Private Function Traz_Grids_Tela(objVou As ClassTRPVouchers) As Long

Dim objVouInfo As ClassTRPVoucherInfo
Dim iLinha As Integer
Dim iLinha1 As Integer
Dim iLinha2 As Integer
Dim objTRPGerComiIntDet As ClassTRPGerComiIntDet
Dim lErro As Long
Dim dValorR As Double
Dim dValorC As Double
Dim dValorCC As Double
Dim dValorA As Double
Dim dValorB As Double
Dim dValorO As Double
Dim dValorE As Double
Dim dValorP As Double
Dim dValorI As Double
Dim objCliente As New ClassCliente
Dim objForn As New ClassFornecedor
Dim objPax As ClassTRPVouPassageiros
Dim objContato As ClassTRPVouContatos

On Error GoTo Erro_Traz_Grids_Tela

    Call Grid_Limpa(objGridRepr)
    Call Grid_Limpa(objGridPromotor)
    Call Grid_Limpa(objGridEmissor)
    Call Grid_Limpa(objGridCorr)
    Call Grid_Limpa(objGridPax)
    Call Grid_Limpa(objGridContatos)
    
    Set gcolCMR = New Collection
    Set gcolCMC = New Collection
    Set gcolCME = New Collection
    Set gcolCMP = New Collection

    For Each objVouInfo In objVou.colTRPVoucherInfo
    
        If objVouInfo.sTipoDoc = TRP_TIPODOC_CMCC_TEXTO Then
            dValorCC = dValorCC + objVouInfo.dValor
        End If
        
        If objVouInfo.sTipoDoc = TRP_TIPODOC_CMA_TEXTO Then
            dValorA = dValorA + objVouInfo.dValor
        End If

        If objVouInfo.sTipoDoc = TRP_TIPODOC_BRUTO_TEXTO Then
            dValorB = dValorB + objVouInfo.dValor
        End If
        
        If objVouInfo.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Then
            
            iLinha = iLinha + 1
            
            'Busca o Cliente no BD
            If objVouInfo.lCliForn <> 0 Then
                Cliente.Text = objVouInfo.lCliForn
                lErro = TP_Cliente_Le2(Cliente, objCliente)
                If lErro <> SUCESSO Then gError 198232
            End If
            
            GridRepr.TextMatrix(iLinha, iGrid_Rep_Col) = Cliente.Text
            GridRepr.TextMatrix(iLinha, iGrid_DataRepr_Col) = Format(objVouInfo.dtData, "dd/mm/yyyy")
            GridRepr.TextMatrix(iLinha, iGrid_ValorRepr_Col) = Format(objVouInfo.dValor, "Standard")
            
            If objVouInfo.lNumTitulo <> 0 Then
                GridRepr.TextMatrix(iLinha, iGrid_NumTitRepr_Col) = CStr(objVouInfo.lNumTitulo)
            Else
                GridRepr.TextMatrix(iLinha, iGrid_NumTitRepr_Col) = ""
            End If
            
            GridRepr.TextMatrix(iLinha, iGrid_HistoricoRepr_Col) = objVouInfo.sHistorico
            GridRepr.TextMatrix(iLinha, iGrid_StatusRepr_Col) = TRPVoucherInfo_Converte_Status(objVouInfo.iStatus)
    
            dValorR = dValorR + objVouInfo.dValor
    
            gcolCMR.Add objVouInfo
    
        End If
            
        If objVouInfo.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Then
            
            iLinha1 = iLinha1 + 1
            
            'Busca o Cliente no BD
            If objVouInfo.lCliForn <> 0 Then
                Cliente.Text = objVouInfo.lCliForn
                lErro = TP_Cliente_Le2(Cliente, objCliente)
                If lErro <> SUCESSO Then gError 198232
            End If
            
            GridCorr.TextMatrix(iLinha1, iGrid_Corr_Col) = Cliente.Text
            GridCorr.TextMatrix(iLinha1, iGrid_DataCorr_Col) = Format(objVouInfo.dtData, "dd/mm/yyyy")
            GridCorr.TextMatrix(iLinha1, iGrid_ValorCorr_Col) = Format(objVouInfo.dValor, "Standard")
            
            If objVouInfo.lNumTitulo <> 0 Then
                GridCorr.TextMatrix(iLinha1, iGrid_NumTitCorr_Col) = CStr(objVouInfo.lNumTitulo)
            Else
                GridCorr.TextMatrix(iLinha1, iGrid_NumTitCorr_Col) = ""
            End If
            
            GridCorr.TextMatrix(iLinha1, iGrid_HistoricoCorr_Col) = objVouInfo.sHistorico
            GridCorr.TextMatrix(iLinha1, iGrid_StatusCorr_Col) = TRPVoucherInfo_Converte_Status(objVouInfo.iStatus)
    
            dValorC = dValorC + objVouInfo.dValor
    
            gcolCMC.Add objVouInfo
    
        End If
            
        If objVouInfo.sTipoDoc = TRP_TIPODOC_OVER_TEXTO Then
            
            iLinha2 = iLinha2 + 1
            
            'Busca o Cliente no BD
            If objVouInfo.lCliForn <> 0 Then
                Cliente.Text = objVouInfo.lCliForn
                lErro = TP_Fornecedor_Le2(Cliente, objForn)
                If lErro <> SUCESSO Then gError 198232
            End If
            
            GridEmissor.TextMatrix(iLinha2, iGrid_Emi_Col) = Cliente.Text
            GridEmissor.TextMatrix(iLinha2, iGrid_DataEmissor_Col) = Format(objVouInfo.dtData, "dd/mm/yyyy")
            GridEmissor.TextMatrix(iLinha2, iGrid_ValorEmissor_Col) = Format(objVouInfo.dValor, "Standard")
            
            If objVouInfo.lNumTitulo <> 0 Then
                GridEmissor.TextMatrix(iLinha2, iGrid_NumTitEmissor_Col) = CStr(objVouInfo.lNumTitulo)
            Else
                GridEmissor.TextMatrix(iLinha2, iGrid_NumTitEmissor_Col) = ""
            End If
            
            GridEmissor.TextMatrix(iLinha2, iGrid_HistoricoEmissor_Col) = objVouInfo.sHistorico
            GridEmissor.TextMatrix(iLinha2, iGrid_StatusEmissor_Col) = TRPVoucherInfo_Converte_Status(objVouInfo.iStatus)
    
            dValorE = dValorE + objVouInfo.dValor
    
            gcolCME.Add objVouInfo
    
        End If
    
    Next

    lErro = CF("TRPVoucher_Le_ComisInt", objVou)
    If lErro <> SUCESSO Then gError 197222
    
    objGridRepr.iLinhasExistentes = iLinha
    objGridCorr.iLinhasExistentes = iLinha1
    objGridEmissor.iLinhasExistentes = iLinha2
    
    iLinha = 0
    
    For Each objTRPGerComiIntDet In objVou.colTRPGerComiIntDet
    
        iLinha = iLinha + 1
        
        GridPromotor.TextMatrix(iLinha, iGrid_VendedorPromo_Col) = objTRPGerComiIntDet.iVendedor & SEPARADOR & objTRPGerComiIntDet.sNomeReduzidoVendedor
        GridPromotor.TextMatrix(iLinha, iGrid_DataPromo_Col) = Format(objTRPGerComiIntDet.dtDataGeracao, "dd/mm/yyyy")
        GridPromotor.TextMatrix(iLinha, iGrid_ValorPromoBase_Col) = Format(objTRPGerComiIntDet.dValorBase, "Standard")
        GridPromotor.TextMatrix(iLinha, iGrid_ValorPromoComiss_Col) = Format(objTRPGerComiIntDet.dValorComissao, "Standard")
        GridPromotor.TextMatrix(iLinha, iGrid_PercPromoComiss_Col) = Format(objTRPGerComiIntDet.dPercComissao, "Percent")
    
        dValorP = dValorP + objTRPGerComiIntDet.dValorComissao
    
        gcolCMP.Add objTRPGerComiIntDet
    
    Next
    
    objGridPromotor.iLinhasExistentes = iLinha
    
    iLinha = 0
    
    For Each objPax In objVou.colPassageiros
    
        iLinha = iLinha + 1
        
        If objPax.iStatus = STATUS_TRP_VOU_CANCELADO Then
            GridPax.TextMatrix(iLinha, iGrid_PaxCancelado_Col) = MARCADO
        Else
            GridPax.TextMatrix(iLinha, iGrid_PaxCancelado_Col) = DESMARCADO
        End If
        
        GridPax.TextMatrix(iLinha, iGrid_PaxTitular_Col) = objPax.iTitular
        GridPax.TextMatrix(iLinha, iGrid_PaxNome_Col) = objPax.sNome
        GridPax.TextMatrix(iLinha, iGrid_PaxDataNasc_Col) = Format(objPax.dtDataNascimento, "dd/mm/yyyy")
        GridPax.TextMatrix(iLinha, iGrid_PaxTipoDoc_Col) = objPax.sTipoDocumento
        GridPax.TextMatrix(iLinha, iGrid_PaxNumDoc_Col) = objPax.sNumeroDocumento
        GridPax.TextMatrix(iLinha, iGrid_PaxSexo_Col) = objPax.sSexo
       
        GridPax.TextMatrix(iLinha, iGrid_PaxValor_Col) = Format(objPax.dValorPago, "STANDARD")
        GridPax.TextMatrix(iLinha, iGrid_PaxValorEmi_Col) = Format(objPax.dValorPagoEmi, "STANDARD")
       
    Next
    
    objGridPax.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridPax)
    
    iLinha = 0
    
    For Each objContato In objVou.colContatos
    
        iLinha = iLinha + 1
        
        GridContatos.TextMatrix(iLinha, iGrid_ContatoNome_Col) = objContato.sNome
        GridContatos.TextMatrix(iLinha, iGrid_ContatoTelefone_Col) = objContato.sTelefone
       
    Next
    
    objGridContatos.iLinhasExistentes = iLinha
    
    TotalComiPro.Caption = Format(dValorP, "STANDARD")
    TotalComiEmi.Caption = Format(dValorE, "STANDARD")
    TotalComiRep.Caption = Format(dValorR, "STANDARD")
    TotalComiCor.Caption = Format(dValorC, "STANDARD")
    
    MoedaRV.Caption = Moeda.Caption
    TarifaUN.Caption = TarifaUnitaria.Caption
    CambioRV.Caption = Cambio.Caption
    
    CMIReal.Caption = Format(dValorP, "STANDARD")
    OverReal.Caption = Format(dValorE, "STANDARD")
    CMRReal.Caption = Format(dValorR, "STANDARD")
    CMCReal.Caption = Format(dValorC, "STANDARD")
    OCRReal.Caption = Format(dValorB, "STANDARD")
    OCRCMAReal.Caption = Format(dValorA, "STANDARD")
    BrutoReal.Caption = Format(objVou.dValorBruto, "STANDARD")
    CMCCReal.Caption = Format(dValorCC, "STANDARD")
    
    If objVou.iCartao = DESMARCADO Then
        CMAReal.Caption = Format(objVou.dValorComissao, "STANDARD")
    Else
        CMAReal.Caption = Format(0, "STANDARD")
    End If
    
    FATReal.Caption = Format(objVou.dValor, "STANDARD")
    LIQCMIReal.Caption = Format(dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC + dValorO, "STANDARD")
    LiqFinalReal.Caption = Format(dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC + dValorO - dValorP, "STANDARD")
    
    If objVou.dCambio <> 0 Then
        CMIMoeda.Caption = Format(dValorP / objVou.dCambio, "STANDARD")
        OverMoeda.Caption = Format(dValorE / objVou.dCambio, "STANDARD")
        CMRMoeda.Caption = Format(dValorR / objVou.dCambio, "STANDARD")
        CMCMoeda.Caption = Format(dValorC / objVou.dCambio, "STANDARD")
        OCRMoeda.Caption = Format(dValorB / objVou.dCambio, "STANDARD")
        OCRCMAMoeda.Caption = Format(dValorA / objVou.dCambio, "STANDARD")
        BrutoMoeda.Caption = Format(objVou.dValorCambio, "STANDARD")
        CMCCMoeda.Caption = Format(dValorCC / objVou.dCambio, "STANDARD")
        
        If objVou.iCartao = DESMARCADO Then
            CMAMoeda.Caption = Format(objVou.dValorComissao / objVou.dCambio, "STANDARD")
        Else
            CMAMoeda.Caption = Format(0, "STANDARD")
        End If
    
        FATMoeda.Caption = Format(objVou.dValor / objVou.dCambio, "STANDARD")
        LIQCMIMoeda.Caption = Format((dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC + dValorO) / objVou.dCambio, "STANDARD")
        LiqFinalMoeda.Caption = Format((dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC + dValorO - dValorP) / objVou.dCambio, "STANDARD")
    
    End If
    
    If dValorB <> 0 Then
        OverPerc.Caption = Format(dValorE / (dValorB + dValorO), "PERCENT")
        CMRPerc.Caption = Format(dValorR / (dValorB + dValorO), "PERCENT")
        CMCPerc.Caption = Format(dValorC / (dValorB + dValorO), "PERCENT")
        CMCCPerc.Caption = Format(dValorCC / (dValorB + dValorO), "PERCENT")
        CMAPerc.Caption = Format(dValorA / (dValorB + dValorO), "PERCENT")
    End If
    
    Exit Function
    
    Traz_Grids_Tela = SUCESSO

Erro_Traz_Grids_Tela:

    Traz_Grids_Tela = gErr

    Select Case gErr
        
        Case 197222
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197214)

    End Select
    
    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 197224
    
    End If
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 197224

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197225)

    End Select

    Exit Function

End Function

Private Sub BotaoTrazerVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers

On Error GoTo Erro_BotaoTrazerVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVouP.Text)
    objVoucher.sSerie = SerieVouP.Text
    objVoucher.sTipVou = TipoVouP.Text
    
    lErro = Trata_Parametros(objVoucher)
    If lErro <> SUCESSO Then gError 196345

    Exit Sub

Erro_BotaoTrazerVou_Click:

    Select Case gErr
    
        Case 196345
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196346)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHist_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHist_Click

    colSelecao.Add StrParaLong(NumeroVou.Caption)
    colSelecao.Add TipoVou.Caption
    colSelecao.Add SerieVou.Caption

    Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "NumVou= ? AND TipVou = ? AND Serie = ?")

    Exit Sub

Erro_BotaoHist_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAbrirFatRep_Click()

Dim lErro As Long
Dim objVouInfo As ClassTRPVoucherInfo

On Error GoTo Erro_BotaoAbrirFatRep_Click

    If GridRepr.Row = 0 Or GridRepr.Row > objGridRepr.iLinhasExistentes Then gError 196389
    
    Set objVouInfo = gcolCMR.Item(GridRepr.Row)
    
    lErro = Abrir_Doc_Destino(objVouInfo.lNumIntDocDestino, objVouInfo.iTipoDocDestino)
    If lErro <> SUCESSO Then gError 196390
    
    Exit Sub

Erro_BotaoAbrirFatRep_Click:

    Select Case gErr
        
        Case 196389
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 196390
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196391)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirFatCor_Click()

Dim lErro As Long
Dim objVouInfo As ClassTRPVoucherInfo

On Error GoTo Erro_BotaoAbrirFatCor_Click

    If GridCorr.Row = 0 Or GridCorr.Row > objGridCorr.iLinhasExistentes Then gError 196389
    
    Set objVouInfo = gcolCMC.Item(GridCorr.Row)
    
    lErro = Abrir_Doc_Destino(objVouInfo.lNumIntDocDestino, objVouInfo.iTipoDocDestino)
    If lErro <> SUCESSO Then gError 196390
    
    Exit Sub

Erro_BotaoAbrirFatCor_Click:

    Select Case gErr
        
        Case 196389
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 196390
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196391)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirFatEmi_Click()

Dim lErro As Long
Dim objVouInfo As ClassTRPVoucherInfo

On Error GoTo Erro_BotaoAbrirFatEmi_Click

    If GridEmissor.Row = 0 Or GridEmissor.Row > objGridEmissor.iLinhasExistentes Then gError 196389
    
    Set objVouInfo = gcolCME.Item(GridEmissor.Row)
    
    lErro = Abrir_Doc_Destino(objVouInfo.lNumIntDocDestino, objVouInfo.iTipoDocDestino)
    If lErro <> SUCESSO Then gError 196390
    
    Exit Sub

Erro_BotaoAbrirFatEmi_Click:

    Select Case gErr
        
        Case 196389
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 196390
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196391)

    End Select

    Exit Sub
    
End Sub

Private Function Abrir_Doc_Destino(ByVal lNumIntDocDestino As Long, ByVal iTipoDocDestino As Integer)

Dim lErro As Long
Dim objObjeto As Object
Dim sTela As String
Dim bExisteDestino As Boolean
Dim lNumTitulo As Long
Dim sDoc As String

On Error GoTo Erro_Abrir_Doc_Destino

    If lNumIntDocDestino <> 0 Then

        lErro = CF("Verifica_Existencia_Doc_TRP", lNumIntDocDestino, iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
        If lErro <> SUCESSO Then gError 196387
        
        Select Case iTipoDocDestino
        
            Case TRP_TIPO_DOC_DESTINO_CREDFORN
                sTela = TRP_TIPO_DOC_DESTINO_CREDFORN_TELA
                Set objObjeto = New ClassCreditoPagar
                
            Case TRP_TIPO_DOC_DESTINO_DEBCLI
                sTela = TRP_TIPO_DOC_DESTINO_DEBCLI_TELA
                Set objObjeto = New ClassDebitoRecCli
        
            Case TRP_TIPO_DOC_DESTINO_TITPAG
                sTela = TRP_TIPO_DOC_DESTINO_TITPAG_TELA
                Set objObjeto = New ClassTituloPagar
        
            Case TRP_TIPO_DOC_DESTINO_TITREC
                sTela = TRP_TIPO_DOC_DESTINO_TITREC_TELA
                Set objObjeto = New ClassTituloReceber
                
            Case TRP_TIPO_DOC_DESTINO_NFSPAG
                sTela = TRP_TIPO_DOC_DESTINO_NFSPAG_TELA
                Set objObjeto = New ClassNFsPag
        
        End Select
        
        If Not (objObjeto Is Nothing) Then
        
            objObjeto.lNumIntDoc = lNumIntDocDestino
            
            Call Chama_Tela(sTela, objObjeto)
            
        End If
    
    End If
    
    Exit Function

Erro_Abrir_Doc_Destino:

    Select Case gErr
        
        Case 196387
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196388)

    End Select

    Exit Function
    
End Function

Private Sub BotaoComissao_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_BotaoComissao_Click

    objVou.lNumVou = StrParaLong(NumeroVouP.Text)
    objVou.sSerie = SerieVouP.Text
    objVou.sTipVou = TipoVouP.Text
    
    Call Chama_Tela("TRPVouComi", objVou)

    Exit Sub

Erro_BotaoComissao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoVendedores_Click()

'BROWSE VENDEDOR :

Dim colSelecao As New Collection
    
    colSelecao.Add StrParaLong(NumeroVou.Caption)
    colSelecao.Add TipoVou.Caption
    colSelecao.Add SerieVou.Caption
    
    colSelecao.Add TipoVou.Caption
    colSelecao.Add SerieVou.Caption
    colSelecao.Add StrParaLong(NumeroVou.Caption)
   
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, Nothing, Nothing, "Codigo IN (SELECT Promotor FROM TRPVouchers WHERE NumVou = ? AND TipVou = ? AND Serie = ?)  OR Codigo IN (SELECT Vendedor FROM TRPVouVendedores WHERE TipVou = ? AND Serie = ? AND NumVou = ?)")

End Sub

Private Function Inicializa_Grid_Pax(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("T")
    objGridInt.colColuna.Add ("C")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Data Nasc")
    objGridInt.colColuna.Add ("Tipo Doc")
    objGridInt.colColuna.Add ("Núm. Doc")
    objGridInt.colColuna.Add ("Sexo")
    objGridInt.colColuna.Add ("Valor Emi.")
    objGridInt.colColuna.Add ("Valor Atu.")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (PaxTitular.Name)
    objGridInt.colCampo.Add (PaxCancelado.Name)
    objGridInt.colCampo.Add (PaxNome.Name)
    objGridInt.colCampo.Add (PaxDataNasc.Name)
    objGridInt.colCampo.Add (PaxTipoDoc.Name)
    objGridInt.colCampo.Add (PaxNumDoc.Name)
    objGridInt.colCampo.Add (PaxSexo.Name)
    objGridInt.colCampo.Add (PaxValorEmi.Name)
    objGridInt.colCampo.Add (PaxValor.Name)

    'Colunas do GridRepr
    iGrid_PaxTitular_Col = 1
    iGrid_PaxCancelado_Col = 2
    iGrid_PaxNome_Col = 3
    iGrid_PaxDataNasc_Col = 4
    iGrid_PaxTipoDoc_Col = 5
    iGrid_PaxNumDoc_Col = 6
    iGrid_PaxSexo_Col = 7
    iGrid_PaxValorEmi_Col = 8
    iGrid_PaxValor_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = GridPax

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Largura da primeira coluna
    GridPax.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Pax = SUCESSO

    Exit Function

End Function

Public Sub GridPax_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPax, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPax, iAlterado)
    End If
    
End Sub

Public Sub GridPax_EnterCell()
    Call Grid_Entrada_Celula(objGridPax, iAlterado)
End Sub

Public Sub GridPax_GotFocus()
    Call Grid_Recebe_Foco(objGridPax)
End Sub

Public Sub GridPax_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridPax)
    
End Sub

Public Sub GridPax_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPax, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPax, iAlterado)
    End If
    
End Sub

Public Sub GridPax_LeaveCell()
    Call Saida_Celula(objGridPax)
End Sub

Public Sub GridPax_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridPax)
End Sub

Public Sub GridPax_RowColChange()
    Call Grid_RowColChange(objGridPax)
End Sub

Public Sub GridPax_Scroll()
    Call Grid_Scroll(objGridPax)
End Sub

Private Function Inicializa_Grid_Contatos(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Telefone")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ContatoNome.Name)
    objGridInt.colCampo.Add (ContatoTelefone.Name)

    'Colunas do GridRepr
    iGrid_ContatoNome_Col = 1
    iGrid_ContatoTelefone_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridContatos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Largura da primeira coluna
    GridContatos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contatos = SUCESSO

    Exit Function

End Function

Public Sub GridContatos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If
    
End Sub

Public Sub GridContatos_EnterCell()
    Call Grid_Entrada_Celula(objGridContatos, iAlterado)
End Sub

Public Sub GridContatos_GotFocus()
    Call Grid_Recebe_Foco(objGridContatos)
End Sub

Public Sub GridContatos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridContatos)
    
End Sub

Public Sub GridContatos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If
    
End Sub

Public Sub GridContatos_LeaveCell()
    Call Saida_Celula(objGridContatos)
End Sub

Public Sub GridContatos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContatos)
End Sub

Public Sub GridContatos_RowColChange()
    Call Grid_RowColChange(objGridContatos)
End Sub

Public Sub GridContatos_Scroll()
    Call Grid_Scroll(objGridContatos)
End Sub

Private Sub BotaoHistAlt_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistAlt_Click

    colSelecao.Add TipoVou.Caption
    colSelecao.Add SerieVou.Caption
    colSelecao.Add StrParaLong(NumeroVou.Caption)

    Call Chama_Tela("TRPVouHistLista", colSelecao, Nothing, Nothing)

    Exit Sub

Erro_BotaoHistAlt_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub
