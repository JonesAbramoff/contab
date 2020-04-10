VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl TRVVoucher 
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
      TabIndex        =   117
      Top             =   870
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame FrameC 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4710
         Index           =   1
         Left            =   30
         TabIndex        =   119
         Top             =   360
         Width           =   9030
         Begin VB.Frame Frame14 
            Caption         =   "Comissão do Emissor (Over)"
            Height          =   4695
            Left            =   45
            TabIndex        =   202
            Top             =   15
            Width           =   8985
            Begin MSMask.MaskEdBox NumNFEmissor 
               Height          =   225
               Left            =   5865
               TabIndex        =   255
               Top             =   1005
               Width           =   705
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
            Begin VB.Frame Frame18 
               Caption         =   "Resumo"
               Height          =   1425
               Left            =   75
               TabIndex        =   230
               Top             =   2850
               Width           =   8835
               Begin VB.Label ResNC 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   975
                  TabIndex        =   254
                  Top             =   195
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NC:"
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
                  Index           =   103
                  Left            =   585
                  TabIndex        =   253
                  Top             =   240
                  Width           =   390
               End
               Begin VB.Label ResNCData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   3150
                  TabIndex        =   252
                  Top             =   195
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NC:"
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
                  Index           =   102
                  Left            =   2010
                  TabIndex        =   251
                  Top             =   240
                  Width           =   1110
               End
               Begin VB.Label ResHist 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   975
                  TabIndex        =   250
                  Top             =   1110
                  Width           =   7725
               End
               Begin VB.Label Label1 
                  Caption         =   "Histórico:"
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
                  Index           =   101
                  Left            =   105
                  TabIndex        =   249
                  Top             =   1140
                  Width           =   810
               End
               Begin VB.Label ResNCValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   5055
                  TabIndex        =   248
                  Top             =   210
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NC:"
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
                  Index           =   100
                  Left            =   4185
                  TabIndex        =   247
                  Top             =   255
                  Width           =   870
               End
               Begin VB.Label ResNCForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   6885
                  TabIndex        =   246
                  Top             =   210
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NC:"
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
                  Index           =   99
                  Left            =   6045
                  TabIndex        =   245
                  Top             =   255
                  Width           =   840
               End
               Begin VB.Label ResNF 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   975
                  TabIndex        =   244
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NF:"
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
                  Index           =   98
                  Left            =   585
                  TabIndex        =   243
                  Top             =   540
                  Width           =   390
               End
               Begin VB.Label ResNFData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   3150
                  TabIndex        =   242
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NF:"
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
                  Index           =   97
                  Left            =   2010
                  TabIndex        =   241
                  Top             =   540
                  Width           =   1110
               End
               Begin VB.Label ResNFValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   5055
                  TabIndex        =   240
                  Top             =   510
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NF:"
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
                  Index           =   96
                  Left            =   4185
                  TabIndex        =   239
                  Top             =   555
                  Width           =   855
               End
               Begin VB.Label ResNFForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   6885
                  TabIndex        =   238
                  Top             =   510
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NF:"
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
                  Index           =   95
                  Left            =   6045
                  TabIndex        =   237
                  Top             =   555
                  Width           =   840
               End
               Begin VB.Label ResUsu 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   975
                  TabIndex        =   236
                  Top             =   795
                  Width           =   3150
               End
               Begin VB.Label Label1 
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
                  Height          =   285
                  Index           =   94
                  Left            =   195
                  TabIndex        =   235
                  Top             =   840
                  Width           =   720
               End
               Begin VB.Label ResData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   5055
                  TabIndex        =   234
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   93
                  Left            =   4515
                  TabIndex        =   233
                  Top             =   840
                  Width           =   540
               End
               Begin VB.Label ResHora 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   0
                  Left            =   6885
                  TabIndex        =   232
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Hora:"
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
                  Index           =   92
                  Left            =   6360
                  TabIndex        =   231
                  Top             =   840
                  Width           =   525
               End
            End
            Begin VB.CommandButton BotaoAbrirNF 
               Caption         =   "Abrir Nota de Fiscal"
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
               Index           =   0
               Left            =   2190
               TabIndex        =   229
               Top             =   4320
               Width           =   2080
            End
            Begin VB.CommandButton BotaoAbrirNC 
               Caption         =   "Abrir Nota de Crédito"
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
               Index           =   0
               Left            =   45
               TabIndex        =   228
               Top             =   4320
               Width           =   2080
            End
            Begin VB.TextBox HistoricoEmissor 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   6645
               MaxLength       =   250
               TabIndex        =   206
               Top             =   1020
               Width           =   2190
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
               Height          =   270
               Left            =   5385
               TabIndex        =   205
               Top             =   195
               Width           =   660
            End
            Begin MSMask.MaskEdBox Emi 
               Height          =   225
               Left            =   435
               TabIndex        =   203
               Top             =   1020
               Width           =   1995
               _ExtentX        =   3519
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
               Left            =   4290
               TabIndex        =   204
               Top             =   1005
               Width           =   795
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumTitEmissor 
               Height          =   225
               Left            =   5115
               TabIndex        =   207
               Top             =   1005
               Width           =   705
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
               Left            =   2460
               TabIndex        =   208
               Top             =   1020
               Width           =   1005
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
               Left            =   3495
               TabIndex        =   209
               Top             =   990
               Width           =   795
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
               Height          =   510
               Left            =   75
               TabIndex        =   210
               Top             =   510
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   900
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
            Begin VB.Label TotalComiEmi 
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   7215
               TabIndex        =   216
               Top             =   4350
               Width           =   1560
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Saldo:"
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
               Index           =   54
               Left            =   6495
               TabIndex        =   215
               Top             =   4380
               Width           =   630
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
               TabIndex        =   214
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label PercComiEmi 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   7830
               TabIndex        =   213
               Top             =   210
               Width           =   1065
            End
            Begin VB.Label Emissor 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   1785
               TabIndex        =   212
               Top             =   210
               Width           =   3600
            End
            Begin VB.Label Label1 
               Caption         =   "Emissor:"
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
               Left            =   990
               TabIndex        =   211
               Top             =   240
               Width           =   690
            End
         End
      End
      Begin VB.Frame FrameC 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4710
         Index           =   5
         Left            =   30
         TabIndex        =   199
         Top             =   360
         Visible         =   0   'False
         Width           =   9030
         Begin VB.Frame Frame15 
            Caption         =   "Comissão do Promotor"
            Height          =   4695
            Left            =   45
            TabIndex        =   217
            Top             =   15
            Width           =   8985
            Begin VB.CommandButton BotaoAbrirProm 
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
               Height          =   270
               Left            =   5385
               TabIndex        =   219
               Top             =   195
               Width           =   660
            End
            Begin MSMask.MaskEdBox VendedorPromo 
               Height          =   225
               Left            =   1995
               TabIndex        =   218
               Top             =   1170
               Width           =   3090
               _ExtentX        =   5450
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
               Left            =   840
               TabIndex        =   220
               Top             =   1380
               Width           =   1400
               _ExtentX        =   2461
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
               TabIndex        =   221
               Top             =   1200
               Width           =   1800
               _ExtentX        =   3175
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
               TabIndex        =   222
               Top             =   1185
               Width           =   1600
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
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridPromotor 
               Height          =   1215
               Left            =   45
               TabIndex        =   223
               Top             =   510
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
               Caption         =   "Saldo:"
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
               Index           =   56
               Left            =   6630
               TabIndex        =   227
               Top             =   4395
               Width           =   630
            End
            Begin VB.Label TotalComiPro 
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   7335
               TabIndex        =   226
               Top             =   4365
               Width           =   1560
            End
            Begin VB.Label Promotor 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   1785
               TabIndex        =   225
               Top             =   210
               Width           =   3600
            End
            Begin VB.Label Label1 
               Caption         =   "Promotor:"
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
               Index           =   63
               Left            =   885
               TabIndex        =   224
               Top             =   225
               Width           =   825
            End
         End
      End
      Begin VB.Frame FrameC 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4710
         Index           =   4
         Left            =   30
         TabIndex        =   201
         Top             =   360
         Visible         =   0   'False
         Width           =   9030
         Begin VB.Frame Frame13 
            Caption         =   "Comissão do Correntista (CMC)"
            Height          =   4695
            Left            =   45
            TabIndex        =   342
            Top             =   15
            Width           =   8985
            Begin VB.Frame Frame21 
               Caption         =   "Resumo"
               Height          =   1425
               Left            =   75
               TabIndex        =   347
               Top             =   2850
               Width           =   8835
               Begin VB.Label ResNC 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   975
                  TabIndex        =   371
                  Top             =   195
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NC:"
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
                  Index           =   127
                  Left            =   585
                  TabIndex        =   370
                  Top             =   240
                  Width           =   390
               End
               Begin VB.Label ResNCData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   3150
                  TabIndex        =   369
                  Top             =   195
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NC:"
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
                  Index           =   126
                  Left            =   2010
                  TabIndex        =   368
                  Top             =   240
                  Width           =   1110
               End
               Begin VB.Label ResHist 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   975
                  TabIndex        =   367
                  Top             =   1110
                  Width           =   7725
               End
               Begin VB.Label Label1 
                  Caption         =   "Histórico:"
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
                  Index           =   125
                  Left            =   105
                  TabIndex        =   366
                  Top             =   1140
                  Width           =   810
               End
               Begin VB.Label ResNCValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   5055
                  TabIndex        =   365
                  Top             =   210
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NC:"
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
                  Index           =   124
                  Left            =   4185
                  TabIndex        =   364
                  Top             =   255
                  Width           =   870
               End
               Begin VB.Label ResNCForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   6885
                  TabIndex        =   363
                  Top             =   210
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NC:"
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
                  Index           =   123
                  Left            =   6045
                  TabIndex        =   362
                  Top             =   255
                  Width           =   840
               End
               Begin VB.Label ResNF 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   975
                  TabIndex        =   361
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NF:"
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
                  Index           =   122
                  Left            =   585
                  TabIndex        =   360
                  Top             =   540
                  Width           =   390
               End
               Begin VB.Label ResNFData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   3150
                  TabIndex        =   359
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NF:"
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
                  Index           =   121
                  Left            =   2010
                  TabIndex        =   358
                  Top             =   540
                  Width           =   1110
               End
               Begin VB.Label ResNFValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   5055
                  TabIndex        =   357
                  Top             =   510
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NF:"
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
                  Index           =   120
                  Left            =   4185
                  TabIndex        =   356
                  Top             =   555
                  Width           =   855
               End
               Begin VB.Label ResNFForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   6885
                  TabIndex        =   355
                  Top             =   510
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NF:"
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
                  Index           =   119
                  Left            =   6045
                  TabIndex        =   354
                  Top             =   555
                  Width           =   840
               End
               Begin VB.Label ResUsu 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   975
                  TabIndex        =   353
                  Top             =   795
                  Width           =   3150
               End
               Begin VB.Label Label1 
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
                  Height          =   285
                  Index           =   60
                  Left            =   195
                  TabIndex        =   352
                  Top             =   840
                  Width           =   720
               End
               Begin VB.Label ResData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   5055
                  TabIndex        =   351
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   59
                  Left            =   4515
                  TabIndex        =   350
                  Top             =   840
                  Width           =   540
               End
               Begin VB.Label ResHora 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   3
                  Left            =   6885
                  TabIndex        =   349
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Hora:"
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
                  Index           =   55
                  Left            =   6360
                  TabIndex        =   348
                  Top             =   840
                  Width           =   525
               End
            End
            Begin VB.CommandButton BotaoAbrirNF 
               Caption         =   "Abrir Nota de Fiscal"
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
               Index           =   3
               Left            =   2190
               TabIndex        =   346
               Top             =   4320
               Width           =   2080
            End
            Begin VB.CommandButton BotaoAbrirNC 
               Caption         =   "Abrir Nota de Crédito"
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
               Index           =   3
               Left            =   45
               TabIndex        =   345
               Top             =   4320
               Width           =   2080
            End
            Begin VB.TextBox HistoricoCorr 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   6645
               MaxLength       =   250
               TabIndex        =   344
               Top             =   1020
               Width           =   2190
            End
            Begin VB.CommandButton BotaoAbrirCorr 
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
               Height          =   270
               Left            =   5385
               TabIndex        =   343
               Top             =   195
               Width           =   660
            End
            Begin MSMask.MaskEdBox NumNFCorr 
               Height          =   225
               Left            =   5865
               TabIndex        =   372
               Top             =   1005
               Width           =   705
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
            Begin MSMask.MaskEdBox Corr 
               Height          =   225
               Left            =   435
               TabIndex        =   373
               Top             =   1020
               Width           =   1995
               _ExtentX        =   3519
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
            Begin MSMask.MaskEdBox StatusCorr 
               Height          =   225
               Left            =   4290
               TabIndex        =   374
               Top             =   1005
               Width           =   795
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumTitCorr 
               Height          =   225
               Left            =   5115
               TabIndex        =   375
               Top             =   1005
               Width           =   705
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
               Left            =   2460
               TabIndex        =   376
               Top             =   1020
               Width           =   1005
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
               Left            =   3495
               TabIndex        =   377
               Top             =   990
               Width           =   795
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
            Begin MSFlexGridLib.MSFlexGrid GridCorr 
               Height          =   510
               Left            =   75
               TabIndex        =   378
               Top             =   510
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   900
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
            Begin VB.Label TotalComiCor 
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   7215
               TabIndex        =   384
               Top             =   4350
               Width           =   1560
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Saldo:"
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
               Index           =   130
               Left            =   6495
               TabIndex        =   383
               Top             =   4380
               Width           =   630
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
               Index           =   129
               Left            =   6750
               TabIndex        =   382
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label PercComiCor 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   7830
               TabIndex        =   381
               Top             =   210
               Width           =   1065
            End
            Begin VB.Label Correntista 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   1785
               TabIndex        =   380
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
               Index           =   128
               Left            =   615
               TabIndex        =   379
               Top             =   240
               Width           =   1350
            End
         End
      End
      Begin VB.Frame FrameC 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4710
         Index           =   3
         Left            =   30
         TabIndex        =   200
         Top             =   360
         Visible         =   0   'False
         Width           =   9030
         Begin VB.Frame Frame12 
            Caption         =   "Comissão do Representante (CMR)"
            Height          =   4695
            Left            =   45
            TabIndex        =   299
            Top             =   15
            Width           =   8985
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
               Height          =   270
               Left            =   5385
               TabIndex        =   329
               Top             =   195
               Width           =   660
            End
            Begin VB.TextBox HistoricoRepr 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   6645
               MaxLength       =   250
               TabIndex        =   328
               Top             =   1020
               Width           =   2190
            End
            Begin VB.CommandButton BotaoAbrirNC 
               Caption         =   "Abrir Nota de Crédito"
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
               Index           =   2
               Left            =   45
               TabIndex        =   327
               Top             =   4320
               Width           =   2080
            End
            Begin VB.CommandButton BotaoAbrirNF 
               Caption         =   "Abrir Nota de Fiscal"
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
               Index           =   2
               Left            =   2190
               TabIndex        =   326
               Top             =   4320
               Width           =   2080
            End
            Begin VB.Frame Frame16 
               Caption         =   "Resumo"
               Height          =   1425
               Left            =   75
               TabIndex        =   301
               Top             =   2850
               Width           =   8835
               Begin VB.Label Label1 
                  Caption         =   "Hora:"
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
                  Index           =   88
                  Left            =   6360
                  TabIndex        =   325
                  Top             =   840
                  Width           =   525
               End
               Begin VB.Label ResHora 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   6885
                  TabIndex        =   324
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   87
                  Left            =   4515
                  TabIndex        =   323
                  Top             =   840
                  Width           =   540
               End
               Begin VB.Label ResData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   5055
                  TabIndex        =   322
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
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
                  Height          =   285
                  Index           =   86
                  Left            =   195
                  TabIndex        =   321
                  Top             =   840
                  Width           =   720
               End
               Begin VB.Label ResUsu 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   975
                  TabIndex        =   320
                  Top             =   795
                  Width           =   3150
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NF:"
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
                  Index           =   85
                  Left            =   6045
                  TabIndex        =   319
                  Top             =   555
                  Width           =   840
               End
               Begin VB.Label ResNFForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   6885
                  TabIndex        =   318
                  Top             =   510
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NF:"
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
                  Index           =   84
                  Left            =   4185
                  TabIndex        =   317
                  Top             =   555
                  Width           =   855
               End
               Begin VB.Label ResNFValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   5055
                  TabIndex        =   316
                  Top             =   510
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NF:"
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
                  Index           =   83
                  Left            =   2010
                  TabIndex        =   315
                  Top             =   540
                  Width           =   1110
               End
               Begin VB.Label ResNFData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   3150
                  TabIndex        =   314
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NF:"
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
                  Index           =   82
                  Left            =   585
                  TabIndex        =   313
                  Top             =   540
                  Width           =   390
               End
               Begin VB.Label ResNF 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   975
                  TabIndex        =   312
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NC:"
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
                  Index           =   81
                  Left            =   6045
                  TabIndex        =   311
                  Top             =   255
                  Width           =   840
               End
               Begin VB.Label ResNCForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   6885
                  TabIndex        =   310
                  Top             =   210
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NC:"
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
                  Index           =   80
                  Left            =   4185
                  TabIndex        =   309
                  Top             =   255
                  Width           =   870
               End
               Begin VB.Label ResNCValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   5055
                  TabIndex        =   308
                  Top             =   210
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Histórico:"
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
                  Index           =   58
                  Left            =   105
                  TabIndex        =   307
                  Top             =   1140
                  Width           =   810
               End
               Begin VB.Label ResHist 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   975
                  TabIndex        =   306
                  Top             =   1110
                  Width           =   7725
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NC:"
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
                  Index           =   57
                  Left            =   2010
                  TabIndex        =   305
                  Top             =   240
                  Width           =   1110
               End
               Begin VB.Label ResNCData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   3150
                  TabIndex        =   304
                  Top             =   195
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NC:"
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
                  Index           =   53
                  Left            =   585
                  TabIndex        =   303
                  Top             =   240
                  Width           =   390
               End
               Begin VB.Label ResNC 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   2
                  Left            =   975
                  TabIndex        =   302
                  Top             =   195
                  Width           =   975
               End
            End
            Begin MSMask.MaskEdBox NumNFRepr 
               Height          =   225
               Left            =   5865
               TabIndex        =   300
               Top             =   1005
               Width           =   705
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
            Begin MSMask.MaskEdBox Rep 
               Height          =   225
               Left            =   435
               TabIndex        =   330
               Top             =   1020
               Width           =   1995
               _ExtentX        =   3519
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
            Begin MSMask.MaskEdBox StatusRepr 
               Height          =   225
               Left            =   4290
               TabIndex        =   331
               Top             =   1005
               Width           =   795
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumTitRepr 
               Height          =   225
               Left            =   5115
               TabIndex        =   332
               Top             =   1005
               Width           =   705
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
               Left            =   2460
               TabIndex        =   333
               Top             =   1020
               Width           =   1005
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
               Left            =   3495
               TabIndex        =   334
               Top             =   990
               Width           =   795
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
               Height          =   510
               Left            =   75
               TabIndex        =   335
               Top             =   510
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   900
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
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
               Index           =   91
               Left            =   345
               TabIndex        =   341
               Top             =   240
               Width           =   1350
            End
            Begin VB.Label Representante 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   1785
               TabIndex        =   340
               Top             =   210
               Width           =   3600
            End
            Begin VB.Label PercComiRep 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   7830
               TabIndex        =   339
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
               Index           =   90
               Left            =   6750
               TabIndex        =   338
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Saldo:"
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
               Index           =   89
               Left            =   6495
               TabIndex        =   337
               Top             =   4380
               Width           =   630
            End
            Begin VB.Label TotalComiRep 
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   7215
               TabIndex        =   336
               Top             =   4350
               Width           =   1560
            End
         End
      End
      Begin VB.Frame FrameC 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4710
         Index           =   2
         Left            =   30
         TabIndex        =   120
         Top             =   360
         Visible         =   0   'False
         Width           =   9030
         Begin VB.Frame Frame19 
            Caption         =   "Comissão de Cartão da Agência (CMCC)"
            Height          =   4695
            Left            =   45
            TabIndex        =   256
            Top             =   15
            Width           =   8985
            Begin VB.CommandButton BotaoAbrirAG 
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
               Height          =   270
               Left            =   5385
               TabIndex        =   286
               Top             =   195
               Width           =   660
            End
            Begin VB.TextBox HistoricoAg 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   6645
               MaxLength       =   250
               TabIndex        =   285
               Top             =   1020
               Width           =   2190
            End
            Begin VB.CommandButton BotaoAbrirNC 
               Caption         =   "Abrir Nota de Crédito"
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
               Index           =   1
               Left            =   45
               TabIndex        =   284
               Top             =   4320
               Width           =   2080
            End
            Begin VB.CommandButton BotaoAbrirNF 
               Caption         =   "Abrir Nota de Fiscal"
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
               Index           =   1
               Left            =   2190
               TabIndex        =   283
               Top             =   4320
               Width           =   2080
            End
            Begin VB.Frame Frame20 
               Caption         =   "Resumo"
               Height          =   1425
               Left            =   75
               TabIndex        =   258
               Top             =   2850
               Width           =   8835
               Begin VB.Label Label1 
                  Caption         =   "Hora:"
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
                  Index           =   115
                  Left            =   6360
                  TabIndex        =   282
                  Top             =   840
                  Width           =   525
               End
               Begin VB.Label ResHora 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   6885
                  TabIndex        =   281
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   114
                  Left            =   4515
                  TabIndex        =   280
                  Top             =   840
                  Width           =   540
               End
               Begin VB.Label ResData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   5055
                  TabIndex        =   279
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label1 
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
                  Height          =   285
                  Index           =   113
                  Left            =   195
                  TabIndex        =   278
                  Top             =   840
                  Width           =   720
               End
               Begin VB.Label ResUsu 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   975
                  TabIndex        =   277
                  Top             =   795
                  Width           =   3150
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NF:"
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
                  Index           =   112
                  Left            =   6045
                  TabIndex        =   276
                  Top             =   555
                  Width           =   840
               End
               Begin VB.Label ResNFForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   6885
                  TabIndex        =   275
                  Top             =   510
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NF:"
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
                  Index           =   111
                  Left            =   4185
                  TabIndex        =   274
                  Top             =   555
                  Width           =   855
               End
               Begin VB.Label ResNFValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   5055
                  TabIndex        =   273
                  Top             =   510
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NF:"
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
                  Index           =   110
                  Left            =   2010
                  TabIndex        =   272
                  Top             =   540
                  Width           =   1110
               End
               Begin VB.Label ResNFData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   3150
                  TabIndex        =   271
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NF:"
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
                  Index           =   109
                  Left            =   585
                  TabIndex        =   270
                  Top             =   540
                  Width           =   390
               End
               Begin VB.Label ResNF 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   975
                  TabIndex        =   269
                  Top             =   495
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Forn. NC:"
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
                  Index           =   108
                  Left            =   6045
                  TabIndex        =   268
                  Top             =   255
                  Width           =   840
               End
               Begin VB.Label ResNCForn 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   6885
                  TabIndex        =   267
                  Top             =   210
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor NC:"
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
                  Index           =   107
                  Left            =   4185
                  TabIndex        =   266
                  Top             =   255
                  Width           =   870
               End
               Begin VB.Label ResNCValor 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   5055
                  TabIndex        =   265
                  Top             =   210
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Histórico:"
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
                  Index           =   106
                  Left            =   105
                  TabIndex        =   264
                  Top             =   1140
                  Width           =   810
               End
               Begin VB.Label ResHist 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   975
                  TabIndex        =   263
                  Top             =   1110
                  Width           =   7725
               End
               Begin VB.Label Label1 
                  Caption         =   "Emissão NC:"
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
                  Index           =   105
                  Left            =   2010
                  TabIndex        =   262
                  Top             =   240
                  Width           =   1110
               End
               Begin VB.Label ResNCData 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   3150
                  TabIndex        =   261
                  Top             =   195
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "NC:"
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
                  Index           =   104
                  Left            =   585
                  TabIndex        =   260
                  Top             =   240
                  Width           =   390
               End
               Begin VB.Label ResNC 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Index           =   1
                  Left            =   975
                  TabIndex        =   259
                  Top             =   195
                  Width           =   975
               End
            End
            Begin MSMask.MaskEdBox NumNFAg 
               Height          =   225
               Left            =   5865
               TabIndex        =   257
               Top             =   1005
               Width           =   705
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
            Begin MSMask.MaskEdBox Ag 
               Height          =   225
               Left            =   435
               TabIndex        =   287
               Top             =   1020
               Width           =   1995
               _ExtentX        =   3519
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
            Begin MSMask.MaskEdBox StatusAg 
               Height          =   225
               Left            =   4290
               TabIndex        =   288
               Top             =   1005
               Width           =   795
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumTitAg 
               Height          =   225
               Left            =   5115
               TabIndex        =   289
               Top             =   1005
               Width           =   705
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
            Begin MSMask.MaskEdBox DataAg 
               Height          =   225
               Left            =   2460
               TabIndex        =   290
               Top             =   1020
               Width           =   1005
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
            Begin MSMask.MaskEdBox ValorAg 
               Height          =   225
               Left            =   3495
               TabIndex        =   291
               Top             =   990
               Width           =   795
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
            Begin MSFlexGridLib.MSFlexGrid GridAg 
               Height          =   510
               Left            =   75
               TabIndex        =   292
               Top             =   510
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   900
               _Version        =   393216
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
            Begin VB.Label Label1 
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
               Height          =   330
               Index           =   118
               Left            =   990
               TabIndex        =   298
               Top             =   240
               Width           =   690
            End
            Begin VB.Label Agencia 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   1785
               TabIndex        =   297
               Top             =   210
               Width           =   3600
            End
            Begin VB.Label PercComiAg 
               BorderStyle     =   1  'Fixed Single
               Height          =   275
               Left            =   7830
               TabIndex        =   296
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
               Index           =   117
               Left            =   6750
               TabIndex        =   295
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Saldo:"
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
               Index           =   116
               Left            =   6495
               TabIndex        =   294
               Top             =   4380
               Width           =   630
            End
            Begin VB.Label TotalComiAg 
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   7215
               TabIndex        =   293
               Top             =   4350
               Width           =   1560
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   5160
         Left            =   0
         TabIndex        =   118
         Top             =   0
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   9102
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "OVER"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "CMCC"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "CMR"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "CMC"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Promotor"
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
      Left            =   210
      TabIndex        =   22
      Top             =   825
      Width           =   9150
      Begin VB.CommandButton BotaoCTB 
         Caption         =   "Contabilização"
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
         Left            =   7335
         TabIndex        =   15
         ToolTipText     =   "Exibe a contabilização do voucher"
         Top             =   4725
         Width           =   1635
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
         Left            =   6135
         TabIndex        =   14
         ToolTipText     =   "Exibe os vendedores que receberão comissão"
         Top             =   4725
         Width           =   1185
      End
      Begin VB.CommandButton BotaoHistImport 
         Caption         =   "Histórico de Importações"
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
         Left            =   3405
         TabIndex        =   12
         ToolTipText     =   "Exibe o histórico de importação"
         Top             =   4725
         Width           =   1350
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
         Left            =   1050
         TabIndex        =   10
         ToolTipText     =   "Permite alterações no comissionamento"
         Top             =   4725
         Width           =   990
      End
      Begin VB.CommandButton BotaoHist 
         Caption         =   "Detalhamento dos valores"
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
         Left            =   2055
         TabIndex        =   11
         ToolTipText     =   "Exibe as informações detalhadas do voucher"
         Top             =   4725
         Width           =   1350
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
         Left            =   4770
         TabIndex        =   13
         ToolTipText     =   "Exibe o histórico de ocorrências"
         Top             =   4725
         Width           =   1350
      End
      Begin VB.Frame Frame7 
         Caption         =   "Complemento"
         Height          =   1320
         Left            =   30
         TabIndex        =   94
         Top             =   3210
         Width           =   8955
         Begin VB.Frame Frame11 
            Caption         =   "Antc"
            Height          =   510
            Left            =   90
            TabIndex        =   110
            Top             =   765
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
               TabIndex        =   112
               Top             =   210
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
               TabIndex        =   111
               Top             =   210
               Width           =   675
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Destino"
            Height          =   1080
            Left            =   5070
            TabIndex        =   103
            Top             =   195
            Width           =   3825
            Begin VB.Label DestinoVou 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   945
               TabIndex        =   107
               Top             =   615
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
               TabIndex        =   106
               Top             =   660
               Width           =   1245
            End
            Begin VB.Label Destino 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   930
               TabIndex        =   105
               Top             =   195
               Width           =   2835
            End
            Begin VB.Label Label1 
               Caption         =   "Dest:"
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
               Left            =   480
               TabIndex        =   104
               Top             =   240
               Width           =   750
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Vigência"
            Height          =   630
            Left            =   1665
            TabIndex        =   98
            Top             =   195
            Width           =   3360
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
               Left            =   1785
               TabIndex        =   102
               Top             =   240
               Width           =   360
            End
            Begin VB.Label VigenciaAte 
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   2160
               TabIndex        =   101
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
               Left            =   255
               TabIndex        =   100
               Top             =   255
               Width           =   360
            End
            Begin VB.Label VigenciaDe 
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   615
               TabIndex        =   99
               Top             =   195
               Width           =   1110
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Inf. Atualizadas"
            Height          =   570
            Left            =   90
            TabIndex        =   95
            Top             =   195
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
               Height          =   240
               Left            =   90
               TabIndex        =   97
               Top             =   225
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
               Height          =   240
               Left            =   750
               TabIndex        =   96
               Top             =   225
               Width           =   675
            End
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
            Left            =   2925
            TabIndex        =   114
            Top             =   915
            Width           =   900
         End
         Begin VB.Label Convenio 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3825
            TabIndex        =   113
            Top             =   855
            Width           =   1125
         End
         Begin VB.Label Idioma 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2280
            TabIndex        =   109
            Top             =   855
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
            Left            =   1605
            TabIndex        =   108
            Top             =   900
            Width           =   750
         End
      End
      Begin VB.CommandButton BotaoExtrairSigav 
         Caption         =   "Extrair dados do Sigav"
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
         Left            =   7230
         TabIndex        =   16
         ToolTipText     =   "Para extrai os dados do Sigav."
         Top             =   4725
         Visible         =   0   'False
         Width           =   1740
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
         Left            =   45
         TabIndex        =   9
         ToolTipText     =   "Cancela\Reativa o voucher"
         Top             =   4725
         Width           =   990
      End
      Begin VB.Frame Frame3 
         Caption         =   "Detalhes"
         Height          =   1350
         Left            =   30
         TabIndex        =   38
         Top             =   1875
         Width           =   8955
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
            Left            =   3780
            TabIndex        =   8
            Top             =   900
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
            Left            =   6600
            TabIndex        =   7
            Top             =   165
            Width           =   765
         End
         Begin VB.Frame Frame4 
            Caption         =   "Cartão"
            Height          =   1080
            Left            =   8130
            TabIndex        =   39
            Top             =   195
            Width           =   765
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
               Left            =   45
               TabIndex        =   41
               Top             =   720
               Width           =   690
            End
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
               Left            =   45
               TabIndex        =   40
               Top             =   360
               Width           =   690
            End
         End
         Begin VB.Label Label1 
            Caption         =   "U.Web:"
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
            Index           =   70
            Left            =   4515
            TabIndex        =   144
            Top             =   960
            Width           =   675
         End
         Begin VB.Label UsuarioWeb 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5190
            TabIndex        =   143
            Top             =   915
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Pax:"
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
            Left            =   6900
            TabIndex        =   138
            Top             =   960
            Width           =   390
         End
         Begin VB.Label Pax 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7335
            TabIndex        =   137
            Top             =   915
            Width           =   765
         End
         Begin VB.Label Moeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3675
            TabIndex        =   136
            Top             =   555
            Width           =   765
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
            Left            =   3000
            TabIndex        =   135
            Top             =   600
            Width           =   615
         End
         Begin VB.Label TarifaUN 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5190
            TabIndex        =   134
            Top             =   555
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "T. Unit.:"
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
            Left            =   4455
            TabIndex        =   133
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Cambio 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7335
            TabIndex        =   132
            Top             =   525
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Cambio:"
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
            Left            =   6615
            TabIndex        =   131
            Top             =   600
            Width           =   750
         End
         Begin VB.Label NumeroFat 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5190
            TabIndex        =   50
            Top             =   195
            Width           =   1440
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
            Left            =   4575
            TabIndex        =   51
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Produto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1155
            TabIndex        =   49
            Top             =   930
            Width           =   2655
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
            Left            =   345
            TabIndex        =   48
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label CondPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1155
            TabIndex        =   47
            Top             =   570
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Cond.Pagto:"
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
            Left            =   45
            TabIndex        =   46
            Top             =   630
            Width           =   1020
         End
         Begin VB.Label Controle 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1155
            TabIndex        =   45
            Top             =   180
            Width           =   3285
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
            Left            =   300
            TabIndex        =   44
            Top             =   210
            Width           =   1020
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Básicos"
         Height          =   1845
         Left            =   15
         TabIndex        =   25
         Top             =   0
         Width           =   8970
         Begin VB.CommandButton BotaoAbrirCli2 
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
            Left            =   4095
            TabIndex        =   6
            Top             =   1020
            Width           =   660
         End
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
            Left            =   4080
            TabIndex        =   5
            Top             =   615
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
            Left            =   7620
            TabIndex        =   4
            Top             =   210
            Width           =   1260
         End
         Begin VB.Label ValorBase 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7770
            TabIndex        =   198
            Top             =   1440
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Base:"
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
            Index           =   78
            Left            =   7245
            TabIndex        =   197
            Top             =   1500
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "CMA\CMCC:"
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
            Index           =   69
            Left            =   4950
            TabIndex        =   142
            Top             =   1095
            Width           =   1080
         End
         Begin VB.Label ValorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6060
            TabIndex        =   141
            Top             =   1035
            Width           =   1125
         End
         Begin VB.Label ClienteVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1200
            TabIndex        =   140
            Top             =   1035
            Width           =   2910
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente Vou:"
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
            Left            =   90
            TabIndex        =   139
            Top             =   1080
            Width           =   1065
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
            Height          =   315
            Index           =   67
            Left            =   2445
            TabIndex        =   130
            Top             =   1500
            Width           =   525
         End
         Begin VB.Label TarifaMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2985
            TabIndex        =   129
            Top             =   1440
            Width           =   1125
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
            Index           =   66
            Left            =   315
            TabIndex        =   125
            Top             =   1500
            Width           =   870
         End
         Begin VB.Label ValorBruto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   124
            Top             =   1440
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Ocr:"
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
            Index           =   52
            Left            =   7350
            TabIndex        =   116
            Top             =   1095
            Width           =   405
         End
         Begin VB.Label ValorOcr 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7785
            TabIndex        =   115
            Top             =   1035
            Width           =   1125
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   6060
            TabIndex        =   42
            Top             =   615
            Width           =   2835
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
            Left            =   5415
            TabIndex        =   43
            Top             =   660
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
            TabIndex        =   37
            Top             =   240
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
            Index           =   2
            Left            =   1725
            TabIndex        =   36
            Top             =   240
            Width           =   480
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
            Index           =   3
            Left            =   2985
            TabIndex        =   35
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente Fat:"
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
            Left            =   150
            TabIndex        =   34
            Top             =   675
            Width           =   1035
         End
         Begin VB.Label TipoVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1170
            TabIndex        =   32
            Top             =   210
            Width           =   510
         End
         Begin VB.Label SerieVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2280
            TabIndex        =   31
            Top             =   210
            Width           =   510
         End
         Begin VB.Label NumeroVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3735
            TabIndex        =   30
            Top             =   210
            Width           =   1440
         End
         Begin VB.Label ClienteFat 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   29
            Top             =   630
            Width           =   2910
         End
         Begin VB.Label DataEmissaoVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   6060
            TabIndex        =   28
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label ValorVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6060
            TabIndex        =   27
            Top             =   1440
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Faturável R$:"
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
            Index           =   4
            Left            =   4830
            TabIndex        =   26
            Top             =   1500
            Width           =   1275
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
            Left            =   5280
            TabIndex        =   33
            Top             =   270
            Width           =   1620
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5070
      Index           =   3
      Left            =   180
      TabIndex        =   23
      Top             =   990
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame Frame6 
         Caption         =   "Dados do Pagamento com Cartão de crédito"
         Height          =   3690
         Left            =   645
         TabIndex        =   81
         Top             =   510
         Width           =   7950
         Begin VB.Label TitularCPF 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   93
            Top             =   3000
            Width           =   885
         End
         Begin VB.Label NumeroCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1650
            TabIndex        =   92
            Top             =   2955
            Width           =   3435
         End
         Begin VB.Label Label1 
            Caption         =   "Quantidade de Parcelas:"
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
            Left            =   3735
            TabIndex        =   91
            Top             =   2370
            Width           =   2325
         End
         Begin VB.Label NumParcelas 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6060
            TabIndex        =   90
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
            TabIndex        =   89
            Top             =   2400
            Width           =   1185
         End
         Begin VB.Label NumAutorizacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   88
            Top             =   2370
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
            TabIndex        =   87
            Top             =   1785
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label ValidadeCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6060
            TabIndex        =   86
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
            TabIndex        =   85
            Top             =   1815
            Width           =   1425
         End
         Begin VB.Label Administradora 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   84
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
            TabIndex        =   83
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Titular 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1635
            TabIndex        =   82
            Top             =   570
            Width           =   6000
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5325
      Index           =   5
      Left            =   165
      TabIndex        =   145
      Top             =   810
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame Frame17 
         BorderStyle     =   0  'None
         Caption         =   "Resumo do Valores"
         Height          =   5445
         Left            =   390
         TabIndex        =   146
         Top             =   -60
         Width           =   8175
         Begin VB.Label CMCPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   196
            Top             =   3870
            Width           =   1245
         End
         Begin VB.Label CMCMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   195
            Top             =   3870
            Width           =   1245
         End
         Begin VB.Label CMCReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   194
            Top             =   3870
            Width           =   1245
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
            TabIndex        =   193
            Top             =   3900
            Width           =   975
         End
         Begin VB.Line Line8 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   4215
            Y2              =   4215
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
            Height          =   210
            Index           =   71
            Left            =   6720
            TabIndex        =   192
            Top             =   510
            Width           =   1245
         End
         Begin VB.Line Line7 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   3825
            Y2              =   3825
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
            Index           =   37
            Left            =   165
            TabIndex        =   191
            Top             =   4275
            Width           =   2475
         End
         Begin VB.Label LIQCMIReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   190
            Top             =   4260
            Width           =   1245
         End
         Begin VB.Label LIQCMIMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   189
            Top             =   4260
            Width           =   1245
         End
         Begin VB.Line Line6 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   750
            Y2              =   750
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
            Index           =   36
            Left            =   180
            TabIndex        =   188
            Top             =   825
            Width           =   1995
         End
         Begin VB.Label BrutoReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   187
            Top             =   795
            Width           =   1245
         End
         Begin VB.Label BrutoMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   186
            Top             =   795
            Width           =   1245
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   3435
            Y2              =   3435
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
            Index           =   33
            Left            =   180
            TabIndex        =   185
            Top             =   3495
            Width           =   915
         End
         Begin VB.Label CMRReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   184
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Label CMRMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   183
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Label CMRPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   182
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   3045
            Y2              =   3045
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
            Index           =   32
            Left            =   180
            TabIndex        =   181
            Top             =   3120
            Width           =   900
         End
         Begin VB.Label OverReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   180
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label OverMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   179
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label OverPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   178
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   2670
            Y2              =   2670
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
            Index           =   31
            Left            =   180
            TabIndex        =   177
            Top             =   1605
            Width           =   2100
         End
         Begin VB.Label FATReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   176
            Top             =   1575
            Width           =   1245
         End
         Begin VB.Label FATMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   175
            Top             =   1575
            Width           =   1245
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   1530
            Y2              =   1530
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
            Index           =   30
            Left            =   180
            TabIndex        =   174
            Top             =   2745
            Width           =   900
         End
         Begin VB.Label CMCCReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   173
            Top             =   2715
            Width           =   1245
         End
         Begin VB.Label CMCCMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   172
            Top             =   2715
            Width           =   1245
         End
         Begin VB.Label CMCCPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   171
            Top             =   2715
            Width           =   1245
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   1140
            Y2              =   1140
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
            Index           =   29
            Left            =   180
            TabIndex        =   170
            Top             =   1215
            Width           =   1680
         End
         Begin VB.Label CMAReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   169
            Top             =   1185
            Width           =   1245
         End
         Begin VB.Label CMAMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   168
            Top             =   1185
            Width           =   1245
         End
         Begin VB.Label CMAPerc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   167
            Top             =   1185
            Width           =   1245
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
            TabIndex        =   166
            Top             =   510
            Width           =   990
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
            TabIndex        =   165
            Top             =   510
            Width           =   1185
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
            Index           =   28
            Left            =   2040
            TabIndex        =   164
            Top             =   180
            Width           =   915
         End
         Begin VB.Label TarifaUNRV 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   163
            Top             =   120
            Width           =   930
         End
         Begin VB.Label CMIMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   162
            Top             =   4635
            Width           =   1245
         End
         Begin VB.Label CMIReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   161
            Top             =   4635
            Width           =   1245
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
            TabIndex        =   160
            Top             =   4665
            Width           =   975
         End
         Begin VB.Line Line9 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Line10 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   4980
            Y2              =   4980
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
            TabIndex        =   159
            Top             =   5040
            Width           =   2205
         End
         Begin VB.Label LiqFinalReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   158
            Top             =   5025
            Width           =   1245
         End
         Begin VB.Label LiqFinalMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   157
            Top             =   5025
            Width           =   1245
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
            TabIndex        =   156
            Top             =   180
            Width           =   630
         End
         Begin VB.Label MoedaRV 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   885
            TabIndex        =   155
            Top             =   135
            Width           =   900
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
            TabIndex        =   154
            Top             =   165
            Width           =   705
         End
         Begin VB.Label CambioRV 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6600
            TabIndex        =   153
            Top             =   120
            Width           =   885
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
            TabIndex        =   152
            Top             =   1995
            Width           =   2085
         End
         Begin VB.Label OCRReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   151
            Top             =   1965
            Width           =   1245
         End
         Begin VB.Label OCRMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   150
            Top             =   1965
            Width           =   1245
         End
         Begin VB.Line Line11 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line12 
            BorderWidth     =   2
            X1              =   90
            X2              =   8000
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Label OCRCMAMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4815
            TabIndex        =   149
            Top             =   2340
            Width           =   1245
         End
         Begin VB.Label OCRCMAReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6615
            TabIndex        =   148
            Top             =   2340
            Width           =   1245
         End
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
            TabIndex        =   147
            Top             =   2370
            Width           =   2085
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   2
      Left            =   225
      TabIndex        =   24
      Top             =   855
      Visible         =   0   'False
      Width           =   9090
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
         Left            =   8340
         TabIndex        =   19
         ToolTipText     =   "Abre a tela de cliente com o passageiro"
         Top             =   270
         Width           =   660
      End
      Begin VB.Frame Frame5 
         Caption         =   "Endereço"
         Height          =   2985
         Left            =   0
         TabIndex        =   56
         Top             =   1860
         Width           =   9060
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
            Left            =   6255
            TabIndex        =   80
            Top             =   1440
            Width           =   525
         End
         Begin VB.Label Telefone2 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6795
            TabIndex        =   79
            Top             =   1380
            Width           =   2175
         End
         Begin VB.Label Telefone1 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4035
            TabIndex        =   77
            Top             =   1365
            Width           =   2115
         End
         Begin VB.Label Contato 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   76
            Top             =   2400
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
            TabIndex        =   75
            Top             =   2445
            Width           =   945
         End
         Begin VB.Label Email 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   74
            Top             =   1920
            Width           =   7785
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
            Height          =   330
            Index           =   21
            Left            =   615
            TabIndex        =   73
            Top             =   1965
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
            Left            =   6390
            TabIndex        =   72
            Top             =   990
            Width           =   330
         End
         Begin VB.Label UF 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6795
            TabIndex        =   71
            Top             =   915
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
            TabIndex        =   70
            Top             =   1485
            Width           =   495
         End
         Begin VB.Label CEP 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   69
            Top             =   1440
            Width           =   2085
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
            TabIndex        =   68
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Cidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4035
            TabIndex        =   67
            Top             =   915
            Width           =   2115
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
            TabIndex        =   66
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Bairro 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   65
            Top             =   915
            Width           =   2085
         End
         Begin VB.Label Endereco 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1185
            TabIndex        =   64
            Top             =   390
            Width           =   7785
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
            TabIndex        =   63
            Top             =   435
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
            TabIndex        =   78
            Top             =   1455
            Width           =   540
         End
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
         Left            =   135
         TabIndex        =   62
         Top             =   1410
         Width           =   990
      End
      Begin VB.Label CartaoFid 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1155
         TabIndex        =   61
         Top             =   1365
         Width           =   2145
      End
      Begin VB.Label DataNasc 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   4710
         TabIndex        =   60
         Top             =   810
         Width           =   1155
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
         Left            =   3585
         TabIndex        =   59
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label CGC 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1155
         TabIndex        =   58
         Top             =   825
         Width           =   2145
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Index           =   13
         Left            =   630
         TabIndex        =   57
         Top             =   870
         Width           =   540
      End
      Begin VB.Label CliPassageiro 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   6870
         TabIndex        =   54
         Top             =   285
         Width           =   1470
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
         Left            =   6090
         TabIndex        =   55
         Top             =   330
         Width           =   1050
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
         Left            =   75
         TabIndex        =   53
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Passageiro 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1155
         TabIndex        =   52
         Top             =   270
         Width           =   4590
      End
   End
   Begin VB.CommandButton BotaoTrazerVou 
      Height          =   315
      Left            =   4155
      Picture         =   "TRVVouchers.ctx":0000
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   30
      Width           =   1080
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   60
         Picture         =   "TRVVouchers.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   540
         Picture         =   "TRVVouchers.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5685
      Left            =   120
      TabIndex        =   21
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
            Caption         =   "Passageiro"
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
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   315
      Left            =   4770
      TabIndex        =   128
      Top             =   60
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
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
      TabIndex        =   121
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
      TabIndex        =   123
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
      TabIndex        =   122
      Top             =   165
      Width           =   750
   End
End
Attribute VB_Name = "TRVVoucher"
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

Dim objGridAg As AdmGrid
Dim objGridRepr As AdmGrid
Dim objGridCorr As AdmGrid
Dim objGridEmissor As AdmGrid
Dim objGridPromotor As AdmGrid

Public gobjVoucher As ClassTRVVouchers
Public gcolCMR As New Collection
Public gcolCMC As New Collection
Public gcolCME As New Collection
Public gcolCMP As New Collection
Public gcolCMCC As New Collection

'Colunas do GridRepr
Dim iGrid_DataRepr_Col As Integer
Dim iGrid_Rep_Col As Integer
Dim iGrid_ValorRepr_Col As Integer
Dim iGrid_NumTitRepr_Col As Integer
Dim iGrid_NumNFRepr_Col As Integer
Dim iGrid_StatusRepr_Col As Integer
Dim iGrid_HistoricoRepr_Col As Integer

Dim iGrid_Corr_Col As Integer
Dim iGrid_DataCorr_Col As Integer
Dim iGrid_ValorCorr_Col As Integer
Dim iGrid_NumTitCorr_Col As Integer
Dim iGrid_NumNFCorr_Col As Integer
Dim iGrid_StatusCorr_Col As Integer
Dim iGrid_HistoricoCorr_Col As Integer

Dim iGrid_Emi_Col As Integer
Dim iGrid_DataEmissor_Col As Integer
Dim iGrid_ValorEmissor_Col As Integer
Dim iGrid_NumTitEmissor_Col As Integer
Dim iGrid_NumNFEmissor_Col As Integer
Dim iGrid_StatusEmissor_Col As Integer
Dim iGrid_HistoricoEmissor_Col As Integer

Dim iGrid_Ag_Col As Integer
Dim iGrid_DataAg_Col As Integer
Dim iGrid_ValorAg_Col As Integer
Dim iGrid_NumTitAg_Col As Integer
Dim iGrid_NumNFAg_Col As Integer
Dim iGrid_StatusAg_Col As Integer
Dim iGrid_HistoricoAg_Col As Integer

Dim iGrid_VendedorPromo_Col As Integer
Dim iGrid_DataPromo_Col As Integer
Dim iGrid_ValorPromoBase_Col As Integer
Dim iGrid_ValorPromoComiss_Col As Integer

Private WithEvents objEventoVoucher As AdmEvento
Attribute objEventoVoucher.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Vouchers"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVVoucher"

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
    Set objGridAg = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
    
    Set gobjVoucher = Nothing
    Set gcolCMR = Nothing
    Set gcolCMC = Nothing
    Set gcolCME = Nothing
    Set gcolCMP = Nothing
    Set gcolCMCC = Nothing

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
    Set objGridAg = New AdmGrid

    lErro = Inicializa_Grid_Representante(objGridRepr)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_Correntista(objGridCorr)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_Emissor(objGridEmissor)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_Promotor(objGridPromotor)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_AG(objGridAg)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0
    
    iFrameAtual = 1
    iFrameAtualC = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190864)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTRVVouchers As ClassTRVVouchers) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTRVVouchers Is Nothing) Then
    
        If Len(Trim(objTRVVouchers.sTipoDoc)) = 0 Then objTRVVouchers.sTipoDoc = TRV_TIPODOC_VOU_TEXTO

        lErro = Traz_TRVVouchers_Tela(objTRVVouchers)
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

Function Move_Tela_Memoria(objTRVVouchers As ClassTRVVouchers, Optional bValida As Boolean = True) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    If bValida Then
    
        If Len(Trim(NumeroVou.Caption)) = 0 Then gError 190889
        If Len(Trim(SerieVou.Caption)) = 0 Then gError 190891
        If Len(Trim(TipoVou.Caption)) = 0 Then gError 190892
        
    End If

    objTRVVouchers.lNumVou = StrParaLong(NumeroVou.Caption)
    objTRVVouchers.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objTRVVouchers.sTipVou = TipoVou.Caption
    objTRVVouchers.sSerie = SerieVou.Caption
    objTRVVouchers.dtData = StrParaDate(DataEmissaoVou.Caption)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 190889 To 190892 'ERRO_TRV_NUMVOU_NAO_PREENCHIDO
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_NUMVOU_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190867)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTRVVouchers As New ClassTRVVouchers

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRVVouchers"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTRVVouchers, False)
    If lErro <> SUCESSO Then gError 190868

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumVou", objTRVVouchers.lNumVou, 0, "NumVou"
    colCampoValor.Add "TipoDoc", objTRVVouchers.sTipoDoc, STRING_TRV_OCR_TIPODOC, "TipoDoc"
    colCampoValor.Add "Serie", objTRVVouchers.sSerie, STRING_TRV_OCR_SERIE, "Serie"
    colCampoValor.Add "TipVou", objTRVVouchers.sTipVou, STRING_TRV_OCR_TIPOVOU, "TipVou"

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
Dim objTRVVouchers As New ClassTRVVouchers

On Error GoTo Erro_Tela_Preenche

    objTRVVouchers.lNumVou = colCampoValor.Item("NumVou").vValor
    objTRVVouchers.sTipoDoc = colCampoValor.Item("TipoDoc").vValor
    objTRVVouchers.sSerie = colCampoValor.Item("Serie").vValor
    objTRVVouchers.sTipVou = colCampoValor.Item("TipVou").vValor

    If objTRVVouchers.lNumVou <> 0 And Len(Trim(objTRVVouchers.sTipoDoc)) > 0 And Len(Trim(objTRVVouchers.sSerie)) > 0 And Len(Trim(objTRVVouchers.sTipVou)) > 0 Then
        
        lErro = Traz_TRVVouchers_Tela(objTRVVouchers)
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

Function Limpa_Tela_TRVVouchers() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVVouchers

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    TipoVou.Caption = ""
    SerieVou.Caption = ""
    NumeroVou.Caption = ""
    ClienteVou.Caption = ""
    ClienteFat.Caption = ""
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
    ValorBruto.Caption = ""
    UsuarioWeb.Caption = ""
    
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
    
    Moeda.Caption = ""
    TarifaUN.Caption = ""
    Cambio.Caption = ""
    
    TarifaUNRV.Caption = ""
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
    Agencia.Caption = ""
    
    TotalComiPro.Caption = ""
    TotalComiEmi.Caption = ""
    TotalComiRep.Caption = ""
    TotalComiCor.Caption = ""
    TotalComiAg.Caption = ""
    
    PercComiAg.Caption = ""
    PercComiEmi.Caption = ""
    PercComiRep.Caption = ""
    PercComiCor.Caption = ""
    
    Set gobjVoucher = Nothing
    Set gcolCMR = New Collection
    Set gcolCMC = New Collection
    Set gcolCME = New Collection
    Set gcolCMP = New Collection
    Set gcolCMCC = New Collection
    
    Call Grid_Limpa(objGridRepr)
    Call Grid_Limpa(objGridPromotor)
    Call Grid_Limpa(objGridEmissor)
    Call Grid_Limpa(objGridCorr)
    Call Grid_Limpa(objGridAg)
    
    Call Limpa_Resumo(0)
    Call Limpa_Resumo(1)
    Call Limpa_Resumo(2)
    Call Limpa_Resumo(3)

    iAlterado = 0

    Limpa_Tela_TRVVouchers = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRVVouchers:

    Limpa_Tela_TRVVouchers = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190881)

    End Select

    Exit Function

End Function

Function Traz_TRVVouchers_Tela(objTRVVouchers As ClassTRVVouchers) As Long

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objTRVVoucherInfo As New ClassTRVVoucherInfo
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Traz_TRVVouchers_Tela

    objTRVVouchers.sTipoDoc = TRV_TIPODOC_VOU_TEXTO

    NumeroVouP.PromptInclude = False
    NumeroVouP.Text = CStr(objTRVVouchers.lNumVou)
    NumeroVouP.PromptInclude = True
    SerieVouP.Text = objTRVVouchers.sSerie
    TipoVouP.Text = objTRVVouchers.sTipVou
    
    'Lê o TRVVouchers que está sendo Passado
    lErro = CF("TRVVouchers_Le", objTRVVouchers)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192646

    If lErro = SUCESSO Then
    
        objTRVVoucherInfo.sSerie = objTRVVouchers.sSerie
        objTRVVoucherInfo.sTipo = objTRVVouchers.sTipVou
        objTRVVoucherInfo.lNumVou = objTRVVouchers.lNumVou

        'Lê o TRVVouchers que está sendo Passado
        lErro = CF("TRVVoucherInfoSigav_Le", objTRVVoucherInfo)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192647
        
        If lErro = SUCESSO Then
        
            OptSigSim.Value = True
        
            lErro = Traz_TRVVoucherInfo_Tela(objTRVVoucherInfo)
            If lErro <> SUCESSO Then gError 192648
            
        Else
            
            OptSigNao.Value = True
            
        End If
        
        TipoVou.Caption = objTRVVouchers.sTipVou
        SerieVou.Caption = objTRVVouchers.sSerie
        NumeroVou.Caption = objTRVVouchers.lNumVou
        
        If objTRVVouchers.iPax <> 0 Then
            Pax.Caption = objTRVVouchers.iPax
        Else
            Pax.Caption = ""
        End If
        
        objcliente.lCodigo = objTRVVouchers.lClienteVou
       
        lErro = CF("Cliente_Le", objcliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 192649
        
        ClienteVou.Caption = objTRVVouchers.lClienteVou & SEPARADOR & objcliente.sNomeReduzido
               
        Set objcliente = New ClassCliente
               
        objcliente.lCodigo = objTRVVouchers.lCliente
       
        lErro = CF("Cliente_Le", objcliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 192649
        
        ClienteFat.Caption = objTRVVouchers.lCliente & SEPARADOR & objcliente.sNomeReduzido
               
        Select Case objTRVVouchers.iStatus
        
            Case STATUS_TRV_VOU_ABERTO
                Status.Caption = STATUS_TRV_VOU_ABERTO_TEXTO
                BotaoCancelar.Caption = "Cancelar"
            
            Case STATUS_TRV_VOU_CANCELADO
                Status.Caption = STATUS_TRV_VOU_CANCELADO_TEXTO
                BotaoCancelar.Caption = "Reativar"
        
        End Select
        
        ValorVou.Caption = Format(objTRVVouchers.dValor, "STANDARD")
        ValorComissao.Caption = Format(objTRVVouchers.dValorComissao, "STANDARD")
        ValorOcr.Caption = Format(objTRVVouchers.dValorOcr, "STANDARD")
        ValorBase.Caption = Format(objTRVVouchers.dValorBaseComis, "STANDARD")
        TarifaMoeda.Caption = Format(objTRVVouchers.dValorCambio, "STANDARD")
        DataEmissaoVou.Caption = Format(objTRVVouchers.dtData, "dd/mm/yyyy")
        Controle.Caption = objTRVVouchers.sControle
        UsuarioWeb.Caption = objTRVVouchers.sUsuarioWeb
        
        objCondicaoPagto.iCodigo = objcliente.iCondicaoPagto

        'Lê Condição Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 192650
        
        CondPagto.Caption = objTRVVouchers.iCondPagto & SEPARADOR & objCondicaoPagto.sDescReduzida
        
        If objTRVVouchers.lNumFatCoinfo <> 0 Then
            NumeroFat.Caption = objTRVVouchers.lNumFatCoinfo
        Else
            NumeroFat.Caption = ""
        End If
        
        lErro = CF("Produto_Formata", objTRVVouchers.sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 192650
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 192650
        
        Produto.Caption = objTRVVouchers.sProduto & SEPARADOR & objProduto.sDescricao
        
        If objTRVVouchers.iCartao = MARCADO Then
            OptSim.Value = True
        Else
            OptNao.Value = True
        End If
        
        Titular.Caption = objTRVVouchers.sTitular
        Administradora.Caption = objTRVVouchers.sCiaCart
        'ValidadeCC.Caption = objTRVVouchers.sValidade
        
'        If objTRVVouchers.lNumAuto > 0 Then
'            NumAutorizacao.Caption = objTRVVouchers.lNumAuto
'        Else
            NumAutorizacao.Caption = objTRVVouchers.sNumAuto
'        End If
        
        If objTRVVouchers.iQuantParc > 0 Then
            NumParcelas.Caption = objTRVVouchers.iQuantParc
        Else
            NumParcelas.Caption = ""
        End If
        
        If Len(Trim(objTRVVouchers.sNumCCred)) > 6 Then
            NumeroCC.Caption = String(Len(Trim(objTRVVouchers.sNumCCred)) - 6, "X") & right(Trim(objTRVVouchers.sNumCCred), 6)
        Else
            NumeroCC.Caption = ""
        End If
        
        'Lê o TRVVouchers que está sendo Passado
        lErro = CF("TRVVoucherInfo_Le", objTRVVouchers, 0, True)
        If lErro <> SUCESSO Then gError 197212
        
        If objTRVVouchers.lCorrentista <> 0 Then
        
            Set objcliente = New ClassCliente
             
            objcliente.lCodigo = objTRVVouchers.lCorrentista
            
            lErro = CF("Cliente_Le", objcliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 192649
        
            Correntista.Caption = objcliente.lCodigo & SEPARADOR & objcliente.sNomeReduzido
            PercComiCor.Caption = Format(objTRVVouchers.dComissaoCorr, "PERCENT")
            
        End If
        
        If objTRVVouchers.lRepresentante <> 0 Then
        
            Set objcliente = New ClassCliente
             
            objcliente.lCodigo = objTRVVouchers.lRepresentante
            
            lErro = CF("Cliente_Le", objcliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 192649
        
            Representante.Caption = objcliente.lCodigo & SEPARADOR & objcliente.sNomeReduzido
            PercComiRep.Caption = Format(objTRVVouchers.dComissaoRep, "PERCENT")
            
        End If
        
        If objTRVVouchers.lEmissor <> 0 Then
                     
            objFornecedor.lCodigo = objTRVVouchers.lEmissor
            
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 192649
        
            Emissor.Caption = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
            PercComiEmi.Caption = Format(objTRVVouchers.dComissaoEmissor, "PERCENT")
            
        End If
        
        Agencia.Caption = ClienteVou.Caption
        If objTRVVouchers.iCartao = MARCADO Then
            PercComiAg.Caption = Format(objTRVVouchers.dComissaoEmissor, "PERCENT")
        Else
            PercComiAg.Caption = ""
        End If
        
        'Busca o Cliente no BD
        If objTRVVouchers.lPromotor <> 0 Then
            Cliente.Text = objTRVVouchers.lPromotor
            lErro = CF("TP_Vendedor_LeTRV", Cliente, objVendedor)
            If lErro <> SUCESSO Then gError 192649
            Promotor.Caption = Cliente.Text
        End If
        
        If objTRVVouchers.dValorBruto <> 0 Then
            ValorBruto.Caption = Format(objTRVVouchers.dValorBruto, "STANDARD")
        End If
        
        lErro = Traz_Grids_Tela(objTRVVouchers)
        If lErro <> SUCESSO Then gError 197213
        
        Select Case Len(Trim(objTRVVouchers.sTitularCPF))
    
            Case STRING_CPF 'CPF
                TitularCPF.Caption = Format(objTRVVouchers.sTitularCPF, "000\.000\.000-00; ; ; ")
    
            Case STRING_CGC 'CGC
                TitularCPF.Caption = Format(objTRVVouchers.sTitularCPF, "00\.000\.000\/0000-00; ; ; ")
                
            Case Else
                TitularCPF.Caption = objTRVVouchers.sTitularCPF
            
        End Select
    
        
    End If
    
    Set gobjVoucher = objTRVVouchers

    iAlterado = 0

    Traz_TRVVouchers_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVVouchers_Tela:

    Traz_TRVVouchers_Tela = gErr

    Select Case gErr

        Case 192646 To 192650, 197212, 197213

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190883)

    End Select

    Exit Function

End Function

Function Traz_TRVVoucherInfo_Tela(objTRVVoucherInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim dTAIValorMoeda As Double
Dim dTAIValorReal As Double
Dim dAGTValorMoeda As Double
Dim dAGTValorReal As Double
Dim objForn As New ClassFornecedor

On Error GoTo Erro_Traz_TRVVoucherInfo_Tela

    Passageiro.Caption = objTRVVoucherInfo.sPassageiroNome & " " & objTRVVoucherInfo.sPassageiroSobreNome
    
    If objTRVVoucherInfo.lCliPassageiro <> 0 Then
        CliPassageiro.Caption = CStr(objTRVVoucherInfo.lCliPassageiro)
    Else
        CliPassageiro.Caption = ""
    End If
    
    CGC.Caption = objTRVVoucherInfo.sPassageiroCGC
    
    If objTRVVoucherInfo.dtDataNasc <> DATA_NULA Then
        DataNasc.Caption = Format(objTRVVoucherInfo.dtDataNasc, "dd/mm/yyyy")
    Else
        DataNasc.Caption = ""
    End If
    
    CartaoFid.Caption = objTRVVoucherInfo.sCartaoFid
    Endereco.Caption = objTRVVoucherInfo.sPassageiroEndereco
    Bairro.Caption = objTRVVoucherInfo.sPassageiroBairro
    Cidade.Caption = objTRVVoucherInfo.sPassageiroCidade
    UF.Caption = objTRVVoucherInfo.sPassageiroUF
    CEP.Caption = objTRVVoucherInfo.sPassageiroCEP
    Telefone1.Caption = objTRVVoucherInfo.sPassageiroTelefone1
    Telefone2.Caption = objTRVVoucherInfo.sPassageiroTelefone2
    Email.Caption = objTRVVoucherInfo.sPassageiroEmail
    Contato.Caption = objTRVVoucherInfo.sPassageiroContato
    
    Moeda.Caption = objTRVVoucherInfo.sMoeda
    
    If objTRVVoucherInfo.dTarifaUnitaria <> 0 Then
        TarifaUN.Caption = Format(objTRVVoucherInfo.dTarifaUnitaria, "STANDARD")
    Else
        TarifaUN.Caption = ""
    End If
    
    If objTRVVoucherInfo.dCambio <> 0 Then
        Cambio.Caption = Format(objTRVVoucherInfo.dCambio, "STANDARD")
    Else
        Cambio.Caption = ""
    End If
    
'    dTAIValorMoeda = objTRVVoucherInfo.dTarifaValorMoeda - objTRVVoucherInfo.dComissaoValorMoeda - objTRVVoucherInfo.dCartaoValorMoeda - objTRVVoucherInfo.dOverValorMoeda - objTRVVoucherInfo.dCMRValorMoeda
'    dTAIValorReal = objTRVVoucherInfo.dTarifaValorReal - objTRVVoucherInfo.dComissaoValorReal - objTRVVoucherInfo.dCartaoValorReal - objTRVVoucherInfo.dOverValorReal - objTRVVoucherInfo.dCMRValorReal
'    dAGTValorMoeda = objTRVVoucherInfo.dTarifaValorMoeda - objTRVVoucherInfo.dComissaoValorMoeda - objTRVVoucherInfo.dCartaoValorMoeda
'    dAGTValorReal = objTRVVoucherInfo.dTarifaValorReal - objTRVVoucherInfo.dComissaoValorReal - objTRVVoucherInfo.dCartaoValorReal
'
'    If objTRVVoucherInfo.dTarifaPerc <> 0 Then
'        TarifaPerc.Caption = Format(objTRVVoucherInfo.dTarifaPerc, "PERCENT")
'    Else
'        TarifaPerc.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dTarifaValorMoeda <> 0 Then
'        TarifaMoeda.Caption = Format(objTRVVoucherInfo.dTarifaValorMoeda, "STANDARD")
'    Else
'        TarifaMoeda.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dTarifaValorReal <> 0 Then
'        TarifaReal.Caption = Format(objTRVVoucherInfo.dTarifaValorReal, "STANDARD")
'    Else
'        TarifaReal.Caption = ""
'    End If

'    If objTRVVoucherInfo.dComissaoPerc <> 0 Then
'        ComissaoPerc.Caption = Format(objTRVVoucherInfo.dComissaoPerc, "PERCENT")
'    Else
'        ComissaoPerc.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dComissaoValorMoeda <> 0 Then
'        ComissaoMoeda.Caption = Format(objTRVVoucherInfo.dComissaoValorMoeda, "STANDARD")
'    Else
'        ComissaoMoeda.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dComissaoValorReal <> 0 Then
'        ComissaoReal.Caption = Format(objTRVVoucherInfo.dComissaoValorReal, "STANDARD")
'    Else
'        ComissaoReal.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dCartaoPerc <> 0 Then
'        CartaoPerc.Caption = Format(objTRVVoucherInfo.dCartaoPerc, "PERCENT")
'    Else
'        CartaoPerc.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dCartaoValorMoeda <> 0 Then
'        CartaoMoeda.Caption = Format(objTRVVoucherInfo.dCartaoValorMoeda, "STANDARD")
'    Else
'        CartaoMoeda.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dCartaoValorReal <> 0 Then
'        CartaoReal.Caption = Format(objTRVVoucherInfo.dCartaoValorReal, "STANDARD")
'    Else
'        CartaoReal.Caption = ""
'    End If
'
'    If dAGTValorMoeda <> 0 Then
'        AGTMoeda.Caption = Format(dAGTValorMoeda, "STANDARD")
'    Else
'        AGTMoeda.Caption = ""
'    End If
'
'    If dAGTValorReal <> 0 Then
'        AGTReal.Caption = Format(dAGTValorReal, "STANDARD")
'    Else
'        AGTReal.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dOverPerc <> 0 Then
'        OverPerc.Caption = Format(objTRVVoucherInfo.dOverPerc, "PERCENT")
'    Else
'        OverPerc.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dOverValorMoeda <> 0 Then
'        OverMoeda.Caption = Format(objTRVVoucherInfo.dOverValorMoeda, "STANDARD")
'    Else
'        OverMoeda.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dOverValorReal <> 0 Then
'        OverReal.Caption = Format(objTRVVoucherInfo.dOverValorReal, "STANDARD")
'    Else
'        OverReal.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dCMRPerc <> 0 Then
'        CMRPerc.Caption = Format(objTRVVoucherInfo.dCMRPerc, "PERCENT")
'    Else
'        CMRPerc.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dCMRValorMoeda <> 0 Then
'        CMRMoeda.Caption = Format(objTRVVoucherInfo.dCMRValorMoeda, "STANDARD")
'    Else
'        CMRMoeda.Caption = ""
'    End If
'
'    If objTRVVoucherInfo.dCMRValorReal <> 0 Then
'        CMRReal.Caption = Format(objTRVVoucherInfo.dCMRValorReal, "STANDARD")
'    Else
'        CMRReal.Caption = ""
'    End If
'
'    If dTAIValorMoeda <> 0 Then
'        TAIMoeda.Caption = Format(dTAIValorMoeda, "STANDARD")
'    Else
'        TAIMoeda.Caption = ""
'    End If
'
'    If dTAIValorReal <> 0 Then
'        TAIReal.Caption = Format(dTAIValorReal, "STANDARD")
'    Else
'        TAIReal.Caption = ""
'    End If
'
'    Titular.Caption = objTRVVoucherInfo.sTitular
'    Administradora.Caption = objTRVVoucherInfo.sCia
'    ValidadeCC.Caption = objTRVVoucherInfo.sValidade
'    NumAutorizacao.Caption = objTRVVoucherInfo.lAprovacao
'    NumParcelas.Caption = objTRVVoucherInfo.lParcela
    
    If Len(Trim(objTRVVoucherInfo.sNumeroCC)) > 6 Then
        NumeroCC.Caption = String(Len(Trim(objTRVVoucherInfo.sNumeroCC)) - 6, "X") & right(Trim(objTRVVoucherInfo.sNumeroCC), 6)
    Else
        NumeroCC.Caption = ""
    End If
    
    If objTRVVoucherInfo.dtDataInicio <> DATA_NULA Then
        VigenciaDe.Caption = Format(objTRVVoucherInfo.dtDataInicio, "dd/mm/yyyy")
    Else
        VigenciaDe.Caption = ""
    End If
    
    If objTRVVoucherInfo.dtDataTermino <> DATA_NULA Then
        VigenciaAte.Caption = Format(objTRVVoucherInfo.dtDataTermino, "dd/mm/yyyy")
    Else
        VigenciaAte.Caption = ""
    End If
    
    If objTRVVoucherInfo.lFornEmissor <> 0 Then
    
        objForn.lCodigo = objTRVVoucherInfo.lFornEmissor
        
        lErro = CF("Fornecedor_Le", objForn)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 190882
    
        Emissor.Caption = objTRVVoucherInfo.lFornEmissor & SEPARADOR & objForn.sNomeReduzido
    Else
        Emissor.Caption = ""
    End If
    
    If objTRVVoucherInfo.iAntc = MARCADO Then
        OptAntcSim.Value = True
    Else
        OptAntcNao.Value = True
    End If
    
    Destino.Caption = objTRVVoucherInfo.sDestino
    DestinoVou.Caption = objTRVVoucherInfo.sDestinoVou
    Idioma.Caption = objTRVVoucherInfo.sIdioma
    Convenio.Caption = objTRVVoucherInfo.sConvenio
    
    Traz_TRVVoucherInfo_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVVoucherInfo_Tela:

    Traz_TRVVoucherInfo_Tela = gErr

    Select Case gErr

        Case 190882

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190883)

    End Select

    Exit Function

End Function

Sub BotaoExtrairSigav_Click()

Dim lErro As Long
Dim objTRVVoucherInfo As New ClassTRVVoucherInfo
Dim objTRVVoucher As New ClassTRVVouchers
Dim objSenha As New ClassSenha

On Error GoTo Erro_BotaoExtrairSigav_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Move_Tela_Memoria(objTRVVoucher)
    If lErro <> SUCESSO Then gError 192640
    
    objTRVVoucherInfo.sTipo = objTRVVoucher.sTipVou
    objTRVVoucherInfo.lNumVou = objTRVVoucher.lNumVou
    objTRVVoucherInfo.sSerie = objTRVVoucher.sSerie
    
    'SetTimer hWnd, NV_INPUTBOX, 10, AddressOf TimerProc
    'sSenha = InputBox("Digite a senha do Sigav", "Extração de Dados")
    
    Load SigavSenha
    lErro = SigavSenha.Trata_Parametros(objSenha)
    If lErro <> SUCESSO Then gError 192669
    SigavSenha.Show vbModal
    
    If Len(Trim(objSenha.sSenha)) = 0 Then gError 192668
    
    lErro = Obter_Dados_Sigav(objTRVVoucherInfo, objSenha.sSenha)
    If lErro <> SUCESSO Then gError 192641
    
    lErro = CF("TRVVoucherInfoSigav_Grava", objTRVVoucherInfo)
    If lErro <> SUCESSO Then gError 192642
    
    lErro = Traz_TRVVoucherInfo_Tela(objTRVVoucherInfo)
    If lErro <> SUCESSO Then gError 192643

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExtrairSigav_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 192640 To 192643
        
        Case 192668
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190895)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190886)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_TRVVouchers

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
Dim objVou As New ClassTRVVouchers
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoCancelar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 190893
    
    lErro = CF("TRVVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198200
    
    If lErro <> SUCESSO Then gError 198201
    
    If objVou.lNumFatCoinfo = 0 Then gError 200712
    
    If UCase(BotaoCancelar.Caption) = "CANCELAR" Then
    
        If Status.Caption = STATUS_TRV_OCR_CANCELADO_TEXTO Then gError 192650
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_CANCELAMENTO_TRVVOUCHERS")
    
        If vbMsgRes = vbYes Then
        
            If Abs(objVou.dValorOcr) > DELTA_VALORMONETARIO Then
                'Pergunta ao usuário se confirma a exclusão
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VOUCHER_COM_OCRS")
                If vbMsgRes = vbNo Then gError 190893
            End If
    
            'Cancela o voucher
            lErro = CF("TRVVoucher_Exclui", objVou)
            If lErro <> SUCESSO Then gError 190894
    
            'Limpa Tela
            Call Limpa_Tela_TRVVouchers
            
            Call Trata_Parametros(objVou)
            
            Call Rotina_Aviso(vbOKOnly, "AVISO_VOUCHER_CANCELADO")
    
        End If
        
    Else
    
        If Status.Caption = STATUS_TRV_VOU_ABERTO_TEXTO Then gError 200711
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_REATIVAMENTO_TRVVOUCHERS")
    
        If vbMsgRes = vbYes Then
    
            'Reativa o voucher
            lErro = CF("TRVVoucher_Reativa", objVou)
            If lErro <> SUCESSO Then gError 190894
    
            'Limpa Tela
            Call Limpa_Tela_TRVVouchers
            
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

        Case 200712
            Call Rotina_Erro(vbOKOnly, "ERRO_USE_SW_CANC_VOU_NAO_FAT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190895)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVoucher_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTRVVouchers As ClassTRVVouchers

On Error GoTo Erro_objEventoVoucher_evSelecao

    Set objTRVVouchers = obj1

    'Mostra os dados do TRVVouchers na tela
    lErro = Traz_TRVVouchers_Tela(objTRVVouchers)
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
Dim objVoucher As New ClassTRVVouchers
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVou.Caption)
    objVoucher.sSerie = SerieVou.Caption
    objVoucher.sTipVou = TipoVou.Caption

    Call Chama_Tela("VoucherRapidoLista", colSelecao, objVoucher, objEventoVoucher)

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

Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirCli_Click

    objcliente.lCodigo = LCodigo_Extrai(ClienteFat.Caption)

    Call Chama_Tela("Clientes", objcliente)

    Exit Sub

Erro_BotaoAbrirCli_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirCli2_Click()

Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirCli2_Click

    objcliente.lCodigo = LCodigo_Extrai(ClienteVou.Caption)

    Call Chama_Tela("Clientes", objcliente)

    Exit Sub

Erro_BotaoAbrirCli2_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirFat_Click()

Dim lErro As Long
'Dim objObjeto As Object
'Dim sTela As String
'Dim bExisteDestino As Boolean
'Dim lNumTitulo As Long
'Dim sDoc As String
Dim objTRVVouchers As New ClassTRVVouchers

On Error GoTo Erro_BotaoAbrirFat_Click

    If Len(Trim(NumeroFat.Caption)) <> 0 Then

        objTRVVouchers.lNumVou = NumeroVou.Caption
        objTRVVouchers.sTipVou = TipoVou.Caption
        objTRVVouchers.sSerie = SerieVou.Caption
        objTRVVouchers.sTipoDoc = TRV_TIPODOC_VOU_TEXTO

        lErro = CF("TRVVouchers_Le", objTRVVouchers)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192882
        
        lErro = Abrir_Doc_Destino(objTRVVouchers.lNumIntDocDestino, objTRVVouchers.iTipoDocDestino, StrParaLong(NumeroFat.Caption))
        If lErro <> SUCESSO Then gError 192883

'        lErro = CF("Verifica_Existencia_Doc_TRV", objTRVVouchers.lNumIntDocDestino, objTRVVouchers.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
'        If lErro <> SUCESSO Then gError 192883
'
'        Select Case objTRVVouchers.iTipoDocDestino
'
'            Case TRV_TIPO_DOC_DESTINO_CREDFORN
'                sTela = TRV_TIPO_DOC_DESTINO_CREDFORN_TELA
'                Set objObjeto = New ClassCreditoPagar
'
'            Case TRV_TIPO_DOC_DESTINO_DEBCLI
'                sTela = TRV_TIPO_DOC_DESTINO_DEBCLI_TELA
'                Set objObjeto = New ClassDebitoRecCli
'
'            Case TRV_TIPO_DOC_DESTINO_TITPAG
'                sTela = TRV_TIPO_DOC_DESTINO_TITPAG_TELA
'                Set objObjeto = New ClassTituloPagar
'
'            Case TRV_TIPO_DOC_DESTINO_TITREC
'                sTela = TRV_TIPO_DOC_DESTINO_TITREC_TELA
'                Set objObjeto = New ClassTituloReceber
'
'            Case TRV_TIPO_DOC_DESTINO_NFSPAG
'                sTela = TRV_TIPO_DOC_DESTINO_NFSPAG_TELA
'                Set objObjeto = New ClassNFsPag
'
'        End Select
'
'        If Not (objObjeto Is Nothing) Then
'
'            objObjeto.lNumIntDoc = objTRVVouchers.lNumIntDocDestino
'
'            Call Chama_Tela(sTela, objObjeto)
'
'        End If
        
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

Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirPax_Click

    objcliente.lCodigo = LCodigo_Extrai(CliPassageiro.Caption)

    Call Chama_Tela("Clientes", objcliente)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objProduto.sCodigo = sProdutoFormatado

    Call Chama_Tela("Produto", objProduto)

    Exit Sub

Erro_BotaoAbrirProd_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
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
    objGridInt.colColuna.Add ("NC/FAT")
    objGridInt.colColuna.Add ("NF")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Histórico")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Rep.Name)
    objGridInt.colCampo.Add (DataRepr.Name)
    objGridInt.colCampo.Add (ValorRepr.Name)
    objGridInt.colCampo.Add (NumTitRepr.Name)
    objGridInt.colCampo.Add (NumNFRepr.Name)
    objGridInt.colCampo.Add (StatusRepr.Name)
    objGridInt.colCampo.Add (HistoricoRepr.Name)

    'Colunas do GridRepr
    iGrid_Rep_Col = 1
    iGrid_DataRepr_Col = 2
    iGrid_ValorRepr_Col = 3
    iGrid_NumTitRepr_Col = 4
    iGrid_NumNFRepr_Col = 5
    iGrid_StatusRepr_Col = 6
    iGrid_HistoricoRepr_Col = 7

    'Grid do GridInterno
    objGridInt.objGrid = GridRepr

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridRepr.ColWidth(0) = 200

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
    Call Traz_Resumo_Tela(2, GridRepr.Row)
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
    objGridInt.colColuna.Add ("NC/FAT")
    objGridInt.colColuna.Add ("NF")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Histórico")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Corr.Name)
    objGridInt.colCampo.Add (DataCorr.Name)
    objGridInt.colCampo.Add (ValorCorr.Name)
    objGridInt.colCampo.Add (NumTitCorr.Name)
    objGridInt.colCampo.Add (NumNFCorr.Name)
    objGridInt.colCampo.Add (StatusCorr.Name)
    objGridInt.colCampo.Add (HistoricoCorr.Name)

    'Colunas do GridRepr
    iGrid_Corr_Col = 1
    iGrid_DataCorr_Col = 2
    iGrid_ValorCorr_Col = 3
    iGrid_NumTitCorr_Col = 4
    iGrid_NumNFCorr_Col = 5
    iGrid_StatusCorr_Col = 6
    iGrid_HistoricoCorr_Col = 7

    'Grid do GridInterno
    objGridInt.objGrid = GridCorr

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridCorr.ColWidth(0) = 200

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
    Call Traz_Resumo_Tela(3, GridCorr.Row)
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
    objGridInt.colColuna.Add ("NC")
    objGridInt.colColuna.Add ("NF")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Histórico")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Emi.Name)
    objGridInt.colCampo.Add (DataEmissor.Name)
    objGridInt.colCampo.Add (ValorEmissor.Name)
    objGridInt.colCampo.Add (NumTitEmissor.Name)
    objGridInt.colCampo.Add (NumNFEmissor.Name)
    objGridInt.colCampo.Add (StatusEmissor.Name)
    objGridInt.colCampo.Add (HistoricoEmissor.Name)

    'Colunas do GridRepr
    iGrid_Emi_Col = 1
    iGrid_DataEmissor_Col = 2
    iGrid_ValorEmissor_Col = 3
    iGrid_NumTitEmissor_Col = 4
    iGrid_NumNFEmissor_Col = 5
    iGrid_StatusEmissor_Col = 6
    iGrid_HistoricoEmissor_Col = 7

    'Grid do GridInterno
    objGridInt.objGrid = GridEmissor

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

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
    Call Traz_Resumo_Tela(0, GridEmissor.Row)
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

    'Controles que participam do Grid
    objGridInt.colCampo.Add (VendedorPromo.Name)
    objGridInt.colCampo.Add (DataPromo.Name)
    objGridInt.colCampo.Add (ValorPromoBase.Name)
    objGridInt.colCampo.Add (ValorPromoComiss.Name)

    'Colunas do GridRepr
    iGrid_VendedorPromo_Col = 1
    iGrid_DataPromo_Col = 2
    iGrid_ValorPromoBase_Col = 3
    iGrid_ValorPromoComiss_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridPromotor

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 14

    'Largura da primeira coluna
    GridPromotor.ColWidth(0) = 400

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

Private Function Traz_Grids_Tela(objTRVVouchers As ClassTRVVouchers) As Long

Dim objTRVVoucherInfoN As ClassTRVVoucherInfoN
Dim iLinha As Integer
Dim iLinha1 As Integer
Dim iLinha2 As Integer
Dim iLinha3 As Integer
Dim objTRVGerComiIntDet As ClassTRVGerComiIntDet
Dim lErro As Long
Dim dValorR As Double
Dim dValorC As Double
Dim dValorE As Double
Dim dValorP As Double
Dim objcliente As New ClassCliente
Dim objForn As New ClassFornecedor
Dim dValorCC As Double
Dim dValorA As Double
Dim dValorB As Double

On Error GoTo Erro_Traz_Grids_Tela

    Call Grid_Limpa(objGridRepr)
    Call Grid_Limpa(objGridPromotor)
    Call Grid_Limpa(objGridEmissor)
    Call Grid_Limpa(objGridCorr)
    Call Grid_Limpa(objGridAg)
    
    Set gcolCMR = New Collection
    Set gcolCMC = New Collection
    Set gcolCMCC = New Collection
    Set gcolCME = New Collection
    Set gcolCMP = New Collection

    For Each objTRVVoucherInfoN In objTRVVouchers.colTRVVoucherInfo

        If objTRVVoucherInfoN.sTipoDoc = TRV_TIPODOC_CMA_TEXTO Then
            dValorA = dValorA + objTRVVoucherInfoN.dValor
        End If

        If objTRVVoucherInfoN.sTipoDoc = TRV_TIPODOC_BRUTO_TEXTO Then
            dValorB = dValorB + objTRVVoucherInfoN.dValor
        End If
        
        If objTRVVoucherInfoN.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Then
            
            iLinha = iLinha + 1
            
            'Busca o Cliente no BD
            If objTRVVoucherInfoN.lCliForn <> 0 Then
                Cliente.Text = objTRVVoucherInfoN.lCliForn
                lErro = TP_Cliente_Le2(Cliente, objcliente)
                If lErro <> SUCESSO Then gError 198232
            End If
            
            GridRepr.TextMatrix(iLinha, iGrid_Rep_Col) = Cliente.Text
            GridRepr.TextMatrix(iLinha, iGrid_DataRepr_Col) = Format(objTRVVoucherInfoN.dtData, "dd/mm/yyyy")
            GridRepr.TextMatrix(iLinha, iGrid_ValorRepr_Col) = Format(objTRVVoucherInfoN.dValor, "Standard")
            
            If objTRVVoucherInfoN.lNumTitulo <> 0 Then
                GridRepr.TextMatrix(iLinha, iGrid_NumTitRepr_Col) = CStr(objTRVVoucherInfoN.lNumTitulo)
            Else
                GridRepr.TextMatrix(iLinha, iGrid_NumTitRepr_Col) = ""
            End If
            
            If objTRVVoucherInfoN.lNumNF <> 0 Then
                GridRepr.TextMatrix(iLinha, iGrid_NumNFRepr_Col) = CStr(objTRVVoucherInfoN.lNumNF)
            Else
                GridRepr.TextMatrix(iLinha, iGrid_NumNFRepr_Col) = ""
            End If
            
            GridRepr.TextMatrix(iLinha, iGrid_HistoricoRepr_Col) = objTRVVoucherInfoN.sHistorico
            GridRepr.TextMatrix(iLinha, iGrid_StatusRepr_Col) = TRVVoucherInfo_Converte_Status(objTRVVoucherInfoN.iStatus)
    
            dValorR = dValorR + objTRVVoucherInfoN.dValor
    
            gcolCMR.Add objTRVVoucherInfoN
    
        End If
            
        If objTRVVoucherInfoN.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Then
            
            iLinha1 = iLinha1 + 1
            
            'Busca o Cliente no BD
            If objTRVVoucherInfoN.lCliForn <> 0 Then
                Cliente.Text = objTRVVoucherInfoN.lCliForn
                lErro = TP_Cliente_Le2(Cliente, objcliente)
                If lErro <> SUCESSO Then gError 198232
            End If
            
            GridCorr.TextMatrix(iLinha1, iGrid_Corr_Col) = Cliente.Text
            GridCorr.TextMatrix(iLinha1, iGrid_DataCorr_Col) = Format(objTRVVoucherInfoN.dtData, "dd/mm/yyyy")
            GridCorr.TextMatrix(iLinha1, iGrid_ValorCorr_Col) = Format(objTRVVoucherInfoN.dValor, "Standard")
            
            If objTRVVoucherInfoN.lNumTitulo <> 0 Then
                GridCorr.TextMatrix(iLinha1, iGrid_NumTitCorr_Col) = CStr(objTRVVoucherInfoN.lNumTitulo)
            Else
                GridCorr.TextMatrix(iLinha1, iGrid_NumTitCorr_Col) = ""
            End If
            
            If objTRVVoucherInfoN.lNumNF <> 0 Then
                GridCorr.TextMatrix(iLinha1, iGrid_NumNFCorr_Col) = CStr(objTRVVoucherInfoN.lNumNF)
            Else
                GridCorr.TextMatrix(iLinha1, iGrid_NumNFCorr_Col) = ""
            End If
            
            GridCorr.TextMatrix(iLinha1, iGrid_HistoricoCorr_Col) = objTRVVoucherInfoN.sHistorico
            GridCorr.TextMatrix(iLinha1, iGrid_StatusCorr_Col) = TRVVoucherInfo_Converte_Status(objTRVVoucherInfoN.iStatus)
    
            dValorC = dValorC + objTRVVoucherInfoN.dValor
    
            gcolCMC.Add objTRVVoucherInfoN
    
        End If
            
        If objTRVVoucherInfoN.sTipoDoc = TRV_TIPODOC_OVER_TEXTO Then
            
            iLinha2 = iLinha2 + 1
            
            'Busca o Cliente no BD
            If objTRVVoucherInfoN.lCliForn <> 0 Then
                Cliente.Text = objTRVVoucherInfoN.lCliForn
                lErro = TP_Fornecedor_Le2(Cliente, objForn)
                If lErro <> SUCESSO Then gError 198232
            End If
            
            GridEmissor.TextMatrix(iLinha2, iGrid_Emi_Col) = Cliente.Text
            GridEmissor.TextMatrix(iLinha2, iGrid_DataEmissor_Col) = Format(objTRVVoucherInfoN.dtData, "dd/mm/yyyy")
            GridEmissor.TextMatrix(iLinha2, iGrid_ValorEmissor_Col) = Format(objTRVVoucherInfoN.dValor, "Standard")
            
            If objTRVVoucherInfoN.lNumTitulo <> 0 Then
                GridEmissor.TextMatrix(iLinha2, iGrid_NumTitEmissor_Col) = CStr(objTRVVoucherInfoN.lNumTitulo)
            Else
                GridEmissor.TextMatrix(iLinha2, iGrid_NumTitEmissor_Col) = ""
            End If
            
            If objTRVVoucherInfoN.lNumNF <> 0 Then
                GridEmissor.TextMatrix(iLinha2, iGrid_NumNFEmissor_Col) = CStr(objTRVVoucherInfoN.lNumNF)
            Else
                GridEmissor.TextMatrix(iLinha2, iGrid_NumNFEmissor_Col) = ""
            End If
            
            GridEmissor.TextMatrix(iLinha2, iGrid_HistoricoEmissor_Col) = objTRVVoucherInfoN.sHistorico
            GridEmissor.TextMatrix(iLinha2, iGrid_StatusEmissor_Col) = TRVVoucherInfo_Converte_Status(objTRVVoucherInfoN.iStatus)
    
            dValorE = dValorE + objTRVVoucherInfoN.dValor
    
            gcolCME.Add objTRVVoucherInfoN
    
        End If
        
        If objTRVVoucherInfoN.sTipoDoc = TRV_TIPODOC_CMCC_TEXTO Then
            
            iLinha3 = iLinha3 + 1
            
            'Busca o Cliente no BD
            If objTRVVoucherInfoN.lCliForn <> 0 Then
                Cliente.Text = objTRVVoucherInfoN.lCliForn
                lErro = TP_Fornecedor_Le2(Cliente, objForn)
                If lErro <> SUCESSO Then gError 198232
            End If
            
            GridAg.TextMatrix(iLinha3, iGrid_Emi_Col) = Cliente.Text
            GridAg.TextMatrix(iLinha3, iGrid_DataAg_Col) = Format(objTRVVoucherInfoN.dtData, "dd/mm/yyyy")
            GridAg.TextMatrix(iLinha3, iGrid_ValorAg_Col) = Format(objTRVVoucherInfoN.dValor, "Standard")
            
            If objTRVVoucherInfoN.lNumTitulo <> 0 Then
                GridAg.TextMatrix(iLinha3, iGrid_NumTitAg_Col) = CStr(objTRVVoucherInfoN.lNumTitulo)
            Else
                GridAg.TextMatrix(iLinha3, iGrid_NumTitAg_Col) = ""
            End If
            
            If objTRVVoucherInfoN.lNumNF <> 0 Then
                GridAg.TextMatrix(iLinha3, iGrid_NumNFAg_Col) = CStr(objTRVVoucherInfoN.lNumNF)
            Else
                GridAg.TextMatrix(iLinha3, iGrid_NumNFAg_Col) = ""
            End If
            
            GridAg.TextMatrix(iLinha3, iGrid_HistoricoAg_Col) = objTRVVoucherInfoN.sHistorico
            GridAg.TextMatrix(iLinha3, iGrid_StatusAg_Col) = TRVVoucherInfo_Converte_Status(objTRVVoucherInfoN.iStatus)
    
            dValorCC = dValorCC + objTRVVoucherInfoN.dValor
    
            gcolCMCC.Add objTRVVoucherInfoN
    
        End If
    
    Next

    lErro = CF("TRVGerComiIntDet_Le_Promotor", objTRVVouchers)
    If lErro <> SUCESSO Then gError 197222
    
    objGridRepr.iLinhasExistentes = iLinha
    objGridCorr.iLinhasExistentes = iLinha1
    objGridEmissor.iLinhasExistentes = iLinha2
    objGridAg.iLinhasExistentes = iLinha3
    
    iLinha = 0
    
    For Each objTRVGerComiIntDet In objTRVVouchers.colTRVGerComiIntDet
    
        iLinha = iLinha + 1
        
        GridPromotor.TextMatrix(iLinha, iGrid_VendedorPromo_Col) = objTRVGerComiIntDet.sNomeReduzidoVendedor
        GridPromotor.TextMatrix(iLinha, iGrid_DataPromo_Col) = Format(objTRVGerComiIntDet.dtDataGeracao, "dd/mm/yyyy")
        GridPromotor.TextMatrix(iLinha, iGrid_ValorPromoBase_Col) = Format(objTRVGerComiIntDet.dValorBase, "Standard")
        GridPromotor.TextMatrix(iLinha, iGrid_ValorPromoComiss_Col) = Format(objTRVGerComiIntDet.dValorComissao, "Standard")
    
        dValorP = dValorP + objTRVGerComiIntDet.dValorComissao
    
        gcolCMP.Add objTRVGerComiIntDet
    
    Next
    
    objGridPromotor.iLinhasExistentes = iLinha
    
    TotalComiPro.Caption = Format(dValorP, "STANDARD")
    TotalComiEmi.Caption = Format(dValorE, "STANDARD")
    TotalComiRep.Caption = Format(dValorR, "STANDARD")
    TotalComiCor.Caption = Format(dValorC, "STANDARD")
    TotalComiAg.Caption = Format(dValorCC, "STANDARD")

    MoedaRV.Caption = Moeda.Caption
    TarifaUNRV.Caption = TarifaUN.Caption
    CambioRV.Caption = Cambio.Caption
    
    CMIReal.Caption = Format(dValorP, "STANDARD")
    OverReal.Caption = Format(dValorE, "STANDARD")
    CMRReal.Caption = Format(dValorR, "STANDARD")
    CMCReal.Caption = Format(dValorC, "STANDARD")
    OCRReal.Caption = Format(dValorB, "STANDARD")
    OCRCMAReal.Caption = Format(dValorA, "STANDARD")
    BrutoReal.Caption = Format(objTRVVouchers.dValorBruto, "STANDARD")
    CMCCReal.Caption = Format(dValorCC, "STANDARD")
    
    If objTRVVouchers.iCartao = DESMARCADO Then
        CMAReal.Caption = Format(objTRVVouchers.dValorComissao, "STANDARD")
    Else
        CMAReal.Caption = Format(0, "STANDARD")
    End If
    
    FATReal.Caption = Format(objTRVVouchers.dValor, "STANDARD")
    LIQCMIReal.Caption = Format(dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC, "STANDARD")
    LiqFinalReal.Caption = Format(dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC - dValorP, "STANDARD")
    
    If objTRVVouchers.dCambio <> 0 Then
        CMIMoeda.Caption = Format(dValorP / objTRVVouchers.dCambio, "STANDARD")
        OverMoeda.Caption = Format(dValorE / objTRVVouchers.dCambio, "STANDARD")
        CMRMoeda.Caption = Format(dValorR / objTRVVouchers.dCambio, "STANDARD")
        CMCMoeda.Caption = Format(dValorC / objTRVVouchers.dCambio, "STANDARD")
        OCRMoeda.Caption = Format(dValorB / objTRVVouchers.dCambio, "STANDARD")
        OCRCMAMoeda.Caption = Format(dValorA / objTRVVouchers.dCambio, "STANDARD")
        BrutoMoeda.Caption = Format(objTRVVouchers.dValorCambio, "STANDARD")
        CMCCMoeda.Caption = Format(dValorCC / objTRVVouchers.dCambio, "STANDARD")
        
        If objTRVVouchers.iCartao = DESMARCADO Then
            CMAMoeda.Caption = Format(objTRVVouchers.dValorComissao / objTRVVouchers.dCambio, "STANDARD")
        Else
            CMAMoeda.Caption = Format(0, "STANDARD")
        End If
    
        FATMoeda.Caption = Format(objTRVVouchers.dValor / objTRVVouchers.dCambio, "STANDARD")
        LIQCMIMoeda.Caption = Format((dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC) / objTRVVouchers.dCambio, "STANDARD")
        LiqFinalMoeda.Caption = Format((dValorB - dValorA - dValorE - dValorR - dValorC - dValorCC - dValorP) / objTRVVouchers.dCambio, "STANDARD")
    
    End If
    
    If dValorB <> 0 Then
        OverPerc.Caption = Format(dValorE / (dValorB), "PERCENT")
        CMRPerc.Caption = Format(dValorR / (dValorB), "PERCENT")
        CMCPerc.Caption = Format(dValorC / (dValorB), "PERCENT")
        CMCCPerc.Caption = Format(dValorCC / (dValorB), "PERCENT")
        CMAPerc.Caption = Format(dValorA / (dValorB), "PERCENT")
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
Dim objVoucher As New ClassTRVVouchers

On Error GoTo Erro_BotaoTrazerVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVouP.Text)
    objVoucher.sSerie = SerieVouP.Text
    objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
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

    Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "NumVou= ? AND TipVou = ? AND Serie = ?")

    Exit Sub

Erro_BotaoHist_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Function Abrir_Doc_Destino(ByVal lNumIntDocDestino As Long, ByVal iTipoDocDestino As Integer, Optional ByVal lNumFat As Long = 0)

Dim lErro As Long
Dim objObjeto As Object
Dim sTela As String
Dim bExisteDestino As Boolean
Dim lNumTitulo As Long
Dim sDoc As String
Dim objFat As New ClassFaturaTRV

On Error GoTo Erro_Abrir_Doc_Destino

    If lNumIntDocDestino <> 0 Then

        If lNumFat = 0 Then
            
            lErro = CF("Verifica_Existencia_Doc_TRV", lNumIntDocDestino, iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
            If lErro <> SUCESSO Then gError 196387
            
            objFat.lNumFat = lNumTitulo
        Else
            objFat.lNumFat = lNumFat
        End If
        
        Call Chama_Tela("TRVConsultaFatura", objFat)
        
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
Dim objVou As New ClassTRVVouchers

On Error GoTo Erro_BotaoComissao_Click

    objVou.lNumVou = StrParaLong(NumeroVouP.Text)
    objVou.sSerie = SerieVouP.Text
    objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVou.sTipVou = TipoVouP.Text
    
    Call Chama_Tela("TRVVouComi", objVou)

    Exit Sub

Erro_BotaoComissao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHistImport_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistImport_Click

    colSelecao.Add StrParaLong(NumeroVou.Caption)
    colSelecao.Add TipoVou.Caption
    colSelecao.Add SerieVou.Caption

    Call Chama_Tela("TRVHistVouArqLista", colSelecao, Nothing, Nothing, "NumVou= ? AND TipVou = ? AND Serie = ?")

    Exit Sub

Erro_BotaoHistImport_Click:

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
    Call Chama_Tela("VendedorLista", colSelecao, Nothing, Nothing, "Vendedor IN (SELECT Promotor FROM TRVVouchers WHERE NumVou = ? AND TipVou = ? AND Serie = ?)  OR Vendedor IN (SELECT Vendedor FROM TRVVouVendedores WHERE TipVou = ? AND Serie = ? AND NumVou = ?)")

End Sub

Public Sub BotaoCTB_Click()

Dim lErro As Long
Dim objVou As New ClassTRVVouchers
Dim objcliente As New ClassCliente
Dim objClienteTRV As ClassClienteTRV

On Error GoTo Erro_BotaoCTB_Click

    objVou.lNumVou = StrParaLong(NumeroVouP.Text)
    objVou.sSerie = SerieVouP.Text
    objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVou.sTipVou = TipoVouP.Text

    lErro = CF("TRVVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192646
    
    objcliente.lCodigo = objVou.lClienteVou
    
    lErro = CF("Cliente_Le_Customizado", objcliente)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192646
    
    Set objClienteTRV = objcliente.objInfoUsu
    
    Call CF("Lancamentos_Abre_Tela", 23, objVou.lNumIntDoc, 0, 0)
    
    Exit Sub

Erro_BotaoCTB_Click:

    Select Case gErr
    
        Case 192646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Public Function Traz_Resumo_Tela(ByVal iIndex As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim colAux As Collection
Dim objVouInfo As ClassTRVVoucherInfoN

On Error GoTo Erro_Traz_Resumo_Tela

    If iLinha > 0 Then
        
        Select Case iIndex
            Case 0
                Set colAux = gcolCME
            Case 1
                Set colAux = gcolCMCC
            Case 2
                Set colAux = gcolCMR
            Case 3
                Set colAux = gcolCMC
            Case 4
                Set colAux = gcolCMP
        End Select
        
        If Not (colAux Is Nothing) Then
        
            If iLinha <= colAux.Count Then
            
                Set objVouInfo = colAux.Item(iLinha)
                Call Limpa_Resumo(0)
                
                If objVouInfo.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_NFSPAG Then
                
                    If objVouInfo.lNumTitulo <> 0 Then
                        ResNC(iIndex).Caption = CStr(objVouInfo.lNumTitulo)
                        ResNCData(iIndex).Caption = Format(objVouInfo.dtDataNC, "dd/mm/yyyy")
                        ResNCValor(iIndex).Caption = Format(objVouInfo.dValorNC, "STANDARD")
                        ResNCForn(iIndex).Caption = CStr(objVouInfo.lFornNC) & SEPARADOR & objVouInfo.sFornNC
                    End If
                
                    If objVouInfo.lNumNF <> 0 Then
                        ResNF(iIndex).Caption = CStr(objVouInfo.lNumNF)
                        ResNFData(iIndex).Caption = Format(objVouInfo.dtDataNF, "dd/mm/yyyy")
                        ResNFValor(iIndex).Caption = Format(objVouInfo.dValorNF, "STANDARD")
                        ResNFForn(iIndex).Caption = CStr(objVouInfo.lFornNF) & SEPARADOR & objVouInfo.sFornNF
                    End If
                    
                End If
                    
                ResUsu(iIndex).Caption = objVouInfo.sUsuario
                ResData(iIndex).Caption = Format(objVouInfo.dtDataRegistro, "dd/mm/yyyy")
                ResHora(iIndex).Caption = Format(objVouInfo.dHoraRegistro, "hh:mm:ss")
                ResHist(iIndex).Caption = objVouInfo.sHistorico
            
            End If
        
        End If
        
    End If
   
    Traz_Resumo_Tela = SUCESSO

    Exit Function

Erro_Traz_Resumo_Tela:

    Traz_Resumo_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197225)

    End Select

    Exit Function

End Function

Private Sub Limpa_Resumo(ByVal iIndex As Integer)
                    
    ResNC(iIndex).Caption = ""
    ResNCData(iIndex).Caption = ""
    ResNCValor(iIndex).Caption = ""
    ResNCForn(iIndex).Caption = ""
    ResNF(iIndex).Caption = ""
    ResNFData(iIndex).Caption = ""
    ResNFValor(iIndex).Caption = ""
    ResNFForn(iIndex).Caption = ""
    ResUsu(iIndex).Caption = ""
    ResData(iIndex).Caption = ""
    ResHora(iIndex).Caption = ""
    ResHist(iIndex).Caption = ""

End Sub

Private Function Inicializa_Grid_AG(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Ag.")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("NC/FAT")
    objGridInt.colColuna.Add ("NF")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Histórico")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Ag.Name)
    objGridInt.colCampo.Add (DataAg.Name)
    objGridInt.colCampo.Add (ValorAg.Name)
    objGridInt.colCampo.Add (NumTitAg.Name)
    objGridInt.colCampo.Add (NumNFAg.Name)
    objGridInt.colCampo.Add (StatusAg.Name)
    objGridInt.colCampo.Add (HistoricoAg.Name)

    'Colunas do GridRepr
    iGrid_Ag_Col = 1
    iGrid_DataAg_Col = 2
    iGrid_ValorAg_Col = 3
    iGrid_NumTitAg_Col = 4
    iGrid_NumNFAg_Col = 5
    iGrid_StatusAg_Col = 6
    iGrid_HistoricoAg_Col = 7

    'Grid do GridInterno
    objGridInt.objGrid = GridAg

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridAg.ColWidth(0) = 200

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_AG = SUCESSO

    Exit Function

End Function

Public Sub GridAG_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAg, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAg, iAlterado)
    End If
    
End Sub

Public Sub GridAG_EnterCell()
    Call Grid_Entrada_Celula(objGridAg, iAlterado)
End Sub

Public Sub GridAG_GotFocus()
    Call Grid_Recebe_Foco(objGridAg)
End Sub

Public Sub GridAG_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridAg)
    
End Sub

Public Sub GridAG_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAg, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAg, iAlterado)
    End If
    
End Sub

Public Sub GridAG_LeaveCell()
    Call Saida_Celula(objGridAg)
End Sub

Public Sub GridAG_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridAg)
End Sub

Public Sub GridAG_RowColChange()
    Call Grid_RowColChange(objGridAg)
    Call Traz_Resumo_Tela(1, GridAg.Row)
End Sub

Public Sub GridAG_Scroll()
    Call Grid_Scroll(objGridAg)
End Sub

Private Sub BotaoAbrirAG_Click()
    Call BotaoAbrirCli2_Click
End Sub

Private Sub BotaoAbrirNC_Click(Index As Integer)

Dim lErro As Long
Dim objVouInfo As ClassTRVVoucherInfoN
Dim colAux As Collection
Dim objGridAux As AdmGrid

On Error GoTo Erro_BotaoAbrirNC_Click

    Select Case Index
        Case 0
            Set colAux = gcolCME
            Set objGridAux = objGridEmissor
        Case 1
            Set colAux = gcolCMCC
            Set objGridAux = objGridAg
        Case 2
            Set colAux = gcolCMR
            Set objGridAux = objGridRepr
        Case 3
            Set colAux = gcolCMC
            Set objGridAux = objGridCorr
    End Select

    If objGridAux.objGrid.Row = 0 Or objGridAux.objGrid.Row > objGridAux.iLinhasExistentes Then gError 196389
    
    Set objVouInfo = colAux.Item(objGridAux.objGrid.Row)
    
    If objVouInfo.lNumIntDocDestino <> 0 Then
        
        lErro = Abrir_Doc_Destino(objVouInfo.lNumIntDocDestino, objVouInfo.iTipoDocDestino, objVouInfo.lNumTitulo)
        If lErro <> SUCESSO Then gError 196390
        
    End If
    
    Exit Sub

Erro_BotaoAbrirNC_Click:

    Select Case gErr
        
        Case 196389
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 196390
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196391)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirNF_Click(Index As Integer)

Dim lErro As Long
Dim objVouInfo As ClassTRVVoucherInfoN
Dim colAux As Collection
Dim objGridAux As AdmGrid
Dim objObjeto As New ClassTituloPagar

On Error GoTo Erro_BotaoAbrirNF_Click

    Select Case Index
        Case 0
            Set colAux = gcolCME
            Set objGridAux = objGridEmissor
        Case 1
            Set colAux = gcolCMCC
            Set objGridAux = objGridAg
        Case 2
            Set colAux = gcolCMR
            Set objGridAux = objGridRepr
        Case 3
            Set colAux = gcolCMC
            Set objGridAux = objGridCorr
    End Select

    If objGridAux.objGrid.Row = 0 Or objGridAux.objGrid.Row > objGridAux.iLinhasExistentes Then gError 196389
    
    Set objVouInfo = colAux.Item(objGridAux.objGrid.Row)
    
    If objVouInfo.lNumIntDocNF <> 0 Then
        
        objObjeto.lNumIntDoc = objVouInfo.lNumIntDocNF
        
        Call Chama_Tela(TRV_TIPO_DOC_DESTINO_TITPAG_TELA, objObjeto)
        
    End If
    
    Exit Sub

Erro_BotaoAbrirNF_Click:

    Select Case gErr
        
        Case 196389
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 196390
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196391)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirCorr_Click()

Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirCorr_Click

    objcliente.lCodigo = LCodigo_Extrai(Correntista.Caption)

    Call Chama_Tela("Clientes", objcliente)

    Exit Sub

Erro_BotaoAbrirCorr_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirRep_Click()

Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirRep_Click

    objcliente.lCodigo = LCodigo_Extrai(Representante.Caption)

    Call Chama_Tela("Clientes", objcliente)

    Exit Sub

Erro_BotaoAbrirRep_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirProm_Click()

Dim objVend As New ClassVendedor

On Error GoTo Erro_BotaoAbrirProm_Click

    objVend.iCodigo = Codigo_Extrai(Promotor.Caption)

    Call Chama_Tela("Vendedores", objVend)

    Exit Sub

Erro_BotaoAbrirProm_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub
