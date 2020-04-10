VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl TabTributacaoFat 
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   8910
   Begin VB.Frame FrameTributacao 
      BorderStyle     =   0  'None
      Height          =   4215
      Index           =   2
      Left            =   105
      TabIndex        =   202
      Top             =   300
      Visible         =   0   'False
      Width           =   8745
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   2
         Left            =   60
         TabIndex        =   221
         Top             =   870
         Visible         =   0   'False
         Width           =   8595
         Begin VB.Frame FrameICMSST 
            Caption         =   "ICMS Substituição Tributária (com FCP incluso)"
            Height          =   1215
            Index           =   1
            Left            =   0
            TabIndex        =   242
            Top             =   2000
            Width           =   8580
            Begin VB.CheckBox ICMSSTBaseDupla 
               Caption         =   "Cálculo por ""Base Dupla"""
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
               Left            =   5745
               TabIndex        =   472
               ToolTipText     =   $"TabTributacaoFat.ctx":0000
               Top             =   210
               Width           =   3540
            End
            Begin VB.ComboBox ICMSSTModalidade 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "TabTributacaoFat.ctx":00D4
               Left            =   1815
               List            =   "TabTributacaoFat.ctx":00D6
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   180
               Width           =   3870
            End
            Begin VB.Frame FrameICMSSTUFDevido 
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   5580
               TabIndex        =   243
               Top             =   870
               Width           =   2940
               Begin VB.ComboBox ICMSUFDevidoST 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2295
                  TabIndex        =   64
                  Top             =   0
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "UF do ICMS ST Devido:"
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
                  Index           =   64
                  Left            =   210
                  TabIndex        =   244
                  Top             =   45
                  Width           =   2055
               End
            End
            Begin MSMask.MaskEdBox ICMSSubstBaseItem 
               Height          =   285
               Left            =   4590
               TabIndex        =   60
               Top             =   540
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSSubstAliquotaItem 
               Height          =   285
               Left            =   1815
               TabIndex        =   62
               Top             =   870
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSSubstValorItem 
               Height          =   285
               Left            =   4590
               TabIndex        =   63
               Top             =   870
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSSubstPercMVAItem 
               Height          =   285
               Left            =   1815
               TabIndex        =   59
               Top             =   540
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSSubstPercRedBaseItem 
               Height          =   285
               Left            =   7875
               TabIndex        =   61
               Top             =   540
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Modalidade:"
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
               Index           =   108
               Left            =   735
               TabIndex        =   250
               Top             =   225
               Width           =   1050
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base de Cálculo:"
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
               Index           =   30
               Left            =   3105
               TabIndex        =   249
               Top             =   585
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   181
               Left            =   990
               TabIndex        =   248
               Top             =   915
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   193
               Left            =   4050
               TabIndex        =   247
               Top             =   900
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "% MVA:"
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
               Left            =   1125
               TabIndex        =   246
               Top             =   585
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "% Redução na BC:"
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
               Left            =   6240
               TabIndex        =   245
               Top             =   585
               Width           =   1590
            End
         End
         Begin VB.Frame FrameICMSST 
            Caption         =   "ICMS Substituição Tributária Cobrado Anteriormente (com FCP incluso)"
            Height          =   1215
            Index           =   2
            Left            =   0
            TabIndex        =   239
            Top             =   2025
            Visible         =   0   'False
            Width           =   8580
            Begin MSMask.MaskEdBox ICMSSTCobrAntBaseItem 
               Height          =   285
               Left            =   1800
               TabIndex        =   65
               Top             =   330
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSSTCobrAntValorItem 
               Height          =   285
               Left            =   1785
               TabIndex        =   67
               Top             =   750
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSSTCobrAntAliquotaItem 
               Height          =   285
               Left            =   4590
               TabIndex        =   66
               Top             =   315
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   131
               Left            =   3735
               TabIndex        =   458
               Top             =   345
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   65
               Left            =   1245
               TabIndex        =   241
               Top             =   780
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base de Cálculo:"
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
               Index           =   67
               Left            =   315
               TabIndex        =   240
               Top             =   375
               Width           =   1440
            End
         End
         Begin VB.Frame FrameICMS 
            Caption         =   "ICMS (com FCP incluso)"
            Height          =   1215
            Index           =   1
            Left            =   0
            TabIndex        =   222
            Top             =   315
            Width           =   8580
            Begin VB.Frame FrameICMS51 
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   75
               TabIndex        =   223
               Top             =   1230
               Visible         =   0   'False
               Width           =   8385
               Begin MSMask.MaskEdBox ICMS51ValorOpItem 
                  Height          =   285
                  Left            =   4515
                  TabIndex        =   47
                  ToolTipText     =   "Valor como se não tivesse o diferimento - vICMSOp"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSPercDiferItem 
                  Height          =   285
                  Left            =   1725
                  TabIndex        =   46
                  ToolTipText     =   "Percentual do diferimento -  pDif"
                  Top             =   0
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSValorDifItem 
                  Height          =   285
                  Left            =   7335
                  TabIndex        =   48
                  ToolTipText     =   "Valor do ICMS diferido - vICMSDif"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS da Operação:"
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
                  Index           =   8
                  Left            =   2835
                  TabIndex        =   226
                  Top             =   45
                  Width           =   1665
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% do Diferimento:"
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
                  Index           =   16
                  Left            =   210
                  TabIndex        =   225
                  Top             =   15
                  Width           =   1500
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS Diferido:"
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
                  Index           =   27
                  Left            =   6045
                  TabIndex        =   224
                  Top             =   45
                  Width           =   1245
               End
            End
            Begin VB.ComboBox ICMSModalidade 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   165
               Width           =   3900
            End
            Begin VB.CheckBox ICMSCredita 
               Caption         =   "Debita"
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
               Left            =   5820
               TabIndex        =   45
               Top             =   930
               Width           =   936
            End
            Begin VB.Frame FrameICMSRedBase 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   285
               Left            =   2970
               TabIndex        =   229
               Top             =   540
               Width           =   2505
               Begin MSMask.MaskEdBox ICMSPercRedBaseItem 
                  Height          =   285
                  Left            =   1620
                  TabIndex        =   42
                  Top             =   0
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% Redução na BC:"
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
                  Index           =   42
                  Left            =   0
                  TabIndex        =   230
                  Top             =   45
                  Width           =   1590
               End
            End
            Begin MSMask.MaskEdBox ICMSBaseItem 
               Height          =   285
               Left            =   1800
               TabIndex        =   41
               Top             =   540
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSAliquotaItem 
               Height          =   285
               Left            =   1800
               TabIndex        =   43
               Top             =   885
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSValorItem 
               Height          =   285
               Left            =   4590
               TabIndex        =   44
               Top             =   885
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Frame FrameICMSBaseOperProp 
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   2295
               TabIndex        =   227
               Top             =   1230
               Width           =   3450
               Begin MSMask.MaskEdBox ICMSpercBaseOperacaoPropria 
                  Height          =   285
                  Left            =   2295
                  TabIndex        =   49
                  Top             =   0
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% BC da operação própria:"
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
                  Index           =   63
                  Left            =   0
                  TabIndex        =   228
                  Top             =   45
                  Width           =   2265
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Modalidade:"
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
               Index           =   107
               Left            =   720
               TabIndex        =   234
               Top             =   210
               Width           =   1050
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base de Cálculo:"
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
               Index           =   206
               Left            =   315
               TabIndex        =   233
               Top             =   585
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   196
               Left            =   975
               TabIndex        =   232
               Top             =   900
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   47
               Left            =   4065
               TabIndex        =   231
               Top             =   930
               Width           =   510
            End
         End
         Begin VB.Frame FrameICMS 
            Caption         =   "ICMS Desonerado"
            Height          =   480
            Index           =   2
            Left            =   0
            TabIndex        =   236
            Top             =   1530
            Visible         =   0   'False
            Width           =   8580
            Begin VB.ComboBox ICMSMotivo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "TabTributacaoFat.ctx":00D8
               Left            =   3705
               List            =   "TabTributacaoFat.ctx":00DA
               TabIndex        =   53
               Text            =   "ICMSMotivo"
               Top             =   135
               Width           =   4770
            End
            Begin MSMask.MaskEdBox ICMSValorIsento 
               Height          =   285
               Left            =   1800
               TabIndex        =   52
               Top             =   150
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   66
               Left            =   1260
               TabIndex        =   238
               Top             =   180
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Motivo:"
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
               Index           =   68
               Left            =   3030
               TabIndex        =   237
               Top             =   180
               Width           =   645
            End
         End
         Begin VB.Frame FrameICMSST 
            Caption         =   "ICMS Substituição Tributária"
            Height          =   1215
            Index           =   3
            Left            =   0
            TabIndex        =   254
            Top             =   2000
            Visible         =   0   'False
            Width           =   8580
            Begin VB.Frame Frame1 
               Caption         =   "Retido na UF remetente"
               Height          =   510
               Index           =   10
               Left            =   60
               TabIndex        =   258
               Top             =   180
               Width           =   8475
               Begin MSMask.MaskEdBox ICMSvBCSTRet 
                  Height          =   285
                  Left            =   1770
                  TabIndex        =   54
                  Top             =   180
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSvICMSSTRet 
                  Height          =   285
                  Left            =   4365
                  TabIndex        =   55
                  Top             =   180
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base de Cálculo:"
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
                  Index           =   69
                  Left            =   285
                  TabIndex        =   260
                  Top             =   225
                  Width           =   1440
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor:"
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
                  Index           =   70
                  Left            =   3825
                  TabIndex        =   259
                  Top             =   210
                  Width           =   510
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Da UF de Destino"
               Height          =   480
               Index           =   11
               Left            =   75
               TabIndex        =   255
               Top             =   675
               Width           =   8475
               Begin MSMask.MaskEdBox ICMSvBCSTDest 
                  Height          =   285
                  Left            =   1770
                  TabIndex        =   56
                  Top             =   150
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSvICMSSTDest 
                  Height          =   285
                  Left            =   4365
                  TabIndex        =   57
                  Top             =   150
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor:"
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
                  Index           =   71
                  Left            =   3825
                  TabIndex        =   257
                  Top             =   180
                  Width           =   510
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base de Cálculo:"
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
                  Index           =   72
                  Left            =   285
                  TabIndex        =   256
                  Top             =   195
                  Width           =   1440
               End
            End
         End
         Begin VB.ComboBox ComboICMSTipo 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   15
            Width           =   3336
         End
         Begin VB.ComboBox OrigemMercadoria 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5880
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   15
            Width           =   2670
         End
         Begin VB.ComboBox ComboICMSSimplesTipo 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   15
            Visible         =   0   'False
            Width           =   3336
         End
         Begin VB.Frame FrameICMSCred 
            Caption         =   "Crédito de ICMS"
            Height          =   465
            Left            =   0
            TabIndex        =   251
            Top             =   285
            Visible         =   0   'False
            Width           =   8580
            Begin MSMask.MaskEdBox ICMSvCredSN 
               Height          =   285
               Left            =   4410
               TabIndex        =   51
               Top             =   135
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSpCredSN 
               Height          =   285
               Left            =   1815
               TabIndex        =   50
               Top             =   135
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   74
               Left            =   3870
               TabIndex        =   253
               Top             =   165
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   73
               Left            =   990
               TabIndex        =   252
               Top             =   180
               Width           =   780
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Situação Tributária:"
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
            Index           =   102
            Left            =   75
            TabIndex        =   262
            Top             =   60
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Origem:"
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
            Index           =   103
            Left            =   5205
            TabIndex        =   261
            Top             =   60
            Width           =   660
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   1
         Left            =   30
         TabIndex        =   291
         Top             =   885
         Width           =   8640
         Begin VB.TextBox cBenefItem 
            Height          =   315
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   464
            Top             =   1170
            Width           =   1350
         End
         Begin VB.ComboBox NatBCCred 
            Height          =   315
            Left            =   3030
            Style           =   2  'Dropdown List
            TabIndex        =   461
            Top             =   810
            Width           =   5490
         End
         Begin VB.Frame Frame1 
            Caption         =   "Totais de Tributos"
            Height          =   795
            Index           =   13
            Left            =   90
            TabIndex        =   375
            Top             =   2460
            Width           =   8490
            Begin VB.ComboBox TotTribItemTipo 
               Height          =   315
               ItemData        =   "TabTributacaoFat.ctx":00DC
               Left            =   90
               List            =   "TabTributacaoFat.ctx":00E9
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   420
               Width           =   1905
            End
            Begin MSMask.MaskEdBox TotTribItem 
               Height          =   315
               Left            =   2340
               TabIndex        =   33
               ToolTipText     =   "Soma dos valores dos tributos correspondentes à este item para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TotTribAliqFedItem 
               Height          =   315
               Left            =   3945
               TabIndex        =   34
               ToolTipText     =   "Soma dos valores dos tributos correspondentes à este item para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TotTribAliqEstItem 
               Height          =   315
               Left            =   5565
               TabIndex        =   35
               ToolTipText     =   "Soma dos valores dos tributos correspondentes à este item para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TotTribAliqMunicItem 
               Height          =   315
               Left            =   7170
               TabIndex        =   36
               ToolTipText     =   "Soma dos valores dos tributos correspondentes à este item para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Municipal(%)"
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
               Index           =   80
               Left            =   7155
               TabIndex        =   380
               Top             =   210
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Estadual(%)"
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
               Index           =   79
               Left            =   5565
               TabIndex        =   379
               Top             =   210
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Federal(%)"
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
               Index           =   78
               Left            =   3945
               TabIndex        =   378
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total (R$)"
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
               Index           =   77
               Left            =   2325
               TabIndex        =   377
               Top             =   210
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fonte"
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
               Index           =   76
               Left            =   90
               TabIndex        =   376
               Top             =   210
               Width           =   495
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Outros Valores"
            Height          =   1020
            Index           =   8
            Left            =   90
            TabIndex        =   292
            Top             =   1455
            Width           =   8490
            Begin VB.Frame Frame1 
               Caption         =   "Unidade Tributável"
               Height          =   885
               Index           =   9
               Left            =   5310
               TabIndex        =   356
               Top             =   105
               Width           =   3105
               Begin VB.Label ValorTrib 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   361
                  Top             =   525
                  Width           =   1635
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor Unitário:"
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
                  Index           =   7
                  Left            =   90
                  TabIndex        =   360
                  Top             =   570
                  Width           =   1215
               End
               Begin VB.Label UMTrib 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   2370
                  TabIndex        =   359
                  Top             =   195
                  Width           =   630
               End
               Begin VB.Label QtdeTrib 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   358
                  Top             =   195
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Quantidade:"
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
                  Index           =   33
                  Left            =   255
                  TabIndex        =   357
                  Top             =   240
                  Width           =   1050
               End
            End
            Begin MSMask.MaskEdBox ValorFreteItem 
               Height          =   315
               Left            =   1050
               TabIndex        =   28
               Top             =   255
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorDescontoItem 
               Height          =   315
               Left            =   3810
               TabIndex        =   31
               Top             =   600
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               Enabled         =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorDespesasItem 
               Height          =   315
               Left            =   3810
               TabIndex        =   29
               Top             =   225
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorSeguroItem 
               Height          =   315
               Left            =   1050
               TabIndex        =   30
               Top             =   600
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Seguro:"
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
               Index           =   907
               Left            =   315
               TabIndex        =   296
               Top             =   630
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Desconto:"
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
               Index           =   150
               Left            =   2880
               TabIndex        =   295
               Top             =   615
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Frete:"
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
               Index           =   59
               Left            =   495
               TabIndex        =   294
               Top             =   300
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "O. Despesas:"
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
               Index           =   60
               Left            =   2610
               TabIndex        =   293
               Top             =   270
               Width           =   1155
            End
         End
         Begin MSMask.MaskEdBox NaturezaOpItem 
            Height          =   315
            Left            =   1080
            TabIndex        =   26
            Top             =   90
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoTributacaoItem 
            Height          =   315
            Left            =   1080
            TabIndex        =   27
            Top             =   435
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Código de Benefício Fiscal na UF:"
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
            Index           =   134
            Left            =   915
            TabIndex        =   463
            Top             =   1200
            Width           =   3330
         End
         Begin VB.Label Label1 
            Caption         =   "Código BC do Crédito:"
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
            Index           =   100
            Left            =   1095
            TabIndex        =   462
            Top             =   840
            Width           =   1950
         End
         Begin VB.Label DescTipoTribItem 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1545
            TabIndex        =   300
            Top             =   435
            Width           =   6975
         End
         Begin VB.Label LabelDescrNatOpItem 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1545
            TabIndex        =   299
            Top             =   90
            Width           =   6975
         End
         Begin VB.Label NaturezaItemLabel 
            AutoSize        =   -1  'True
            Caption         =   "CFOP:"
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
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   298
            Top             =   135
            Width           =   555
         End
         Begin VB.Label LblTipoTribItem 
            Caption         =   "Tipo Trib.:"
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   297
            Top             =   480
            Width           =   990
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   8
         Left            =   60
         TabIndex        =   381
         Top             =   855
         Visible         =   0   'False
         Width           =   8595
         Begin VB.Frame FrameICMSInterest 
            Caption         =   "ICMS em Operações Interestaduais"
            Height          =   3165
            Index           =   0
            Left            =   0
            TabIndex        =   382
            Top             =   60
            Width           =   8580
            Begin VB.CheckBox ICMSInterestBaseDupla 
               Caption         =   "Cálculo por ""Base Dupla"""
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
               Left            =   1815
               TabIndex        =   473
               ToolTipText     =   "Inclusão do ICMS relativo à alíquota interna no destino"
               Top             =   285
               Width           =   3540
            End
            Begin VB.ComboBox ICMSInterestPercPartilhaItem 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "TabTributacaoFat.ctx":0123
               Left            =   1815
               List            =   "TabTributacaoFat.ctx":013B
               Style           =   2  'Dropdown List
               TabIndex        =   312
               Top             =   1620
               Width           =   6690
            End
            Begin VB.ComboBox ICMSInterestAliqItem 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "TabTributacaoFat.ctx":0180
               Left            =   1815
               List            =   "TabTributacaoFat.ctx":0191
               Style           =   2  'Dropdown List
               TabIndex        =   311
               Top             =   1245
               Width           =   6690
            End
            Begin MSMask.MaskEdBox ICMSInterestBCUFDestItem 
               Height          =   285
               Left            =   1815
               TabIndex        =   309
               Top             =   555
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSInterestAliqUFDestItem 
               Height          =   285
               Left            =   1815
               TabIndex        =   310
               Top             =   885
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSInterestPercFCPUFDestItem 
               Height          =   285
               Left            =   7830
               TabIndex        =   308
               Top             =   525
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSInterestVlrUFDestItem 
               Height          =   285
               Left            =   1815
               TabIndex        =   313
               Top             =   1995
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSInterestVlrUFRemetItem 
               Height          =   285
               Left            =   1815
               TabIndex        =   314
               Top             =   2355
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSInterestVlrFCPUFDestItem 
               Height          =   285
               Left            =   1815
               TabIndex        =   315
               Top             =   2700
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSInterestBCFCPUFDestItem 
               Height          =   285
               Left            =   4845
               TabIndex        =   460
               Top             =   540
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "BC FCP UF Dest.:"
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
               Index           =   133
               Left            =   3300
               TabIndex        =   459
               Top             =   585
               Width           =   1530
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "BC ICMS UF Dest.:"
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
               Index           =   86
               Left            =   165
               TabIndex        =   390
               Top             =   555
               Width           =   1635
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Alíquota Interna:"
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
               Index           =   84
               Left            =   345
               TabIndex        =   389
               Top             =   900
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "% FCP UF Dest.:"
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
               Index           =   89
               Left            =   6375
               TabIndex        =   388
               Top             =   570
               Width           =   1425
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Alíquota Interest.:"
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
               Index           =   90
               Left            =   240
               TabIndex        =   387
               Top             =   1290
               Width           =   1545
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "% Partilha:"
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
               Index           =   81
               Left            =   870
               TabIndex        =   386
               Top             =   1665
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor UF Dest.:"
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
               Index           =   82
               Left            =   465
               TabIndex        =   385
               Top             =   2055
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor UF Remet.:"
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
               Index           =   83
               Left            =   315
               TabIndex        =   384
               Top             =   2385
               Width           =   1470
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor FCP UF Dest.:"
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
               Index           =   85
               Left            =   60
               TabIndex        =   383
               Top             =   2730
               Width           =   1725
            End
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   10
         Left            =   60
         TabIndex        =   466
         Top             =   855
         Visible         =   0   'False
         Width           =   8580
         Begin VB.Frame Frame1 
            Caption         =   "IPI"
            Height          =   900
            Index           =   20
            Left            =   120
            TabIndex        =   469
            Top             =   690
            Width           =   8340
            Begin MSMask.MaskEdBox IPIVlrDevolvidoItem 
               Height          =   285
               Left            =   1710
               TabIndex        =   470
               Top             =   330
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor Devolvido:"
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
               Index           =   135
               Left            =   225
               TabIndex        =   471
               Top             =   375
               Width           =   1425
            End
         End
         Begin MSMask.MaskEdBox pDevolItem 
            Height          =   285
            Left            =   1830
            TabIndex        =   467
            Top             =   300
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Devolvido:"
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
            Index           =   136
            Left            =   645
            TabIndex        =   468
            Top             =   345
            Width           =   1125
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   9
         Left            =   60
         TabIndex        =   406
         Top             =   870
         Visible         =   0   'False
         Width           =   8595
         Begin VB.Frame FrameFCP 
            Caption         =   " FCP retido anteriormente por ST"
            Height          =   750
            Index           =   2
            Left            =   0
            TabIndex        =   441
            Top             =   2145
            Width           =   8580
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   315
               Index           =   19
               Left            =   2295
               TabIndex        =   449
               Top             =   1230
               Width           =   3450
               Begin MSMask.MaskEdBox MaskEdBox18 
                  Height          =   285
                  Left            =   2295
                  TabIndex        =   450
                  Top             =   0
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% BC da operação própria:"
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
                  Index           =   127
                  Left            =   0
                  TabIndex        =   451
                  Top             =   45
                  Width           =   2265
               End
            End
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   315
               Index           =   18
               Left            =   75
               TabIndex        =   442
               Top             =   1230
               Visible         =   0   'False
               Width           =   8385
               Begin MSMask.MaskEdBox MaskEdBox15 
                  Height          =   285
                  Left            =   4515
                  TabIndex        =   443
                  ToolTipText     =   "Valor como se não tivesse o diferimento - vICMSOp"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox MaskEdBox16 
                  Height          =   285
                  Left            =   1725
                  TabIndex        =   444
                  ToolTipText     =   "Percentual do diferimento -  pDif"
                  Top             =   0
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox MaskEdBox17 
                  Height          =   285
                  Left            =   7335
                  TabIndex        =   445
                  ToolTipText     =   "Valor do ICMS diferido - vICMSDif"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS Diferido:"
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
                  Index           =   126
                  Left            =   6045
                  TabIndex        =   448
                  Top             =   45
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% do Diferimento:"
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
                  Index           =   125
                  Left            =   210
                  TabIndex        =   447
                  Top             =   15
                  Width           =   1500
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS da Operação:"
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
                  Index           =   124
                  Left            =   2835
                  TabIndex        =   446
                  Top             =   45
                  Width           =   1665
               End
            End
            Begin MSMask.MaskEdBox ICMSvBCFCPSTRetItem 
               Height          =   285
               Left            =   1800
               TabIndex        =   452
               Top             =   300
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSpFCPSTRetItem 
               Height          =   285
               Left            =   4095
               TabIndex        =   453
               Top             =   315
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSvFCPSTRetItem 
               Height          =   285
               Left            =   6885
               TabIndex        =   454
               Top             =   315
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   130
               Left            =   6360
               TabIndex        =   457
               Top             =   360
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   129
               Left            =   3270
               TabIndex        =   456
               Top             =   330
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base de Cálculo:"
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
               Index           =   128
               Left            =   315
               TabIndex        =   455
               Top             =   345
               Width           =   1440
            End
         End
         Begin VB.Frame FrameFCP 
            Caption         =   " FCP retido por Substituição Tributária"
            Height          =   750
            Index           =   1
            Left            =   0
            TabIndex        =   424
            Top             =   1110
            Width           =   8580
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   315
               Index           =   17
               Left            =   75
               TabIndex        =   428
               Top             =   1230
               Visible         =   0   'False
               Width           =   8385
               Begin MSMask.MaskEdBox MaskEdBox9 
                  Height          =   285
                  Left            =   4515
                  TabIndex        =   429
                  ToolTipText     =   "Valor como se não tivesse o diferimento - vICMSOp"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox MaskEdBox10 
                  Height          =   285
                  Left            =   1725
                  TabIndex        =   430
                  ToolTipText     =   "Percentual do diferimento -  pDif"
                  Top             =   0
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox MaskEdBox11 
                  Height          =   285
                  Left            =   7335
                  TabIndex        =   431
                  ToolTipText     =   "Valor do ICMS diferido - vICMSDif"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS da Operação:"
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
                  Index           =   120
                  Left            =   2835
                  TabIndex        =   434
                  Top             =   45
                  Width           =   1665
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% do Diferimento:"
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
                  Index           =   119
                  Left            =   210
                  TabIndex        =   433
                  Top             =   15
                  Width           =   1500
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS Diferido:"
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
                  Index           =   115
                  Left            =   6045
                  TabIndex        =   432
                  Top             =   45
                  Width           =   1245
               End
            End
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   315
               Index           =   15
               Left            =   2295
               TabIndex        =   425
               Top             =   1230
               Width           =   3450
               Begin MSMask.MaskEdBox MaskEdBox4 
                  Height          =   285
                  Left            =   2295
                  TabIndex        =   426
                  Top             =   0
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% BC da operação própria:"
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
                  Index           =   113
                  Left            =   0
                  TabIndex        =   427
                  Top             =   45
                  Width           =   2265
               End
            End
            Begin MSMask.MaskEdBox ICMSvBCFCPSTItem 
               Height          =   285
               Left            =   1800
               TabIndex        =   435
               Top             =   300
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSpFCPSTItem 
               Height          =   285
               Left            =   4095
               TabIndex        =   436
               Top             =   315
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSvFCPSTItem 
               Height          =   285
               Left            =   6885
               TabIndex        =   437
               Top             =   315
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base de Cálculo:"
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
               Index           =   123
               Left            =   315
               TabIndex        =   440
               Top             =   345
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   122
               Left            =   3270
               TabIndex        =   439
               Top             =   330
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   121
               Left            =   6360
               TabIndex        =   438
               Top             =   360
               Width           =   510
            End
         End
         Begin VB.Frame FrameFCP 
            Caption         =   "Fundo de Combate à Pobreza (FCP)"
            Height          =   750
            Index           =   0
            Left            =   0
            TabIndex        =   407
            Top             =   165
            Width           =   8580
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   315
               Index           =   16
               Left            =   2295
               TabIndex        =   418
               Top             =   1230
               Width           =   3450
               Begin MSMask.MaskEdBox MaskEdBox8 
                  Height          =   285
                  Left            =   2295
                  TabIndex        =   419
                  Top             =   0
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% BC da operação própria:"
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
                  Index           =   114
                  Left            =   0
                  TabIndex        =   420
                  Top             =   45
                  Width           =   2265
               End
            End
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   315
               Index           =   14
               Left            =   75
               TabIndex        =   408
               Top             =   1230
               Visible         =   0   'False
               Width           =   8385
               Begin MSMask.MaskEdBox MaskEdBox1 
                  Height          =   285
                  Left            =   4515
                  TabIndex        =   409
                  ToolTipText     =   "Valor como se não tivesse o diferimento - vICMSOp"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox MaskEdBox2 
                  Height          =   285
                  Left            =   1725
                  TabIndex        =   410
                  ToolTipText     =   "Percentual do diferimento -  pDif"
                  Top             =   0
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox MaskEdBox3 
                  Height          =   285
                  Left            =   7335
                  TabIndex        =   411
                  ToolTipText     =   "Valor do ICMS diferido - vICMSDif"
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS Diferido:"
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
                  Index           =   112
                  Left            =   6045
                  TabIndex        =   414
                  Top             =   45
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% do Diferimento:"
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
                  Index           =   111
                  Left            =   210
                  TabIndex        =   413
                  Top             =   15
                  Width           =   1500
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ICMS da Operação:"
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
                  Index           =   110
                  Left            =   2835
                  TabIndex        =   412
                  Top             =   45
                  Width           =   1665
               End
            End
            Begin MSMask.MaskEdBox ICMSvBCFCPItem 
               Height          =   285
               Left            =   1800
               TabIndex        =   415
               Top             =   300
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSpFCPItem 
               Height          =   285
               Left            =   4095
               TabIndex        =   416
               Top             =   315
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ICMSvFCPItem 
               Height          =   285
               Left            =   6885
               TabIndex        =   417
               Top             =   315
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   118
               Left            =   6360
               TabIndex        =   423
               Top             =   360
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   117
               Left            =   3270
               TabIndex        =   422
               Top             =   330
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base de Cálculo:"
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
               Index           =   116
               Left            =   315
               TabIndex        =   421
               Top             =   345
               Width           =   1440
            End
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3225
         Index           =   7
         Left            =   30
         TabIndex        =   301
         Top             =   870
         Visible         =   0   'False
         Width           =   8445
         Begin VB.Frame FrameImpostoImportacao 
            Caption         =   "Imposto de Importação"
            Height          =   2190
            Left            =   645
            TabIndex        =   303
            Top             =   135
            Width           =   5130
            Begin MSMask.MaskEdBox IIBaseItem 
               Height          =   285
               Left            =   2235
               TabIndex        =   304
               Top             =   345
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IIDespAduItem 
               Height          =   285
               Left            =   2235
               TabIndex        =   305
               Top             =   810
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IIIOFItem 
               Height          =   285
               Left            =   2235
               TabIndex        =   306
               Top             =   1275
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IIValorItem 
               Height          =   285
               Left            =   2235
               TabIndex        =   307
               Top             =   1755
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   58
               Left            =   1590
               TabIndex        =   319
               Top             =   1785
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor do IOF:"
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
               Index           =   57
               Left            =   945
               TabIndex        =   318
               Top             =   1305
               Width           =   1140
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Despesas Aduaneiras:"
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
               Index           =   56
               Left            =   195
               TabIndex        =   317
               Top             =   840
               Width           =   1905
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base de Cálculo:"
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
               Index           =   55
               Left            =   645
               TabIndex        =   316
               Top             =   375
               Width           =   1440
            End
         End
         Begin VB.TextBox FCI 
            Height          =   300
            Left            =   1140
            MaxLength       =   36
            TabIndex        =   302
            Top             =   2535
            Width           =   4050
         End
         Begin VB.Label Label1 
            Caption         =   "FCI:"
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
            Index           =   132
            Left            =   690
            TabIndex        =   320
            Top             =   2580
            Width           =   435
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3285
         Index           =   6
         Left            =   60
         TabIndex        =   263
         Top             =   855
         Visible         =   0   'False
         Width           =   8595
         Begin VB.Frame Frame1 
            Caption         =   "Invisível"
            Height          =   375
            Index           =   12
            Left            =   60
            TabIndex        =   364
            Top             =   1470
            Visible         =   0   'False
            Width           =   795
            Begin MSMask.MaskEdBox ISSValorDescIncondItem 
               Height          =   315
               Left            =   1245
               TabIndex        =   365
               Top             =   585
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ISSValorDescCondItem 
               Height          =   315
               Left            =   1770
               TabIndex        =   367
               Top             =   210
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Desc.Condicionado:"
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
               Index           =   34
               Left            =   0
               TabIndex        =   368
               Top             =   255
               Width           =   1725
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Desc.Incond.:"
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
               Index           =   32
               Left            =   0
               TabIndex        =   366
               Top             =   630
               Width           =   1215
            End
         End
         Begin VB.ComboBox ISSCodPaisItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6600
            TabIndex        =   113
            Top             =   750
            Width           =   1935
         End
         Begin VB.CheckBox ISSIndIncentivoItem 
            Caption         =   "Incentivo Fiscal"
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
            Left            =   6600
            TabIndex        =   351
            Top             =   2985
            Width           =   2085
         End
         Begin VB.ComboBox ISSIndExigibilidadeItem 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "TabTributacaoFat.ctx":024D
            Left            =   1890
            List            =   "TabTributacaoFat.ctx":0269
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   1830
            Width           =   6630
         End
         Begin VB.ComboBox ISSListaServ 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1905
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   1110
            Width           =   6630
         End
         Begin VB.ComboBox ComboISSTipo 
            Height          =   315
            Left            =   1905
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   30
            Width           =   3930
         End
         Begin MSMask.MaskEdBox ISSBaseItem 
            Height          =   315
            Left            =   1905
            TabIndex        =   109
            Top             =   390
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSAliquotaItem 
            Height          =   315
            Left            =   4710
            TabIndex        =   110
            Top             =   390
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSCodIBGE 
            Height          =   315
            Left            =   1905
            TabIndex        =   112
            Top             =   750
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   7
            Mask            =   "#######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSValorItem 
            Height          =   315
            Left            =   7410
            TabIndex        =   111
            Top             =   390
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSValorDeducaoItem 
            Height          =   315
            Left            =   1905
            TabIndex        =   115
            Top             =   1470
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSValorOutrasRetItem 
            Height          =   315
            Left            =   4710
            TabIndex        =   116
            Top             =   1470
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSValorRetItem 
            Height          =   315
            Left            =   7425
            TabIndex        =   117
            Top             =   1470
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSMunicIncidImpItem 
            Height          =   315
            Left            =   1890
            TabIndex        =   120
            Top             =   2565
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   7
            Mask            =   "#######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSNumProcessoItem 
            Height          =   315
            Left            =   1890
            TabIndex        =   121
            Top             =   2940
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   30
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSCodServItem 
            Height          =   315
            Left            =   1890
            TabIndex        =   119
            Top             =   2190
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSMunicFaturamento 
            Height          =   315
            Left            =   6705
            TabIndex        =   369
            Top             =   0
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   7
            Mask            =   "#######"
            PromptChar      =   " "
         End
         Begin VB.Label ISSCodServItemLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cód.Serviço Munic.:"
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
            Left            =   90
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   363
            Top             =   2235
            Width           =   1725
         End
         Begin VB.Label ISSCodServDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3510
            TabIndex        =   362
            Top             =   2190
            Width           =   5010
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Núm.Processo.:"
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
            Index           =   38
            Left            =   465
            TabIndex        =   355
            Top             =   2985
            Width           =   1335
         End
         Begin VB.Label ISSPaisLabel 
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
            Left            =   6105
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   354
            Top             =   780
            Width           =   495
         End
         Begin VB.Label ISSMunicIncidDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2655
            TabIndex        =   353
            Top             =   2565
            Width           =   5850
         End
         Begin VB.Label ISSLabelMunicIncidImp 
            AutoSize        =   -1  'True
            Caption         =   "Munic. Imposto:"
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
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   352
            Top             =   2625
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ind.Exigibilidade:"
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
            Index           =   44
            Left            =   360
            TabIndex        =   350
            Top             =   1890
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Retenção:"
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
            Index           =   35
            Left            =   6495
            TabIndex        =   349
            Top             =   1500
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Outras Retenções:"
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
            Index           =   31
            Left            =   3075
            TabIndex        =   348
            Top             =   1515
            Width           =   1590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Dedução:"
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
            Index           =   28
            Left            =   1020
            TabIndex        =   347
            Top             =   1515
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota (%):"
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
            Index           =   51
            Left            =   3555
            TabIndex        =   270
            Top             =   435
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base de Cálculo:"
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
            Index           =   52
            Left            =   405
            TabIndex        =   269
            Top             =   435
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Lista de Serviço:"
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
            Index           =   49
            Left            =   405
            TabIndex        =   268
            Top             =   1140
            Width           =   1440
         End
         Begin VB.Label ISSDescIBGE 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2670
            TabIndex        =   267
            Top             =   750
            Width           =   3150
         End
         Begin VB.Label ISSLabelCodIBGE 
            AutoSize        =   -1  'True
            Caption         =   "Munic. Fato Gerador:"
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
            Left            =   30
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   266
            Top             =   795
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Index           =   53
            Left            =   6870
            TabIndex        =   265
            Top             =   435
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Situação Tributária:"
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
            Index           =   99
            Left            =   165
            TabIndex        =   264
            Top             =   75
            Width           =   1680
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3225
         Index           =   4
         Left            =   30
         TabIndex        =   321
         Top             =   870
         Visible         =   0   'False
         Width           =   8445
         Begin VB.Frame Frame1 
            Caption         =   "PIS"
            Height          =   1770
            Index           =   3
            Left            =   75
            TabIndex        =   331
            Top             =   -30
            Width           =   8355
            Begin VB.ComboBox ComboPISTipo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1995
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   135
               Width           =   6225
            End
            Begin VB.Frame FramePISTipoCalculoBase 
               BorderStyle     =   0  'None
               Height          =   390
               Left            =   330
               TabIndex        =   332
               Top             =   720
               Width           =   7125
               Begin MSMask.MaskEdBox PISBaseItem 
                  Height          =   285
                  Left            =   1665
                  TabIndex        =   84
                  Top             =   90
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox PISAliquotaItem 
                  Height          =   285
                  Left            =   4905
                  TabIndex        =   85
                  Top             =   90
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (%):"
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
                  Index           =   11
                  Left            =   3690
                  TabIndex        =   334
                  Top             =   135
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base de Cálculo:"
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
                  Index           =   10
                  Left            =   90
                  TabIndex        =   333
                  Top             =   120
                  Width           =   1440
               End
            End
            Begin VB.ComboBox PISTipoCalculo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1995
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   465
               Width           =   4350
            End
            Begin MSMask.MaskEdBox PISValorItem 
               Height          =   285
               Left            =   1995
               TabIndex        =   88
               Top             =   1440
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Frame FramePISTipoCalculoQtd 
               BorderStyle     =   0  'None
               Height          =   420
               Left            =   480
               TabIndex        =   335
               Top             =   990
               Width           =   6960
               Begin MSMask.MaskEdBox PISAliquotaRSItem 
                  Height          =   285
                  Left            =   1515
                  TabIndex        =   86
                  Top             =   135
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox PISQtdeItem 
                  Height          =   285
                  Left            =   4755
                  TabIndex        =   87
                  Top             =   135
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Quantidade Vendida:"
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
                  Index           =   14
                  Left            =   2850
                  TabIndex        =   337
                  Top             =   165
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (R$):"
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
                  Index           =   13
                  Left            =   150
                  TabIndex        =   336
                  Top             =   165
                  Width           =   1200
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Situação Tributária:"
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
               Index           =   104
               Left            =   195
               TabIndex        =   340
               Top             =   195
               Width           =   1680
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   12
               Left            =   1320
               TabIndex        =   339
               Top             =   1470
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Cálculo:"
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
               Index           =   94
               Left            =   465
               TabIndex        =   338
               Top             =   525
               Width           =   1395
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "PIS Substituição Tributária"
            Height          =   1485
            Index           =   4
            Left            =   75
            TabIndex        =   322
            Top             =   1725
            Width           =   8355
            Begin VB.ComboBox PISSTTipoCalculo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1965
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   225
               Width           =   4350
            End
            Begin VB.Frame FramePISSTTipoCalculoBase 
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   330
               TabIndex        =   326
               Top             =   555
               Width           =   7275
               Begin MSMask.MaskEdBox PISSTBaseItem 
                  Height          =   285
                  Left            =   1635
                  TabIndex        =   90
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox PISSTAliquotaItem 
                  Height          =   285
                  Left            =   4875
                  TabIndex        =   91
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (%):"
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
                  Index           =   17
                  Left            =   3660
                  TabIndex        =   328
                  Top             =   15
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base de Cálculo:"
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
                  Index           =   15
                  Left            =   60
                  TabIndex        =   327
                  Top             =   30
                  Width           =   1440
               End
            End
            Begin VB.Frame FramePISSTTipoCalculoQtd 
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   450
               TabIndex        =   323
               Top             =   840
               Width           =   7185
               Begin MSMask.MaskEdBox PISSTAliquotaRSItem 
                  Height          =   285
                  Left            =   1515
                  TabIndex        =   92
                  Top             =   15
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox PISSTQtdeItem 
                  Height          =   285
                  Left            =   4755
                  TabIndex        =   93
                  Top             =   15
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Quantidade Vendida:"
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
                  Index           =   20
                  Left            =   2850
                  TabIndex        =   325
                  Top             =   45
                  Width           =   1800
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (R$):"
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
                  Index           =   19
                  Left            =   150
                  TabIndex        =   324
                  Top             =   45
                  Width           =   1200
               End
            End
            Begin MSMask.MaskEdBox PISSTValorItem 
               Height          =   285
               Left            =   1965
               TabIndex        =   94
               Top             =   1155
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   18
               Left            =   1290
               TabIndex        =   330
               Top             =   1185
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Cálculo:"
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
               Index           =   95
               Left            =   435
               TabIndex        =   329
               Top             =   285
               Width           =   1395
            End
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3270
         Index           =   3
         Left            =   45
         TabIndex        =   203
         Top             =   855
         Visible         =   0   'False
         Width           =   8595
         Begin VB.ComboBox ComboIPITipo 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   105
            Width           =   6120
         End
         Begin VB.Frame FrameIPIItem 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   1455
            Left            =   75
            TabIndex        =   210
            Top             =   495
            Width           =   8415
            Begin VB.ComboBox IPITipoCalculo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2400
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   0
               Width           =   4035
            End
            Begin VB.Frame FrameIPITipoCalculoBase 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   840
               TabIndex        =   214
               Top             =   375
               Width           =   7515
               Begin MSMask.MaskEdBox IPIBaseItem 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   70
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIPercRedBaseItem 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   72
                  Top             =   15
                  Visible         =   0   'False
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIAliquotaItem 
                  Height          =   315
                  Left            =   4680
                  TabIndex        =   71
                  Top             =   15
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
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
                  Height          =   195
                  Index           =   209
                  Left            =   3855
                  TabIndex        =   217
                  Top             =   60
                  Width           =   780
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "% de Redução na BC:"
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
                  Index           =   171
                  Left            =   5835
                  TabIndex        =   216
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   1860
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base de Cálculo:"
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
                  Index           =   43
                  Left            =   90
                  TabIndex        =   215
                  Top             =   60
                  Width           =   1440
               End
            End
            Begin VB.Frame FrameIPITipoCalculoQtd 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   240
               TabIndex        =   211
               Top             =   705
               Width           =   8100
               Begin MSMask.MaskEdBox IPIUnidadePadraoQtdeItem 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   73
                  Top             =   45
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIUnidadePadraoValorItem 
                  Height          =   315
                  Left            =   5280
                  TabIndex        =   74
                  Top             =   45
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor Unitário:"
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
                  Index           =   62
                  Left            =   4020
                  TabIndex        =   213
                  Top             =   105
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Quantidade UM Padrão:"
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
                  Index           =   61
                  Left            =   90
                  TabIndex        =   212
                  Top             =   90
                  Width           =   2040
               End
            End
            Begin VB.CheckBox IPICredita 
               Caption         =   "Debita"
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
               Height          =   264
               Left            =   3645
               TabIndex        =   76
               Top             =   1140
               Width           =   936
            End
            Begin MSMask.MaskEdBox IPIValorItem 
               Height          =   315
               Left            =   2400
               TabIndex        =   75
               Top             =   1125
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   40
               Left            =   1875
               TabIndex        =   219
               Top             =   1185
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Cálculo:"
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
               Index           =   106
               Left            =   975
               TabIndex        =   218
               Top             =   45
               Width           =   1395
            End
         End
         Begin VB.Frame FrameIPIItem2 
            BorderStyle     =   0  'None
            Height          =   1425
            Left            =   -15
            TabIndex        =   204
            Top             =   1950
            Width           =   8580
            Begin MSMask.MaskEdBox IPIClasseEnq 
               Height          =   315
               Left            =   2490
               TabIndex        =   77
               Top             =   45
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   5
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IPICodEnq 
               Height          =   315
               Left            =   2505
               TabIndex        =   79
               Top             =   405
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   3
               Format          =   "000"
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IPICNPJProdutor 
               Height          =   315
               Left            =   5595
               TabIndex        =   78
               Top             =   45
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "##############"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IPICodSelo 
               Height          =   315
               Left            =   2490
               TabIndex        =   80
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IPISeloQtde 
               Height          =   315
               Left            =   5625
               TabIndex        =   81
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   5
               Format          =   "#,###"
               Mask            =   "#####"
               PromptChar      =   " "
            End
            Begin VB.Label IPICodEnqDesc 
               BorderStyle     =   1  'Fixed Single
               Height          =   525
               Left            =   3240
               TabIndex        =   398
               Top             =   405
               Width           =   5340
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Classe de enquadramento:"
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
               Left            =   195
               TabIndex        =   209
               Top             =   90
               Width           =   2265
            End
            Begin VB.Label IPICodEnqLabel 
               AutoSize        =   -1  'True
               Caption         =   "Cód. de enquadramento:"
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
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   208
               Top             =   450
               Width           =   2085
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CNPJ do Produtor:"
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
               Index           =   105
               Left            =   3960
               TabIndex        =   207
               Top             =   120
               Width           =   1590
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Código do selo de controle:"
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
               Index           =   4
               Left            =   90
               TabIndex        =   206
               Top             =   1005
               Width           =   2340
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Qtd do selo de  controle:"
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
               Index           =   5
               Left            =   3495
               TabIndex        =   205
               Top             =   990
               Width           =   2130
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Situação Tributária:"
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
            Index           =   109
            Left            =   750
            TabIndex        =   220
            Top             =   150
            Width           =   1680
         End
      End
      Begin VB.Frame FrameTabTribDet 
         BorderStyle     =   0  'None
         Height          =   3225
         Index           =   5
         Left            =   90
         TabIndex        =   271
         Top             =   870
         Visible         =   0   'False
         Width           =   8445
         Begin VB.Frame Frame1 
            Caption         =   "COFINS"
            Height          =   1770
            Index           =   5
            Left            =   15
            TabIndex        =   281
            Top             =   -30
            Width           =   8355
            Begin VB.Frame FrameCOFINSTipoCalculoBase 
               BorderStyle     =   0  'None
               Height          =   390
               Left            =   330
               TabIndex        =   282
               Top             =   720
               Width           =   7125
               Begin MSMask.MaskEdBox COFINSBaseItem 
                  Height          =   285
                  Left            =   1665
                  TabIndex        =   97
                  Top             =   90
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox COFINSAliquotaItem 
                  Height          =   285
                  Left            =   4905
                  TabIndex        =   98
                  Top             =   90
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base de Cálculo:"
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
                  Index           =   23
                  Left            =   90
                  TabIndex        =   284
                  Top             =   120
                  Width           =   1440
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (%):"
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
                  Index           =   24
                  Left            =   3690
                  TabIndex        =   283
                  Top             =   135
                  Width           =   1095
               End
            End
            Begin VB.Frame FrameCOFINSTipoCalculoQtd 
               BorderStyle     =   0  'None
               Height          =   420
               Left            =   480
               TabIndex        =   285
               Top             =   990
               Width           =   6960
               Begin MSMask.MaskEdBox COFINSAliquotaRSItem 
                  Height          =   285
                  Left            =   1515
                  TabIndex        =   99
                  Top             =   135
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox COFINSQtdeItem 
                  Height          =   285
                  Left            =   4755
                  TabIndex        =   100
                  Top             =   135
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (R$):"
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
                  Index           =   25
                  Left            =   150
                  TabIndex        =   287
                  Top             =   165
                  Width           =   1200
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Quantidade Vendida:"
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
                  Index           =   26
                  Left            =   2850
                  TabIndex        =   286
                  Top             =   165
                  Width           =   1800
               End
            End
            Begin VB.ComboBox ComboCOFINSTipo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1995
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   135
               Width           =   6225
            End
            Begin VB.ComboBox COFINSTipoCalculo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1995
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   465
               Width           =   4350
            End
            Begin MSMask.MaskEdBox COFINSValorItem 
               Height          =   285
               Left            =   1995
               TabIndex        =   101
               Top             =   1440
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   36
               Left            =   1320
               TabIndex        =   290
               Top             =   1470
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Situação Tributária:"
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
               Index           =   97
               Left            =   195
               TabIndex        =   289
               Top             =   195
               Width           =   1680
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Cálculo:"
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
               Index           =   96
               Left            =   465
               TabIndex        =   288
               Top             =   525
               Width           =   1395
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "COFINS Substituição Tributária"
            Height          =   1485
            Index           =   6
            Left            =   15
            TabIndex        =   272
            Top             =   1725
            Width           =   8355
            Begin VB.Frame FrameCOFINSSTTipoCalculoBase 
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   330
               TabIndex        =   276
               Top             =   555
               Width           =   7275
               Begin MSMask.MaskEdBox COFINSSTBaseItem 
                  Height          =   285
                  Left            =   1635
                  TabIndex        =   103
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox COFINSSTAliquotaItem 
                  Height          =   285
                  Left            =   4875
                  TabIndex        =   104
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base de Cálculo:"
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
                  Index           =   41
                  Left            =   60
                  TabIndex        =   278
                  Top             =   30
                  Width           =   1440
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (%):"
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
                  Index           =   45
                  Left            =   3660
                  TabIndex        =   277
                  Top             =   15
                  Width           =   1095
               End
            End
            Begin VB.ComboBox COFINSSTTipoCalculo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1965
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   225
               Width           =   4350
            End
            Begin VB.Frame FrameCOFINSSTTipoCalculoQtd 
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   450
               TabIndex        =   273
               Top             =   840
               Width           =   7185
               Begin MSMask.MaskEdBox COFINSSTAliquotaRSItem 
                  Height          =   285
                  Left            =   1515
                  TabIndex        =   105
                  Top             =   15
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox COFINSSTQtdeItem 
                  Height          =   285
                  Left            =   4755
                  TabIndex        =   106
                  Top             =   15
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00##"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Alíquota (R$):"
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
                  Index           =   37
                  Left            =   150
                  TabIndex        =   275
                  Top             =   45
                  Width           =   1200
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Quantidade Vendida:"
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
                  Index           =   39
                  Left            =   2850
                  TabIndex        =   274
                  Top             =   45
                  Width           =   1800
               End
            End
            Begin MSMask.MaskEdBox COFINSSTValorItem 
               Height          =   285
               Left            =   1965
               TabIndex        =   107
               Top             =   1155
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Cálculo:"
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
               Index           =   98
               Left            =   435
               TabIndex        =   280
               Top             =   285
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   46
               Left            =   1290
               TabIndex        =   279
               Top             =   1185
               Width           =   510
            End
         End
      End
      Begin VB.Frame FrameItensTrib 
         Caption         =   "Item"
         Height          =   540
         Left            =   0
         TabIndex        =   341
         Top             =   -15
         Width           =   8700
         Begin VB.ComboBox ComboItensTrib 
            Height          =   315
            Left            =   45
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   180
            Width           =   3990
         End
         Begin VB.Label Label1 
            Caption         =   "Quantidade:"
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
            Index           =   9
            Left            =   5880
            TabIndex        =   346
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Valor:"
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
            Index           =   191
            Left            =   4125
            TabIndex        =   345
            Top             =   225
            Width           =   570
         End
         Begin VB.Label LabelValorItem 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4680
            TabIndex        =   344
            Top             =   165
            Width           =   1140
         End
         Begin VB.Label LabelQtdeItem 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6975
            TabIndex        =   343
            Top             =   165
            Width           =   945
         End
         Begin VB.Label LabelUMItem 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7905
            TabIndex        =   342
            Top             =   165
            Width           =   765
         End
      End
      Begin MSComctlLib.TabStrip OpcaoTribDet 
         Height          =   3660
         Left            =   0
         TabIndex        =   37
         Top             =   525
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   6456
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   10
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Geral"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ICMS"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "IPI"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PIS"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "COFINS"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ISS"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "II"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ICMS Inter."
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "FCP"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Trib.Devolvidos"
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
   Begin VB.Frame FrameTributacao 
      BorderStyle     =   0  'None
      Height          =   4215
      Index           =   1
      Left            =   60
      TabIndex        =   123
      Top             =   315
      Width           =   8745
      Begin VB.Frame FrameResumo 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3825
         Index           =   1
         Left            =   120
         TabIndex        =   154
         Top             =   345
         Width           =   8580
         Begin VB.Frame Frame1 
            Caption         =   "ICMS"
            Height          =   1695
            Index           =   35
            Left            =   15
            TabIndex        =   168
            Top             =   1215
            Width           =   6225
            Begin VB.Frame Frame1 
               Caption         =   "Substituição Tributária"
               Height          =   1470
               Index           =   36
               Left            =   4305
               TabIndex        =   169
               Top             =   120
               Width           =   1830
               Begin VB.Label ICMSVlrFCPST 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   645
                  TabIndex        =   400
                  Top             =   1005
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FCP:"
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
                  Index           =   3
                  Left            =   150
                  TabIndex        =   399
                  Top             =   1050
                  Width           =   420
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor:"
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
                  Index           =   195
                  Left            =   90
                  TabIndex        =   173
                  Top             =   690
                  Width           =   510
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
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
                  Height          =   195
                  Index           =   217
                  Left            =   105
                  TabIndex        =   172
                  Top             =   330
                  Width           =   495
               End
               Begin VB.Label ICMSSubstBase 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   660
                  TabIndex        =   171
                  Top             =   285
                  Width           =   1080
               End
               Begin VB.Label ICMSSubstValor 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   660
                  TabIndex        =   170
                  Top             =   645
                  Width           =   1080
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "FCP Ret.Ant.:"
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
               Index           =   93
               Left            =   1905
               TabIndex        =   404
               Top             =   1185
               Width           =   1185
            End
            Begin VB.Label ICMSVlrFCPSTRet 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3120
               TabIndex        =   403
               Top             =   1140
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "FCP:"
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
               Index           =   92
               Left            =   2670
               TabIndex        =   402
               Top             =   795
               Width           =   420
            End
            Begin VB.Label ICMSVlrFCP 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3120
               TabIndex        =   401
               Top             =   765
               Width           =   1080
            End
            Begin VB.Label ICMSDesonerado 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   3120
               TabIndex        =   181
               Top             =   405
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Desonerado:"
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
               Index           =   215
               Left            =   1995
               TabIndex        =   180
               Top             =   435
               Width           =   1095
            End
            Begin VB.Label LabelICMSCredito 
               AutoSize        =   -1  'True
               Caption         =   "Crédito:"
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
               Left            =   60
               TabIndex        =   179
               Top             =   1155
               Width           =   660
            End
            Begin VB.Label ICMSCredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   795
               TabIndex        =   178
               Top             =   1125
               Width           =   1080
            End
            Begin VB.Label ICMSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   795
               TabIndex        =   177
               Top             =   765
               Width           =   1080
            End
            Begin VB.Label ICMSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   795
               TabIndex        =   176
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   203
               Left            =   240
               TabIndex        =   175
               Top             =   420
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   204
               Left            =   225
               TabIndex        =   174
               Top             =   795
               Width           =   510
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Totais de Tributos"
            Height          =   810
            Index           =   39
            Left            =   15
            TabIndex        =   166
            Top             =   2940
            Width           =   8490
            Begin VB.ComboBox TotTribTipo 
               Height          =   315
               ItemData        =   "TabTributacaoFat.ctx":031B
               Left            =   90
               List            =   "TabTributacaoFat.ctx":0328
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   420
               Width           =   2385
            End
            Begin MSMask.MaskEdBox TotTrib 
               Height          =   315
               Left            =   2700
               TabIndex        =   5
               ToolTipText     =   "Soma dos valores dos tributos para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "###,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TotTribFed 
               Height          =   315
               Left            =   4185
               TabIndex        =   6
               ToolTipText     =   "Soma dos valores dos tributos para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "###,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TotTribEst 
               Height          =   315
               Left            =   5685
               TabIndex        =   7
               ToolTipText     =   "Soma dos valores dos tributos para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "###,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TotTribMunic 
               Height          =   315
               Left            =   7170
               TabIndex        =   8
               ToolTipText     =   "Soma dos valores dos tributos para atender à Lei da Transparencia."
               Top             =   420
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "###,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Municipal (R$)"
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
               Index           =   75
               Left            =   7155
               TabIndex        =   374
               Top             =   195
               Width           =   1245
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Estadual (R$)"
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
               Index           =   54
               Left            =   5670
               TabIndex        =   373
               Top             =   195
               Width           =   1170
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Federal (R$)"
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
               Index           =   50
               Left            =   4185
               TabIndex        =   372
               Top             =   195
               Width           =   1065
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total (R$)"
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
               Index           =   48
               Left            =   2670
               TabIndex        =   371
               Top             =   195
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fonte"
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
               Index           =   29
               Left            =   90
               TabIndex        =   370
               Top             =   195
               Width           =   495
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "IPI"
            Height          =   1710
            Index           =   38
            Left            =   6300
            TabIndex        =   159
            Top             =   1215
            Width           =   2205
            Begin VB.Label IPIVlrDevolvido 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1050
               TabIndex        =   465
               Top             =   1260
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Devolvido:"
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
               Index           =   101
               Left            =   105
               TabIndex        =   405
               Top             =   1335
               Width           =   930
            End
            Begin VB.Label LabelIPICredito 
               AutoSize        =   -1  'True
               Caption         =   "Crédito:"
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
               Left            =   345
               TabIndex        =   165
               Top             =   960
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   207
               Left            =   510
               TabIndex        =   164
               Top             =   585
               Width           =   510
            End
            Begin VB.Label IPIValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1050
               TabIndex        =   163
               Top             =   555
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   211
               Left            =   525
               TabIndex        =   162
               Top             =   225
               Width           =   495
            End
            Begin VB.Label IPIBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1050
               TabIndex        =   161
               Top             =   180
               Width           =   1080
            End
            Begin VB.Label IPICredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1050
               TabIndex        =   160
               Top             =   915
               Width           =   1080
            End
         End
         Begin VB.CheckBox IndConsumidorFinal 
            Caption         =   "Consumidor Final"
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
            Left            =   6705
            TabIndex        =   3
            ToolTipText     =   "Indicador de consumidor final"
            Top             =   930
            Width           =   1830
         End
         Begin VB.ComboBox IndPresenca 
            Height          =   315
            ItemData        =   "TabTributacaoFat.ctx":0362
            Left            =   1605
            List            =   "TabTributacaoFat.ctx":037B
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Indicador de Presença"
            Top             =   900
            Width           =   4995
         End
         Begin VB.CommandButton TributacaoRecalcular 
            Caption         =   "Recalcular Tributação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   7350
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   75
            Width           =   1125
         End
         Begin MSMask.MaskEdBox TipoTributacao 
            Height          =   315
            Left            =   1065
            TabIndex        =   1
            Top             =   510
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NatOpInterna 
            Height          =   315
            Left            =   1065
            TabIndex        =   0
            Top             =   120
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ind. Presença:"
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
            Index           =   220
            Left            =   285
            TabIndex        =   167
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label LblNatOpInterna 
            AutoSize        =   -1  'True
            Caption         =   "CFOP:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   158
            Top             =   180
            Width           =   555
         End
         Begin VB.Label DescNatOpInterna 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1605
            TabIndex        =   157
            Top             =   120
            Width           =   5730
         End
         Begin VB.Label LblTipoTrib 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Trib.:"
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
            Left            =   105
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   156
            Top             =   570
            Width           =   900
         End
         Begin VB.Label DescTipoTrib 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1605
            TabIndex        =   155
            Top             =   510
            Width           =   5730
         End
      End
      Begin VB.Frame FrameResumo 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3780
         Index           =   2
         Left            =   105
         TabIndex        =   124
         Top             =   300
         Visible         =   0   'False
         Width           =   8505
         Begin VB.Frame Frame1 
            Caption         =   "ISS"
            Height          =   2250
            Index           =   29
            Left            =   225
            TabIndex        =   133
            Top             =   30
            Width           =   8100
            Begin VB.CheckBox ISSIncluso 
               BackColor       =   &H00C0C0FF&
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
               Height          =   192
               Left            =   8340
               TabIndex        =   134
               Top             =   2055
               Value           =   1  'Checked
               Width           =   195
            End
            Begin MSMask.MaskEdBox ISSRetido 
               Height          =   315
               Left            =   6735
               TabIndex        =   13
               Top             =   1770
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownPrestServico 
               Height          =   270
               Left            =   3495
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   1785
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   476
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataPrestServico 
               Height          =   315
               Left            =   2460
               TabIndex        =   11
               ToolTipText     =   "Data de prestação do serviço"
               Top             =   1785
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Deduções:"
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
               Index           =   168
               Left            =   1485
               TabIndex        =   148
               Top             =   705
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   167
               Left            =   1920
               TabIndex        =   147
               Top             =   225
               Width           =   495
            End
            Begin VB.Label ISSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2460
               TabIndex        =   146
               Top             =   180
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Prestação do  Serviço:"
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
               Index           =   165
               Left            =   465
               TabIndex        =   145
               Top             =   1830
               Width           =   1950
            End
            Begin VB.Label ISSValorDeducao 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   2460
               TabIndex        =   144
               Top             =   675
               Width           =   1080
            End
            Begin VB.Label ISSValorOutrasRet 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   6720
               TabIndex        =   143
               Top             =   675
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Outras Retenções:"
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
               Index           =   162
               Left            =   5070
               TabIndex        =   142
               Top             =   705
               Width           =   1590
            End
            Begin VB.Label ISSValorDescIncond 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2475
               TabIndex        =   141
               Top             =   1230
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Desconto Incondicionado:"
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
               Index           =   160
               Left            =   165
               TabIndex        =   140
               Top             =   1275
               Width           =   2250
            End
            Begin VB.Label ISSValorDescCond 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6735
               TabIndex        =   139
               Top             =   1215
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Desconto Condicionado:"
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
               Index           =   158
               Left            =   4590
               TabIndex        =   138
               Top             =   1245
               Width           =   2100
            End
            Begin VB.Label ISSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6720
               TabIndex        =   137
               Top             =   180
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   156
               Left            =   6165
               TabIndex        =   136
               Top             =   210
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor Retido:"
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
               Index           =   154
               Left            =   5565
               TabIndex        =   135
               Top             =   1800
               Width           =   1125
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "INSS"
            Height          =   660
            Index           =   28
            Left            =   225
            TabIndex        =   129
            Top             =   2325
            Width           =   8100
            Begin VB.CheckBox INSSRetido 
               Caption         =   "Retido"
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
               Left            =   6735
               TabIndex        =   17
               Top             =   255
               Width           =   930
            End
            Begin MSMask.MaskEdBox INSSValor 
               Height          =   315
               Left            =   5040
               TabIndex        =   16
               Top             =   225
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSBase 
               Height          =   315
               Left            =   870
               TabIndex        =   14
               Top             =   210
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSDeducoes 
               Height          =   315
               Left            =   3105
               TabIndex        =   15
               Top             =   210
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Deduções:"
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
               Index           =   153
               Left            =   2145
               TabIndex        =   132
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   152
               Left            =   300
               TabIndex        =   131
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   151
               Left            =   4485
               TabIndex        =   130
               Top             =   255
               Width           =   510
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "IR"
            Height          =   585
            Index           =   27
            Left            =   240
            TabIndex        =   125
            Top             =   3060
            Width           =   8100
            Begin MSMask.MaskEdBox IRAliquota 
               Height          =   315
               Left            =   3135
               TabIndex        =   19
               Top             =   195
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   315
               Left            =   5055
               TabIndex        =   20
               Top             =   195
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox IRBase 
               Height          =   315
               Left            =   870
               TabIndex        =   18
               Top             =   195
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   149
               Left            =   4515
               TabIndex        =   128
               Top             =   225
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%:"
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
               Index           =   148
               Left            =   2910
               TabIndex        =   127
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   147
               Left            =   345
               TabIndex        =   126
               Top             =   225
               Width           =   495
            End
         End
      End
      Begin VB.Frame FrameResumo 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3780
         Index           =   3
         Left            =   120
         TabIndex        =   149
         Top             =   360
         Visible         =   0   'False
         Width           =   8580
         Begin VB.Frame Frame1 
            Caption         =   "ICMS em Operações Interestaduais"
            Height          =   1575
            Index           =   2
            Left            =   5175
            TabIndex        =   391
            Top             =   1875
            Width           =   3135
            Begin VB.Label ICMSInterestVlrUFRemet 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1110
               TabIndex        =   397
               Top             =   1065
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Remetente:"
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
               Index           =   91
               Left            =   60
               TabIndex        =   396
               Top             =   1095
               Width           =   990
            End
            Begin VB.Label ICMSInterestVlrUFDest 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1110
               TabIndex        =   395
               Top             =   690
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   88
               Left            =   330
               TabIndex        =   394
               Top             =   720
               Width           =   720
            End
            Begin VB.Label ICMSInterestVlrFCPUFDest 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1110
               TabIndex        =   393
               Top             =   300
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "FCP:"
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
               Index           =   87
               Left            =   630
               TabIndex        =   392
               Top             =   330
               Width           =   420
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "COFINS"
            Height          =   1575
            Index           =   0
            Left            =   210
            TabIndex        =   193
            Top             =   1875
            Width           =   4830
            Begin VB.Frame Frame1 
               Caption         =   "Substituição Tributária"
               Height          =   795
               Index           =   1
               Left            =   2415
               TabIndex        =   194
               Top             =   225
               Width           =   2010
               Begin VB.Label COFINSValorST 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   660
                  TabIndex        =   196
                  Top             =   345
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor:"
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
                  Index           =   6
                  Left            =   150
                  TabIndex        =   195
                  Top             =   390
                  Width           =   510
               End
            End
            Begin MSMask.MaskEdBox COFINSRetido 
               Height          =   315
               Left            =   855
               TabIndex        =   22
               Top             =   1080
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCOFINSCredito 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Débito:"
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
               TabIndex        =   201
               Top             =   735
               Width           =   720
            End
            Begin VB.Label COFINSCredito 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   870
               TabIndex        =   200
               Top             =   690
               Width           =   1080
            End
            Begin VB.Label COFINSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   885
               TabIndex        =   199
               Top             =   270
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   22
               Left            =   165
               TabIndex        =   198
               Top             =   1125
               Width           =   630
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   21
               Left            =   315
               TabIndex        =   197
               Top             =   300
               Width           =   510
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "PIS"
            Height          =   1575
            Index           =   30
            Left            =   210
            TabIndex        =   184
            Top             =   210
            Width           =   4830
            Begin MSMask.MaskEdBox PISRetido 
               Height          =   315
               Left            =   855
               TabIndex        =   21
               Top             =   1080
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Frame Frame1 
               Caption         =   "Substituição Tributária"
               Height          =   720
               Index           =   31
               Left            =   2400
               TabIndex        =   185
               Top             =   240
               Width           =   2010
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor:"
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
                  Index           =   172
                  Left            =   225
                  TabIndex        =   187
                  Top             =   330
                  Width           =   510
               End
               Begin VB.Label PISValorST 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   735
                  TabIndex        =   186
                  Top             =   285
                  Width           =   1080
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   177
               Left            =   270
               TabIndex        =   192
               Top             =   345
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   176
               Left            =   165
               TabIndex        =   191
               Top             =   1125
               Width           =   630
            End
            Begin VB.Label PISValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   840
               TabIndex        =   190
               Top             =   315
               Width           =   1080
            End
            Begin VB.Label PISCredito 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   840
               TabIndex        =   189
               Top             =   705
               Width           =   1080
            End
            Begin VB.Label LabelPISCredito 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Débito:"
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
               Left            =   90
               TabIndex        =   188
               Top             =   750
               Width           =   675
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "CSLL"
            Height          =   705
            Index           =   37
            Left            =   5190
            TabIndex        =   182
            Top             =   195
            Width           =   3135
            Begin MSMask.MaskEdBox CSLLRetido 
               Height          =   315
               Left            =   1185
               TabIndex        =   23
               Top             =   240
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   205
               Left            =   480
               TabIndex        =   183
               Top             =   285
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "II"
            Height          =   825
            Index           =   34
            Left            =   5175
            TabIndex        =   150
            Top             =   960
            Width           =   3135
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   219
               Left            =   465
               TabIndex        =   152
               Top             =   300
               Width           =   510
            End
            Begin VB.Label IIValor 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1110
               TabIndex        =   151
               Top             =   270
               Width           =   1080
            End
         End
      End
      Begin MSComctlLib.TabStrip OpcaoResumo 
         Height          =   4200
         Left            =   30
         TabIndex        =   10
         Top             =   0
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7408
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ICMS"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ISS"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Outros"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Presença:"
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
         Index           =   214
         Left            =   5040
         TabIndex        =   153
         Top             =   705
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4590
      Index           =   7
      Left            =   -15
      TabIndex        =   122
      Top             =   -15
      Width           =   8895
      Begin MSComctlLib.TabStrip OpcaoTributacao 
         Height          =   4560
         Left            =   45
         TabIndex        =   24
         Top             =   0
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   8043
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Resumo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhamento"
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
End
Attribute VB_Name = "TabTributacaoFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gobjForm As Object

Public iValorIRRFAlterado As Integer
Public iISSRetidoAlterado As Integer
Public iPISRetidoAlterado As Integer
Public iPISValorAlterado As Integer
Public iCOFINSRetidoAlterado As Integer
Public iCOFINSValorAlterado As Integer
Public iCSLLRetidoAlterado As Integer

Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get gobjTribTab() As ClassTribTab
    If Not (gobjForm Is Nothing) Then Set gobjTribTab = gobjForm.gobjTribTab
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Property Get iAlterado() As Integer
    If Not (gobjForm Is Nothing) Then iAlterado = gobjForm.iAlterado
End Property

Public Property Let iAlterado(vData As Integer)
    If Not (gobjForm Is Nothing) Then gobjForm.iAlterado = iAlterado
End Property

Public Function Limpa_TabTrib()
    Call Limpa_Tela(Me)
End Function

Private Sub FCI_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.FCI_Change
End Sub

Private Sub FCI_Validate(Cancel As Boolean)
    Call gobjTribTab.FCI_Validate(Cancel)
End Sub

Private Sub IPICNPJProdutor_GotFocus()
    Call gobjTribTab.IPICNPJProdutor_GotFocus(iAlterado)
End Sub

Private Sub IPICodSelo_GotFocus()
    Call gobjTribTab.IPICodSelo_GotFocus(iAlterado)
End Sub


Private Sub IPISeloQtde_GotFocus()
    Call gobjTribTab.IPISeloQtde_GotFocus(iAlterado)
End Sub

Public Sub NatOpInterna_GotFocus()
    
    Call gobjTribTab.NatOpInterna_GotFocus(iAlterado)

End Sub

Public Sub NaturezaItemLabel_Click()

    Call gobjTribTab.NaturezaItemLabel_Click

End Sub

Public Sub NaturezaOpItem_GotFocus()
    
    Call gobjTribTab.NaturezaOpItem_GotFocus(iAlterado)

End Sub


Public Sub TipoTributacao_GotFocus()
    
    Call gobjTribTab.TipoTributacao_GotFocus(iAlterado)

End Sub

Public Sub TipoTributacaoItem_GotFocus()
    
    Call gobjTribTab.TipoTributacaoItem_GotFocus(iAlterado)

End Sub

Private Sub TotTrib_Change()
    Call gobjTribTab.TotTrib_Change
End Sub

Private Sub TotTrib_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTrib_Validate(Cancel)
End Sub

Private Sub TotTribItem_Change()
    Call gobjTribTab.TotTribItem_Change
End Sub

Private Sub TotTribItem_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTribItem_Validate(Cancel)
End Sub

Private Sub TotTribItemTipo_Change()
    Call gobjTribTab.TotTribItemTipo_Click
End Sub

Private Sub TotTribItemTipo_Click()
    Call gobjTribTab.TotTribItemTipo_Click
End Sub

Private Sub TotTribTipo_Change()
    Call gobjTribTab.TotTribTipo_Click
End Sub

Private Sub TotTribTipo_Click()
    Call gobjTribTab.TotTribTipo_Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NatOpInterna Then
            Call LblNatOpInterna_Click
        ElseIf Me.ActiveControl Is TipoTributacao Then
            Call LblTipoTrib_Click
        End If

    End If

End Sub

Public Sub ValorIRRF_Validate(Cancel As Boolean)

    Call gobjTribTab.ValorIRRF_Validate(Cancel)

End Sub

Public Sub ComboICMSTipo_Click()

    Call gobjTribTab.ComboICMSTipo_Click

End Sub

Public Sub ComboICMSTipo_Change()

    Call gobjTribTab.ComboICMSTipo_Click

End Sub

Public Sub ComboIPITipo_Click()

    Call gobjTribTab.ComboIPITipo_Click

End Sub

Public Sub ComboIPITipo_Change()

    Call gobjTribTab.ComboIPITipo_Click

End Sub

Public Sub ComboItensTrib_Click()

    Call gobjTribTab.ComboItensTrib_Click

End Sub

Public Sub LblNatOpInterna_Click()

    Call gobjTribTab.LblNatOpInterna_Click

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set gobjTribTab = Nothing
    
End Sub

Public Sub LblTipoTrib_Click()

    Call gobjTribTab.LblTipoTrib_Click

End Sub

Public Sub LblTipoTribItem_Click()

    Call gobjTribTab.LblTipoTribItem_Click

End Sub

Public Sub NaturezaOpItem_Change()

    Call gobjTribTab.NaturezaOpItem_Change

End Sub

Public Sub NaturezaOpItem_Validate(Cancel As Boolean)

    Call gobjTribTab.NaturezaOpItem_Validate(Cancel)

End Sub

Public Sub TipoTributacao_Change()

    iAlterado = REGISTRO_ALTERADO

    Call gobjTribTab.TipoTributacao_Change

End Sub

Public Sub TipoTributacao_Validate(Cancel As Boolean)

    Call gobjTribTab.TipoTributacao_Validate(Cancel)

End Sub

Public Sub TipoTributacaoItem_Change()

    Call gobjTribTab.TipoTributacaoItem_Change

End Sub

Public Sub TipoTributacaoItem_Validate(Cancel As Boolean)

    Call gobjTribTab.TipoTributacaoItem_Validate(Cancel)

End Sub

Public Sub TributacaoRecalcular_Click()

    Call gobjTribTab.TributacaoRecalcular_Click

End Sub

Public Sub OpcaoTributacao_Click()

    Call gobjTribTab.OpcaoTributacao_Click

End Sub

Public Sub ValorIRRF_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorIRRFAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ValorIRRF_Change

End Sub

Public Sub ICMSAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSAliquotaItem_Change

End Sub

Public Sub ICMSAliquotaItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSAliquotaItem_Validate(Cancel)

End Sub

Public Sub ICMSBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSBaseItem_Change

End Sub

Public Sub ICMSBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSBaseItem_Validate(Cancel)

End Sub

Public Sub ICMSPercRedBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSPercRedBaseItem_Change

End Sub

Public Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSPercRedBaseItem_Validate(Cancel)

End Sub

Public Sub ICMSSubstAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstAliquotaItem_Change

End Sub

Public Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSubstAliquotaItem_Validate(Cancel)

End Sub

Public Sub ICMSSubstBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstBaseItem_Change

End Sub

Public Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSubstBaseItem_Validate(Cancel)

End Sub

Public Sub ICMSSubstValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstValorItem_Change

End Sub

Public Sub ICMSSubstValorItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSubstValorItem_Validate(Cancel)

End Sub

Public Sub ICMSValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSValorItem_Change

End Sub

Public Sub ICMSValorItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSValorItem_Validate(Cancel)

End Sub

Public Sub IPIAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIAliquotaItem_Change

End Sub

Public Sub IPIAliquotaItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIAliquotaItem_Validate(Cancel)

End Sub

Public Sub IPIBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIBaseItem_Change

End Sub

Public Sub IPIBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIBaseItem_Validate(Cancel)

End Sub

Public Sub IPIPercRedBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIPercRedBaseItem_Change

End Sub

Public Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIPercRedBaseItem_Validate(Cancel)

End Sub

Public Sub IPIValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIValorItem_Change

End Sub

Public Sub IPIValorItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIValorItem_Validate(Cancel)

End Sub

Public Sub IRAliquota_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IRAliquota_Change

End Sub

Public Sub IRAliquota_Validate(Cancel As Boolean)

    Call gobjTribTab.IRAliquota_Validate(Cancel)
    
End Sub

Public Sub INSSBase_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSBase_Change
End Sub

Public Sub INSSBase_Validate(Cancel As Boolean)
    Call gobjTribTab.INSSBase_Validate(Cancel)
End Sub

Public Sub INSSDeducoes_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSDeducoes_Change
End Sub

Public Sub INSSDeducoes_Validate(Cancel As Boolean)
    Call gobjTribTab.INSSDeducoes_Validate(Cancel)
End Sub

Public Sub INSSRetido_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSRetido_Click
End Sub

Public Sub INSSValor_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSValor_Change
End Sub

Public Sub INSSValor_Validate(Cancel As Boolean)
    Call gobjTribTab.INSSValor_Validate(Cancel)
End Sub

Public Sub PISRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISRetido_Change
    iPISRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PISRetido_Validate(Cancel As Boolean)
    
    If iPISRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.PISRetido_Validate(Cancel)
    iPISRetidoAlterado = 0

End Sub

Public Sub ISSRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSRetido_Change
    iISSRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ISSRetido_Validate(Cancel As Boolean)
    
    If iISSRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.ISSRetido_Validate(Cancel)
    iISSRetidoAlterado = 0

End Sub

Public Sub COFINSRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSRetido_Change
    iCOFINSRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub COFINSRetido_Validate(Cancel As Boolean)
    
    If iCOFINSRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.COFINSRetido_Validate(Cancel)
    iCOFINSRetidoAlterado = 0

End Sub

Public Sub CSLLRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.CSLLRetido_Change
    iCSLLRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CSLLRetido_Validate(Cancel As Boolean)
    
    If iCSLLRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.CSLLRetido_Validate(Cancel)
    iCSLLRetidoAlterado = 0

End Sub

Public Sub NatOpInterna_Change()
    Call gobjTribTab.NatOpInterna_Change
End Sub

Public Sub NatOpInterna_Validate(Cancel As Boolean)
    Call gobjTribTab.NatOpInterna_Validate(Cancel)
End Sub

'##################################################################
'Novo código de Tributação NFe - 19/06/2009
'ALTERAÇÕES FEITAS
'=================
'RETIRADO
'--------
'TribSobreDesconto
'TribSobreFrete
'TribSobreItem
'TribSobreSeguro
'TribSobreOutrasDesp
'ISSAliquota
'
'DE MASK PARA LABEL
'------------------
'ISSValor
'PISValor
'COFINSValor
'
'DE LABEL PARA MASK
'------------------
'IRBase
'
'OUTROS
'------
'ISSIncluso -Invisível(Value = True)
'
'INCLUÍDO
'--------
'OpcaoTribDet (TabStrip)
'FrameTabTribDet (1 ao 7)
'ValorDescontoItem (Mask)
'ValorFreteItem (Mask)
'ValorSeguroItem (Mask)
'ValorDespesasItem (Mask)
'UMTrib (Label)
'QtdeTrib (Label)
'ValorTrib (Label)
'OrigemMercadoria (Combo)
'ICMSModalidade (Combo)
'ICMSSTModalidade (Combo)
'ICMSSubstVlrAddItem (Mask)
'ICMSSTPercRedBaseItem (Mask)
'IPITipoCalculo (Combo)
'IPICodEnq (Mask)
'IPIClasseEnq (Mask)
'IPICNPJProdutor (Mask)
'IPICodSelo (Mask)
'IPISeloQtde (Mask)
'IPIValorUnitario (Mask)
'IPIQtd (Mask)
'FrameIPITipoCalculoBase (Para abilitar ou desabilitar conforme IPITipoCalculo )
'FrameIPITipoCalculoQtd (Para abilitar ou desabilitar conforme IPITipoCalculo )
'ComboPISTipo (Combo)
'PISTipoCalculo (Combo)
'PISSTTipoCalculo (Combo)
'FramePISTipoCalculoBase (Para abilitar ou desabilitar conforme PISTipoCalculo )
'FramePISTipoCalculoQtd (Para abilitar ou desabilitar conforme PISTipoCalculo )
'FramePISSTTipoCalculoBase (Para abilitar ou desabilitar conforme PISSTTipoCalculo )
'FramePISSTTipoCalculoQtd (Para abilitar ou desabilitar conforme PISSTTipoCalculo )
'PISBaseItem (Mask)
'PISAliquotaItem (Mask)
'PISAliquotaRSItem (Mask)
'PISQtdeItem (Mask)
'PISValorItem (Mask)
'PISSTBaseItem (Mask)
'PISSTAliquotaItem (Mask)
'PISSTAliquotaRSItem (Mask)
'PISSTQtdeItem (Mask)
'PISSTValorItem (Mask)
'ComboCOFINSTipo (Combo)
'COFINSTipoCalculo (Combo)
'COFINSSTTipoCalculo (Combo)
'FrameCOFINSTipoCalculoBase (Para abilitar ou desabilitar conforme COFINSTipoCalculo )
'FrameCOFINSTipoCalculoQtd (Para abilitar ou desabilitar conforme COFINSTipoCalculo )
'FrameCOFINSSTTipoCalculoBase (Para abilitar ou desabilitar conforme COFINSSTTipoCalculo )
'FrameCOFINSSTTipoCalculoQtd (Para abilitar ou desabilitar conforme COFINSSTTipoCalculo )
'COFINSBaseItem (Mask)
'COFINSAliquotaItem (Mask)
'COFINSAliquotaRSItem (Mask)
'COFINSQtdeItem (Mask)
'COFINSValorItem (Mask)
'COFINSSTBaseItem (Mask)
'COFINSSTAliquotaItem (Mask)
'COFINSSTAliquotaRSItem (Mask)
'COFINSSTQtdeItem (Mask)
'COFINSSTValorItem (Mask)
'ISSBaseItem (Mask)
'ISSAliquotaItem (Mask)
'ISSValorItem (Mask)
'ISSListaServ (Combo)
'ISSCodIBGE (Mask)
'ISSDescIBGE (Label)
'IIBaseItem (Mask)
'IIDespAduItem (Mask)
'IIIOFItem (Mask)
'IIValorItem (Mask)
'ISSLabelCodIBGE (Caption - Tem que chamar o Browser)

Public Sub OpcaoTribDet_Click()
    Call gobjTribTab.OpcaoTribDet_Click
End Sub

Private Sub IPITipoCalculo_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPITipoCalculo_Click
End Sub

Private Sub PISTipoCalculo_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISTipoCalculo_Click
End Sub

Private Sub PISSTTipoCalculo_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISSTTipoCalculo_Click
End Sub

Private Sub COFINSTipoCalculo_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSTipoCalculo_Click
End Sub

Private Sub COFINSSTTipoCalculo_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSSTTipoCalculo_Click
End Sub

Private Sub ComboPISTipo_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ComboPISTipo_Click
End Sub

Private Sub ComboPISTipo_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ComboPISTipo_Click
End Sub

Private Sub ComboCOFINSTipo_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ComboCOFINSTipo_Click
End Sub

Private Sub ComboCOFINSTipo_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ComboCOFINSTipo_Click
End Sub

Private Sub ISSValorItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSValorItem_Validate(Cancel)
End Sub

Private Sub ISSAliquotaItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSAliquotaItem_Change
End Sub

Private Sub ISSAliquotaItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSAliquotaItem_Validate(Cancel)
End Sub

Private Sub ISSBaseItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSBaseItem_Change
End Sub

Private Sub ISSBaseItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSBaseItem_Validate(Cancel)
End Sub

Private Sub ISSValorItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSValorItem_Change
End Sub

Private Sub PISValorItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISValorItem_Validate(Cancel)
End Sub

Private Sub PISAliquotaItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISAliquotaItem_Change
End Sub

Private Sub PISAliquotaItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISAliquotaItem_Validate(Cancel)
End Sub

Private Sub PISBaseItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISBaseItem_Change
End Sub

Private Sub PISBaseItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISBaseItem_Validate(Cancel)
End Sub

Private Sub PISValorItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISValorItem_Change
End Sub

Private Sub PISAliquotaRSItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISAliquotaRSItem_Validate(Cancel)
End Sub

Private Sub PISAliquotaRSItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISAliquotaRSItem_Change
End Sub

Private Sub PISQtdeItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISQtdeItem_Validate(Cancel)
End Sub

Private Sub PISQtdeItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISQtdeItem_Change
End Sub

Private Sub PISSTValorItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISSTValorItem_Validate(Cancel)
End Sub

Private Sub PISSTAliquotaItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISSTAliquotaItem_Change
End Sub

Private Sub PISSTAliquotaItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISSTAliquotaItem_Validate(Cancel)
End Sub

Private Sub PISSTBaseItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISSTBaseItem_Change
End Sub

Private Sub PISSTBaseItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISSTBaseItem_Validate(Cancel)
End Sub

Private Sub PISSTValorItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISSTValorItem_Change
End Sub

Private Sub PISSTAliquotaRSItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISSTAliquotaRSItem_Validate(Cancel)
End Sub

Private Sub PISSTAliquotaRSItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISSTAliquotaRSItem_Change
End Sub

Private Sub PISSTQtdeItem_Validate(Cancel As Boolean)
    Call gobjTribTab.PISSTQtdeItem_Validate(Cancel)
End Sub

Private Sub PISSTQtdeItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISSTQtdeItem_Change
End Sub

Private Sub COFINSValorItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSValorItem_Validate(Cancel)
End Sub

Private Sub COFINSAliquotaItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSAliquotaItem_Change
End Sub

Private Sub COFINSAliquotaItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSAliquotaItem_Validate(Cancel)
End Sub

Private Sub COFINSBaseItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSBaseItem_Change
End Sub

Private Sub COFINSBaseItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSBaseItem_Validate(Cancel)
End Sub

Private Sub COFINSValorItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSValorItem_Change
End Sub

Private Sub COFINSAliquotaRSItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSAliquotaRSItem_Validate(Cancel)
End Sub

Private Sub COFINSAliquotaRSItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSAliquotaRSItem_Change
End Sub

Private Sub COFINSQtdeItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSQtdeItem_Validate(Cancel)
End Sub

Private Sub COFINSQtdeItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSQtdeItem_Change
End Sub

Private Sub COFINSSTValorItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSSTValorItem_Validate(Cancel)
End Sub

Private Sub COFINSSTAliquotaItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSSTAliquotaItem_Change
End Sub

Private Sub COFINSSTAliquotaItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSSTAliquotaItem_Validate(Cancel)
End Sub

Private Sub COFINSSTBaseItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSSTBaseItem_Change
End Sub

Private Sub COFINSSTBaseItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSSTBaseItem_Validate(Cancel)
End Sub

Private Sub COFINSSTValorItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSSTValorItem_Change
End Sub

Private Sub COFINSSTAliquotaRSItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSSTAliquotaRSItem_Validate(Cancel)
End Sub

Private Sub COFINSSTAliquotaRSItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSSTAliquotaRSItem_Change
End Sub

Private Sub COFINSSTQtdeItem_Validate(Cancel As Boolean)
    Call gobjTribTab.COFINSSTQtdeItem_Validate(Cancel)
End Sub

Private Sub COFINSSTQtdeItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSSTQtdeItem_Change
End Sub

Private Sub ISSCodIBGE_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSCodIBGE_Change
End Sub

Private Sub ISSCodIBGE_GotFocus()
    Call gobjTribTab.ISSCodIBGE_GotFocus
End Sub

Private Sub ISSCodIBGE_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSCodIBGE_Validate(Cancel)
End Sub

Private Sub ISSLabelCodIBGE_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSLabelCodIBGE_Click
End Sub

Private Sub IIBaseItem_Validate(Cancel As Boolean)
    Call gobjTribTab.IIBaseItem_Validate(Cancel)
End Sub

Private Sub IIBaseItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IIBaseItem_Change
End Sub

Private Sub IIValorItem_Validate(Cancel As Boolean)
    Call gobjTribTab.IIValorItem_Validate(Cancel)
End Sub

Private Sub IIValorItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IIValorItem_Change
End Sub

Private Sub IIIOFItem_Validate(Cancel As Boolean)
    Call gobjTribTab.IIIOFItem_Validate(Cancel)
End Sub

Private Sub IIIOFItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IIIOFItem_Change
End Sub

Private Sub IIDespAduItem_Validate(Cancel As Boolean)
    Call gobjTribTab.IIDespAduItem_Validate(Cancel)
End Sub

Private Sub IIDespAduItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IIDespAduItem_Change
End Sub

Private Sub ValorFreteItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ValorFreteItem_Validate(Cancel)
End Sub

Private Sub ValorFreteItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ValorFreteItem_Change
End Sub

Private Sub ValorSeguroItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ValorSeguroItem_Validate(Cancel)
End Sub

Private Sub ValorSeguroItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ValorSeguroItem_Change
End Sub

Private Sub ValorDescontoItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ValorDescontoItem_Validate(Cancel)
End Sub

Private Sub ValorDescontoItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ValorDescontoItem_Change
End Sub

Private Sub ValorDespesasItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ValorDespesasItem_Validate(Cancel)
End Sub

Private Sub ValorDespesasItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ValorDespesasItem_Change
End Sub

Private Sub ICMSModalidade_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSModalidade_Click
End Sub

Private Sub ICMSSTModalidade_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSTModalidade_Click
End Sub

Private Sub OrigemMercadoria_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.OrigemMercadoria_Click
End Sub

Private Sub ICMSSubstPercRedBaseItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstPercRedBaseItem_Change
End Sub

Private Sub ICMSSubstPercRedBaseItem_Validate(Cancel As Boolean)
   Call gobjTribTab.ICMSSubstPercRedBaseItem_Validate(Cancel)
End Sub

Private Sub ICMSSubstPercMVAItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstPercMVAItem_Change
End Sub

Private Sub ICMSSubstPercMVAItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSSubstPercMVAItem_Validate(Cancel)
End Sub

Private Sub IPICodEnq_Validate(Cancel As Boolean)
    Call gobjTribTab.IPICodEnq_Validate(Cancel)
End Sub

Private Sub IPICodEnq_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIClasseEnq_Validate(Cancel As Boolean)
    Call gobjTribTab.IPIClasseEnq_Validate(Cancel)
End Sub

Private Sub IPIClasseEnq_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPICodSelo_Validate(Cancel As Boolean)
    Call gobjTribTab.IPICodSelo_Validate(Cancel)
End Sub

Private Sub IPICodSelo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPISeloQtde_Validate(Cancel As Boolean)
    Call gobjTribTab.IPISeloQtde_Validate(Cancel)
End Sub

Private Sub IPISeloQtde_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPICNPJProdutor_Validate(Cancel As Boolean)
    Call gobjTribTab.IPICNPJProdutor_Validate(Cancel)
End Sub

Private Sub IPICNPJProdutor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIUnidadePadraoValorItem_Validate(Cancel As Boolean)
    Call gobjTribTab.IPIUnidadePadraoValorItem_Validate(Cancel)
End Sub

Private Sub IPIUnidadePadraoValorItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIUnidadePadraoValorItem_Change
End Sub

Private Sub IPIUnidadePadraoQtdeItem_Validate(Cancel As Boolean)
    Call gobjTribTab.IPIUnidadePadraoQtdeItem_Validate(Cancel)
End Sub

Private Sub IPIUnidadePadraoQtdeItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIUnidadePadraoQtdeItem_Change
End Sub

Private Sub IRBase_Validate(Cancel As Boolean)
    Call gobjTribTab.IRBase_Validate(Cancel)
End Sub

Private Sub IRBase_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IRBase_Change
End Sub

Private Sub ISSListaServ_Click()
    Call gobjTribTab.ISSListaServ_Click
End Sub


Public Sub ICMSSTCobrAntBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSTCobrAntBaseItem_Change

End Sub

Public Sub ICMSSTCobrAntBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSTCobrAntBaseItem_Validate(Cancel)

End Sub

Public Sub ICMSSTCobrAntValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSTCobrAntValorItem_Change

End Sub

Public Sub ICMSSTCobrAntValorItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSTCobrAntValorItem_Validate(Cancel)

End Sub

Public Sub ICMSpercBaseOperacaoPropria_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSpercBaseOperacaoPropria_Change
End Sub

Public Sub ICMSpercBaseOperacaoPropria_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSpercBaseOperacaoPropria_Validate(Cancel)
End Sub

Public Sub ICMSUFDevidoST_Change()
    iAlterado = REGISTRO_ALTERADO
    'Call gobjTribTab.ICMSUFDevidoST_Change
End Sub

Public Sub ICMSUFDevidoST_Click()
    Call gobjTribTab.ICMSUFDevidoST_Click
End Sub

Public Sub ICMSvBCSTRet_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvBCSTRet_Change
End Sub

Public Sub ICMSvBCSTRet_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvBCSTRet_Validate(Cancel)
End Sub

Public Sub ICMSvICMSSTRet_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvICMSSTRet_Change
End Sub

Public Sub ICMSvICMSSTRet_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvICMSSTRet_Validate(Cancel)
End Sub

Public Sub ICMSvBCSTDest_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvBCSTDest_Change
End Sub

Public Sub ICMSvBCSTDest_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvBCSTDest_Validate(Cancel)
End Sub

Public Sub ICMSvICMSSTDest_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvICMSSTDest_Change
End Sub

Public Sub ICMSvICMSSTDest_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvICMSSTDest_Validate(Cancel)
End Sub

Public Sub ICMSpCredSN_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSpCredSN_Change
End Sub

Public Sub ICMSpCredSN_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSpCredSN_Validate(Cancel)
End Sub

Public Sub ICMSvCredSN_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvCredSN_Change
End Sub

Public Sub ICMSvCredSN_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvCredSN_Validate(Cancel)
End Sub

Public Sub ICMSValorIsento_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSValorIsento_Change
End Sub

Public Sub ICMSValorIsento_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSValorIsento_Validate(Cancel)
End Sub

Public Sub ICMSMotivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ICMSMotivo_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSMotivo_Validate(Cancel)
End Sub

Public Sub ICMSMotivo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ComboISSTipo_Click()
    Call gobjTribTab.ComboISSTipo_Click
End Sub

Public Sub ComboICMSSimplesTipo_Click()
    Call gobjTribTab.ComboICMSSimplesTipo_Click
End Sub

Private Sub ComboICMSSimplesTipo_Change()
    Call gobjTribTab.ComboICMSSimplesTipo_Click
End Sub

Private Sub NatBCCred_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NatBCCred_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.NatBCCred_Click
End Sub
'##################################################################

'--------------------------------NFe 3.10 --------------------------------------------
Private Sub OpcaoResumo_Click()
    Call gobjTribTab.OpcaoResumo_Click
End Sub

Private Sub IndPresenca_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IndPresenca_Click
End Sub

Private Sub IndPresenca_Click()
    Call gobjTribTab.IndPresenca_Click
End Sub

Private Sub IndConsumidorFinal_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IndConsumidorFinal_Click
End Sub

'--ICMSDesonerado-Caption = Total do ICMS Isento
'--ISSValorDeducao-Caption = Total das Deduções de ISS
'--ISSValorOutrasRet-Caption = Total do ISS Outras Retenções
'--ISSValorDescIncond-Caption = Total do ISS Desconto Incondicionado
'--ISSValorDescCond-Caption = Total do ISS Desconto Condicionado

Private Sub DataPrestServico_Change()
    iAlterado = REGISTRO_ALTERADO
     Call gobjTribTab.DataPrestServico_Change
End Sub

Private Sub DataPrestServico_GotFocus()
     Call gobjTribTab.DataPrestServico_GotFocus(iAlterado)
End Sub

Private Sub DataPrestServico_Validate(Cancel As Boolean)
     Call gobjTribTab.DataPrestServico_Validate(Cancel)
End Sub

Private Sub UpDownPrestServico_DownClick()
     Call gobjTribTab.UpDownPrestServico_DownClick
End Sub

Private Sub UpDownPrestServico_UpClick()
     Call gobjTribTab.UpDownPrestServico_UpClick
End Sub

'--PISBase-Caption = Soma da Base do PIS nos itens
'--PISBaseST-Caption = Soma da Base ST do PIS nos itens
'--PISValorST-Caption = Soma do Valor do PIS nos itens
'--COFINSBase-Caption = Soma da Base do COFINS nos itens
'--COFINSBaseST-Caption = Soma da Base ST do COFINS nos itens
'--COFINSValorST-Caption = Soma do Valor do COFINS nos itens
'--IIBase-Caption = Soma da Base do II nos itens
'--IIDespAdu-Caption = Soma da Despesa Aduaneira nos itens
'--IIIOF-Caption = Soma do IOF nos itens
'--IIValor-Caption = Soma do valor do II nos itens

Private Sub ICMS51ValorOpItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMS51ValorOpItem_Change
End Sub

Private Sub ICMS51ValorOpItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMS51ValorOpItem_Validate(Cancel)
End Sub

Private Sub ICMSPercDiferItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSPercDiferItem_Change
End Sub

Private Sub ICMSPercDiferItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSPercDiferItem_Validate(Cancel)
End Sub

Private Sub ICMSValorDifItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSValorDifItem_Change
End Sub

Private Sub ICMSValorDifItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSValorDifItem_Validate(Cancel)
End Sub

Private Sub ISSIndIncentivoItem_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSIndIncentivoItem_Click
End Sub

Private Sub ISSValorDeducaoItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSValorDeducaoItem_Change
End Sub

Private Sub ISSValorDeducaoItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSValorDeducaoItem_Validate(Cancel)
End Sub

Private Sub ISSValorOutrasRetItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSValorOutrasRetItem_Change
End Sub

Private Sub ISSValorOutrasRetItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSValorOutrasRetItem_Validate(Cancel)
End Sub

Private Sub ISSValorDescIncondItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSValorDescIncondItem_Change
End Sub

Private Sub ISSValorDescIncondItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSValorDescIncondItem_Validate(Cancel)
End Sub

Private Sub ISSValorDescCondItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSValorDescCondItem_Change
End Sub

Private Sub ISSValorDescCondItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSValorDescCondItem_Validate(Cancel)
End Sub

Private Sub ISSValorRetItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSValorRetItem_Change
End Sub

Private Sub ISSValorRetItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSValorRetItem_Validate(Cancel)
End Sub

Private Sub ISSIndExigibilidadeItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSIndExigibilidadeItem_Click
End Sub

Private Sub ISSIndExigibilidadeItem_Click()
    Call gobjTribTab.ISSIndExigibilidadeItem_Click
End Sub

Private Sub ISSCodServItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSCodServItem_Change
End Sub

Private Sub ISSCodServItemLabel_Click()
    Call gobjTribTab.ISSCodServItemLabel_Click
End Sub

Private Sub ISSCodServItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSCodServItem_Validate(Cancel)
End Sub

'--ISSCodServDesc-Caption=Descrição do código de ISSCodServItem

Private Sub ISSLabelMunicIncidImp_Click()
    Call gobjTribTab.ISSLabelMunicIncidImp_Click
End Sub

Private Sub ISSMunicIncidImpItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSMunicIncidImpItem_Change
End Sub

Private Sub ISSMunicIncidImpItem_GotFocus()
    Call gobjTribTab.ISSMunicIncidImpItem_GotFocus(iAlterado)
End Sub

Private Sub ISSMunicIncidImpItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSMunicIncidImpItem_Validate(Cancel)
End Sub

'--ISSMunicIncidDesc-Caption=Descrição do município de ISSMunicIncidImpItem

Private Sub ISSPaisLabel_Click()
    Call gobjTribTab.ISSPaisLabel_Click
End Sub

Private Sub ISSCodPaisItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSCodPaisItem_Change
End Sub

Private Sub ISSCodPaisItem_Click()
    Call gobjTribTab.ISSCodPaisItem_Click
End Sub

Private Sub ISSCodPaisItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSCodPaisItem_Validate(Cancel)
End Sub

Private Sub ISSNumProcessoItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSNumProcessoItem_Change
End Sub

Private Sub ISSNumProcessoItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ISSNumProcessoItem_Validate(Cancel)
End Sub
'--------------------------------NFe 3.10 --------------------------------------------

'##################################
'IBPTax 0.0.9
Private Sub TotTribFed_Change()
    Call gobjTribTab.TotTribFed_Change
End Sub

Private Sub TotTribFed_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTribFed_Validate(Cancel)
End Sub

Private Sub TotTribEst_Change()
    Call gobjTribTab.TotTribEst_Change
End Sub

Private Sub TotTribEst_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTribEst_Validate(Cancel)
End Sub

Private Sub TotTribMunic_Change()
    Call gobjTribTab.TotTribMunic_Change
End Sub

Private Sub TotTribMunic_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTribMunic_Validate(Cancel)
End Sub

Private Sub TotTribAliqFedItem_Change()
    Call gobjTribTab.TotTribAliqFedItem_Change
End Sub

Private Sub TotTribAliqFedItem_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTribAliqFedItem_Validate(Cancel)
End Sub

Private Sub TotTribAliqEstItem_Change()
    Call gobjTribTab.TotTribAliqEstItem_Change
End Sub

Private Sub TotTribAliqEstItem_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTribAliqEstItem_Validate(Cancel)
End Sub

Private Sub TotTribAliqMunicItem_Change()
    Call gobjTribTab.TotTribAliqMunicItem_Change
End Sub

Private Sub TotTribAliqMunicItem_Validate(Cancel As Boolean)
    Call gobjTribTab.TotTribAliqMunicItem_Validate(Cancel)
End Sub
'##################################

Public Sub ICMSInterestPercFCPUFDestItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestPercFCPUFDestItem_Change
End Sub

Public Sub ICMSInterestPercFCPUFDestItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSInterestPercFCPUFDestItem_Validate(Cancel)
End Sub

Public Sub ICMSInterestBCUFDestItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestBCUFDestItem_Change
End Sub

Public Sub ICMSInterestBCUFDestItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSInterestBCUFDestItem_Validate(Cancel)
End Sub

Public Sub ICMSInterestAliqUFDestItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestAliqUFDestItem_Change
End Sub

Public Sub ICMSInterestAliqUFDestItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSInterestAliqUFDestItem_Validate(Cancel)
End Sub

Private Sub ICMSInterestAliqItem_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestAliqItem_Click
End Sub

Private Sub ICMSInterestPercPartilhaItem_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestPercPartilhaItem_Click
End Sub

Public Sub ICMSInterestVlrUFDestItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestVlrUFDestItem_Change
End Sub

Public Sub ICMSInterestVlrUFDestItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSInterestVlrUFDestItem_Validate(Cancel)
End Sub

Public Sub ICMSInterestVlrUFRemetItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestVlrUFRemetItem_Change
End Sub

Public Sub ICMSInterestVlrUFRemetItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSInterestVlrUFRemetItem_Validate(Cancel)
End Sub

Public Sub ICMSInterestVlrFCPUFDestItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestVlrFCPUFDestItem_Change
End Sub

Public Sub ICMSInterestVlrFCPUFDestItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSInterestVlrFCPUFDestItem_Validate(Cancel)
End Sub

Private Sub IPICodEnqLabel_Click()
    Call gobjTribTab.IPICodEnqLabel_Click
End Sub

Private Sub IPICodEnq_GotFocus()
    Call gobjTribTab.IPICodEnq_GotFocus(iAlterado)
End Sub

'nfe 4.00
Private Sub cBenefItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.cBenefItem_Change
End Sub

Public Sub cBenefItem_Validate(Cancel As Boolean)
    Call gobjTribTab.cBenefItem_Validate(Cancel)
End Sub

Private Sub ICMSpFCPItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSpFCPItem_Change
End Sub

Private Sub ICMSvBCFCPItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvBCFCPItem_Change
End Sub

Private Sub ICMSvFCPItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvFCPItem_Change
End Sub

Public Sub ICMSpFCPItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSpFCPItem_Validate(Cancel)
End Sub

Public Sub ICMSvBCFCPItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvBCFCPItem_Validate(Cancel)
End Sub

Public Sub ICMSvFCPItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvFCPItem_Validate(Cancel)
End Sub

Private Sub ICMSpFCPSTItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSpFCPSTItem_Change
End Sub

Private Sub ICMSvBCFCPSTItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvBCFCPSTItem_Change
End Sub

Private Sub ICMSvFCPSTItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvFCPSTItem_Change
End Sub

Public Sub ICMSpFCPSTItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSpFCPSTItem_Validate(Cancel)
End Sub

Public Sub ICMSvBCFCPSTItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvBCFCPSTItem_Validate(Cancel)
End Sub

Public Sub ICMSvFCPSTItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvFCPSTItem_Validate(Cancel)
End Sub

Private Sub ICMSpFCPSTRetItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSpFCPSTRetItem_Change
End Sub

Private Sub ICMSvBCFCPSTRetItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvBCFCPSTRetItem_Change
End Sub

Private Sub ICMSvFCPSTRetItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSvFCPSTRetItem_Change
End Sub

Public Sub ICMSpFCPSTRetItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSpFCPSTRetItem_Validate(Cancel)
End Sub

Public Sub ICMSvBCFCPSTRetItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvBCFCPSTRetItem_Validate(Cancel)
End Sub

Public Sub ICMSvFCPSTRetItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSvFCPSTRetItem_Validate(Cancel)
End Sub

Private Sub ICMSInterestBCFCPUFDestItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestBCFCPUFDestItem_Change
End Sub

Private Sub ICMSInterestBCFCPUFDestItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSInterestBCFCPUFDestItem_Validate(Cancel)
End Sub

Private Sub ICMSSTCobrAntAliquotaItem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSTCobrAntAliquotaItem_Change
End Sub

Private Sub ICMSSTCobrAntAliquotaItem_Validate(Cancel As Boolean)
    Call gobjTribTab.ICMSSTCobrAntAliquotaItem_Validate(Cancel)
End Sub

Private Sub IPIVlrDevolvidoItem_Change()
    Call gobjTribTab.IPIVlrDevolvidoItem_Change
End Sub

Private Sub IPIVlrDevolvidoItem_Validate(Cancel As Boolean)
    Call gobjTribTab.IPIVlrDevolvidoItem_Validate(Cancel)
End Sub

Private Sub pDevolItem_Change()
    Call gobjTribTab.pDevolItem_Change
End Sub

Private Sub pDevolItem_Validate(Cancel As Boolean)
    Call gobjTribTab.pDevolItem_Validate(Cancel)
End Sub

Private Sub ICMSSTBaseDupla_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSTBaseDupla_Click
End Sub

Private Sub ICMSInterestBaseDupla_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSInterestBaseDupla_Click
End Sub
