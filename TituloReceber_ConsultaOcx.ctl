VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TituloReceber_ConsultaOcx 
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   ScaleHeight     =   6825
   ScaleWidth      =   10215
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   9930
      Begin VB.CommandButton BotaoProjetos 
         Caption         =   "..."
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
         Left            =   3480
         TabIndex        =   225
         Top             =   1095
         Width           =   495
      End
      Begin VB.ComboBox Etapa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6570
         Style           =   2  'Dropdown List
         TabIndex        =   221
         Top             =   1095
         Width           =   2295
      End
      Begin VB.Frame Frame7 
         Caption         =   "Reajuste"
         Height          =   555
         Left            =   3645
         TabIndex        =   188
         Top             =   4650
         Width           =   6255
         Begin VB.Label ReajusteBase 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4830
            TabIndex        =   194
            Top             =   150
            Width           =   1305
         End
         Begin VB.Label Moeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3135
            TabIndex        =   193
            Top             =   165
            Width           =   960
         End
         Begin VB.Label ReajustePeriodicidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   960
            TabIndex        =   192
            Top             =   180
            Width           =   1305
         End
         Begin VB.Label Label46 
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
            Height          =   240
            Left            =   4275
            TabIndex        =   191
            Top             =   195
            Width           =   525
         End
         Begin VB.Label LabelMoeda 
            AutoSize        =   -1  'True
            Caption         =   "Índice:"
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
            Height          =   180
            Left            =   2475
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   190
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Periodic.:"
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
            Height          =   180
            Left            =   75
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   189
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.CommandButton BotaoDocOriginal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   7815
         Picture         =   "TituloReceber_ConsultaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   -45
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tributos"
         Height          =   2040
         Left            =   60
         TabIndex        =   58
         Top             =   3165
         Width           =   3465
         Begin VB.Frame Frame6 
            Caption         =   "Retenções"
            Height          =   1245
            Left            =   150
            TabIndex        =   175
            Top             =   165
            Width           =   3210
            Begin VB.Label Label56 
               Caption         =   "ISS:"
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
               Left            =   450
               TabIndex        =   217
               Top             =   915
               Width           =   375
            End
            Begin VB.Label ISSRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   840
               TabIndex        =   216
               Top             =   885
               Width           =   855
            End
            Begin VB.Label COFINSRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   825
               TabIndex        =   183
               Top             =   540
               Width           =   840
            End
            Begin VB.Label Label45 
               Caption         =   "COFINS:"
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
               Left            =   75
               TabIndex        =   182
               Top             =   585
               Width           =   720
            End
            Begin VB.Label CSLLRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2250
               TabIndex        =   181
               Top             =   555
               Width           =   855
            End
            Begin VB.Label Label43 
               Caption         =   "CSLL:"
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
               Left            =   1710
               TabIndex        =   180
               Top             =   600
               Width           =   555
            End
            Begin VB.Label PISRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2250
               TabIndex        =   179
               Top             =   195
               Width           =   855
            End
            Begin VB.Label Label32 
               Caption         =   "PIS:"
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
               Left            =   1860
               TabIndex        =   178
               Top             =   225
               Width           =   375
            End
            Begin VB.Label ValorIRRF 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   810
               TabIndex        =   177
               Top             =   210
               Width           =   855
            End
            Begin VB.Label Label20 
               Caption         =   "IR:"
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
               Left            =   555
               TabIndex        =   176
               Top             =   225
               Width           =   270
            End
         End
         Begin VB.Frame SSFrame6 
            Caption         =   "INSS"
            Height          =   600
            Left            =   150
            TabIndex        =   59
            Top             =   1395
            Width           =   3210
            Begin VB.CheckBox INSSRetido 
               Caption         =   "Retido"
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
               Left            =   2100
               TabIndex        =   5
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label30 
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
               Height          =   210
               Left            =   105
               TabIndex        =   61
               Top             =   240
               Width           =   510
            End
            Begin VB.Label ValorINSS 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   645
               TabIndex        =   60
               Top             =   210
               Width           =   1230
            End
         End
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Parcelas"
         Height          =   1710
         Left            =   60
         TabIndex        =   62
         Top             =   1425
         Width           =   9840
         Begin VB.TextBox DescPrev 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4530
            MaxLength       =   255
            TabIndex        =   195
            Top             =   705
            Width           =   3990
         End
         Begin VB.TextBox StatusParcela 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3810
            TabIndex        =   9
            Top             =   690
            Width           =   645
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   255
            Left            =   1410
            TabIndex        =   7
            Top             =   690
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   255
            Left            =   2550
            TabIndex        =   8
            Top             =   690
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   255
            Left            =   210
            TabIndex        =   6
            Top             =   690
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1140
            Left            =   90
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   495
            Width           =   9630
            _ExtentX        =   16986
            _ExtentY        =   2011
            _Version        =   393216
            Rows            =   50
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label CondicaoPagamento 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3105
            TabIndex        =   162
            Top             =   150
            Width           =   2100
         End
         Begin VB.Label CondPagtoLabel 
            Caption         =   "Condição de Pagamento:"
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
            Left            =   840
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   161
            Top             =   195
            Width           =   2175
         End
      End
      Begin VB.Frame SSFrame7 
         Caption         =   "Comissões na Emissão"
         Height          =   1485
         Left            =   3660
         TabIndex        =   63
         Top             =   3165
         Width           =   6240
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   255
            Left            =   4440
            TabIndex        =   14
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox ValorBaseEmissao 
            Height          =   255
            Left            =   3030
            TabIndex        =   13
            Top             =   300
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   255
            Left            =   1650
            TabIndex        =   12
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VendedorEmissao 
            Height          =   255
            Left            =   390
            TabIndex        =   11
            Top             =   330
            Width           =   1735
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissoesEmissao 
            Height          =   870
            Left            =   75
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   210
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   1535
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label TotalPercentualEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1965
            TabIndex        =   66
            Top             =   1170
            Width           =   945
         End
         Begin VB.Label TotalValorEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3285
            TabIndex        =   65
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label LabelTotaisEmissao 
            Caption         =   "Totais:"
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
            Height          =   225
            Left            =   1200
            TabIndex        =   64
            Top             =   1185
            Width           =   705
         End
      End
      Begin MSMask.MaskEdBox Projeto 
         Height          =   285
         Left            =   1050
         TabIndex        =   222
         Top             =   1095
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProjeto 
         AutoSize        =   -1  'True
         Caption         =   "Projeto:"
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
         Left            =   285
         TabIndex        =   224
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Etapa:"
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
         Height          =   180
         Index           =   62
         Left            =   5925
         TabIndex        =   223
         Top             =   1155
         Width           =   570
      End
      Begin VB.Label LabelNatureza 
         AutoSize        =   -1  'True
         Caption         =   "Natureza:"
         BeginProperty Font 
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
         TabIndex        =   201
         Top             =   765
         Width           =   840
      End
      Begin VB.Label LabelNaturezaDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1050
         TabIndex        =   200
         Top             =   720
         Width           =   3360
      End
      Begin VB.Label LabelCcl 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6555
         TabIndex        =   199
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label CclLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo/Lucro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4485
         TabIndex        =   198
         Top             =   765
         Width           =   2010
      End
      Begin VB.Label Label49 
         Caption         =   "Reajustado Até:"
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
         Left            =   2250
         TabIndex        =   197
         Top             =   75
         Width           =   1440
      End
      Begin VB.Label ReajustadoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   3705
         TabIndex        =   196
         Top             =   -15
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   165
         TabIndex        =   158
         Top             =   75
         Width           =   555
      End
      Begin VB.Label Saldo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   795
         TabIndex        =   157
         Top             =   45
         Width           =   1245
      End
      Begin VB.Label JurosMensais 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6570
         TabIndex        =   108
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Juros Mensais:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5220
         TabIndex        =   107
         Top             =   405
         Width           =   1305
      End
      Begin VB.Label PercMulta 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6570
         TabIndex        =   106
         Top             =   -15
         Width           =   1095
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Multa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5955
         TabIndex        =   105
         Top             =   45
         Width           =   540
      End
      Begin VB.Label ValorTitulo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3705
         TabIndex        =   68
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Valor Original:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2415
         TabIndex        =   67
         Top             =   435
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1500
      Visible         =   0   'False
      Width           =   9930
      Begin VB.CommandButton BotaoDif 
         Caption         =   "Diferenças nas Parcelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   204
         ToolTipText     =   "Lista de Diferenças nas Parcelas"
         Top             =   4230
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificação da Parcela"
         Height          =   4695
         Left            =   60
         TabIndex        =   96
         Top             =   105
         Width           =   3210
         Begin MSComCtl2.UpDown UpDownParcela 
            Height          =   300
            Index           =   0
            Left            =   2145
            TabIndex        =   52
            Top             =   375
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Valor Reajustado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   203
            Top             =   2055
            Width           =   1530
         End
         Begin VB.Label ValorParc 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   202
            Top             =   2010
            Width           =   1185
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Nosso Número:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   185
            Top             =   3900
            Width           =   1305
         End
         Begin VB.Label NossoNumero 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   184
            Top             =   4170
            Width           =   2175
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Cobrador:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   810
            TabIndex        =   174
            Top             =   3195
            Width           =   840
         End
         Begin VB.Label Cobrador 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   173
            Top             =   3150
            Width           =   1185
         End
         Begin VB.Label Carteira 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   172
            Top             =   3540
            Width           =   1185
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Carteira:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   915
            TabIndex        =   171
            Top             =   3585
            Width           =   735
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1035
            TabIndex        =   166
            Top             =   2820
            Width           =   615
         End
         Begin VB.Label StatusParc 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   165
            Top             =   2775
            Width           =   1185
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Reajustado:"
            BeginProperty Font 
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
            TabIndex        =   164
            Top             =   2445
            Width           =   1575
         End
         Begin VB.Label SaldoParc 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1785
            TabIndex        =   163
            Top             =   2400
            Width           =   1185
         End
         Begin VB.Label ValorOriginalParc 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   104
            Top             =   1605
            Width           =   1185
         End
         Begin VB.Label DataVenctParcReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   103
            Top             =   1230
            Width           =   1185
         End
         Begin VB.Label DataVenctoParc 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   102
            Top             =   855
            Width           =   1185
         End
         Begin VB.Label Parcela 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   1770
            TabIndex        =   101
            Top             =   375
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "Parcela:"
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
            Left            =   945
            TabIndex        =   100
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Valor Original:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   435
            TabIndex        =   99
            Top             =   1650
            Width           =   1215
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Vencto Real:"
            BeginProperty Font 
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
            TabIndex        =   98
            Top             =   1275
            Width           =   1125
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   585
            TabIndex        =   97
            Top             =   900
            Width           =   1065
         End
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   1830
         Left            =   3360
         TabIndex        =   71
         Top             =   2040
         Width           =   6510
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   4860
            TabIndex        =   26
            Top             =   180
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   3540
            TabIndex        =   25
            Top             =   180
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   225
            Left            =   2220
            TabIndex        =   24
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   840
            TabIndex        =   23
            Top             =   180
            Width           =   1735
            _ExtentX        =   3069
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1125
            Left            =   45
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   300
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   1984
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label LabelTotaisComissoes 
            Caption         =   "Totais:"
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
            Height          =   225
            Left            =   525
            TabIndex        =   74
            Top             =   1530
            Width           =   705
         End
         Begin VB.Label TotalValorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2475
            TabIndex        =   73
            Top             =   1500
            Width           =   1155
         End
         Begin VB.Label TotalPercentualComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1380
            TabIndex        =   72
            Top             =   1500
            Width           =   945
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Descontos"
         Height          =   1845
         Left            =   3360
         TabIndex        =   75
         Top             =   120
         Width           =   6510
         Begin VB.ComboBox TipoDesconto 
            Height          =   315
            ItemData        =   "TituloReceber_ConsultaOcx.ctx":2F16
            Left            =   120
            List            =   "TituloReceber_ConsultaOcx.ctx":2F18
            TabIndex        =   18
            Top             =   300
            Width           =   1890
         End
         Begin MSMask.MaskEdBox Percentual1 
            Height          =   225
            Left            =   4170
            TabIndex        =   21
            Top             =   300
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   225
            Left            =   3120
            TabIndex        =   20
            Top             =   330
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   285
            Left            =   2070
            TabIndex        =   19
            Tag             =   "1"
            Top             =   300
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDescontos 
            Height          =   1110
            Left            =   60
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   270
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   1958
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5220
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   1485
      Visible         =   0   'False
      Width           =   9945
      Begin VB.CommandButton BotaoCTBBaixa 
         Caption         =   "Consultar contabilização da baixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Abre contabilização da baixa"
         Top             =   4365
         Width           =   2265
      End
      Begin VB.CommandButton BotaoImprimirRecibo 
         Caption         =   "Imprimir Recibo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Imprime o recibo"
         Top             =   4365
         Width           =   2055
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Cartão"
         Height          =   1770
         Index           =   4
         Left            =   2745
         TabIndex        =   205
         Top             =   2190
         Visible         =   0   'False
         Width           =   6105
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1185
            TabIndex        =   209
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Label52 
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
            Height          =   255
            Left            =   3750
            TabIndex        =   208
            Top             =   810
            Width           =   705
         End
         Begin VB.Label ClienteTitCartao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1935
            TabIndex        =   215
            Top             =   315
            Width           =   1740
         End
         Begin VB.Label NumTitCartao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4560
            TabIndex        =   214
            Top             =   780
            Width           =   720
         End
         Begin VB.Label FilialTitCartao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4560
            TabIndex        =   213
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label TipoTitCartao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1935
            TabIndex        =   212
            Top             =   780
            Width           =   1080
         End
         Begin VB.Label DataEmissaoTitCartao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1935
            TabIndex        =   211
            Top             =   1230
            Width           =   1095
         End
         Begin VB.Label Label55 
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
            Left            =   3990
            TabIndex        =   210
            Top             =   345
            Width           =   465
         End
         Begin VB.Label Label51 
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
            Height          =   255
            Left            =   1395
            TabIndex        =   207
            Top             =   810
            Width           =   450
         End
         Begin VB.Label Label47 
            Caption         =   "Data Emissão:"
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
            Left            =   600
            TabIndex        =   206
            Top             =   1260
            Width           =   1245
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Débito"
         Height          =   1770
         Index           =   2
         Left            =   2745
         TabIndex        =   141
         Top             =   2190
         Visible         =   0   'False
         Width           =   6105
         Begin VB.Label Label34 
            Caption         =   "Data Emissão:"
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
            Left            =   570
            TabIndex        =   153
            Top             =   405
            Width           =   1245
         End
         Begin VB.Label Label48 
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
            Height          =   255
            Left            =   4005
            TabIndex        =   152
            Top             =   405
            Width           =   450
         End
         Begin VB.Label Label37 
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
            Height          =   255
            Left            =   1110
            TabIndex        =   151
            Top             =   810
            Width           =   705
         End
         Begin VB.Label Label38 
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
            Height          =   255
            Left            =   3945
            TabIndex        =   150
            Top             =   810
            Width           =   510
         End
         Begin VB.Label Label39 
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
            Left            =   1260
            TabIndex        =   149
            Top             =   1215
            Width           =   555
         End
         Begin VB.Label Label40 
            Caption         =   "Filial Empresa:"
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
            Left            =   3210
            TabIndex        =   148
            Top             =   1215
            Width           =   1245
         End
         Begin VB.Label DataEmissaoCred 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DataEmissao"
            Height          =   300
            Left            =   1950
            TabIndex        =   147
            Top             =   382
            Width           =   1095
         End
         Begin VB.Label SiglaDocumentoCR 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SiglaDoc"
            Height          =   300
            Left            =   4560
            TabIndex        =   146
            Top             =   382
            Width           =   1080
         End
         Begin VB.Label ValorDebito 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor"
            Height          =   300
            Left            =   4560
            TabIndex        =   145
            Top             =   787
            Width           =   1080
         End
         Begin VB.Label FilialEmpresaCR 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FilEmpr"
            Height          =   300
            Left            =   4560
            TabIndex        =   144
            Top             =   1192
            Width           =   525
         End
         Begin VB.Label NumTitulo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Numero"
            Height          =   300
            Left            =   1950
            TabIndex        =   143
            Top             =   787
            Width           =   720
         End
         Begin VB.Label SaldoDebito 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo"
            Height          =   300
            Left            =   1950
            TabIndex        =   142
            Top             =   1192
            Width           =   1080
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Adiantamento de Cliente"
         Height          =   1770
         Index           =   1
         Left            =   2745
         TabIndex        =   132
         Top             =   2190
         Visible         =   0   'False
         Width           =   6105
         Begin VB.Label Label12 
            Caption         =   "Data Movimto:"
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
            Left            =   315
            TabIndex        =   140
            Top             =   585
            Width           =   1245
         End
         Begin VB.Label Label6 
            Caption         =   "Conta Corrente:"
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
            Left            =   3120
            TabIndex        =   139
            Top             =   585
            Width           =   1350
         End
         Begin VB.Label Label7 
            Caption         =   "Meio Pagto:"
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
            Left            =   525
            TabIndex        =   138
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label23 
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
            Height          =   255
            Left            =   3960
            TabIndex        =   137
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label DataMovimento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DataMovto"
            Height          =   300
            Left            =   1665
            TabIndex        =   136
            Top             =   555
            Width           =   1095
         End
         Begin VB.Label MeioPagtoDescricao 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MeioPagto"
            Height          =   300
            Left            =   1665
            TabIndex        =   135
            Top             =   1050
            Width           =   960
         End
         Begin VB.Label ValorPA 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ValorPagtoAnt"
            Height          =   300
            Left            =   4590
            TabIndex        =   134
            Top             =   1050
            Width           =   1080
         End
         Begin VB.Label CCIntNomeReduzido 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CCorrente"
            Height          =   300
            Left            =   4590
            TabIndex        =   133
            Top             =   555
            Width           =   1335
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Dados do Recebimento"
         Height          =   1770
         Index           =   0
         Left            =   2730
         TabIndex        =   125
         Top             =   2190
         Width           =   6105
         Begin VB.Label Label4 
            Caption         =   "Conta:"
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
            Left            =   315
            TabIndex        =   131
            Top             =   540
            Width           =   555
         End
         Begin VB.Label Label13 
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
            Height          =   255
            Left            =   2775
            TabIndex        =   130
            Top             =   540
            Width           =   810
         End
         Begin VB.Label Label25 
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
            Height          =   255
            Left            =   375
            TabIndex        =   129
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label ContaCorrente 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   960
            TabIndex        =   128
            Top             =   510
            Width           =   1590
         End
         Begin VB.Label ValorRecebimento 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   960
            TabIndex        =   127
            Top             =   997
            Width           =   1590
         End
         Begin VB.Label Historico 
            BorderStyle     =   1  'Fixed Single
            Height          =   720
            Left            =   3630
            TabIndex        =   126
            Top             =   510
            Width           =   2340
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Perda"
         Height          =   1770
         Index           =   3
         Left            =   2745
         TabIndex        =   154
         Top             =   2190
         Visible         =   0   'False
         Width           =   6105
         Begin VB.Label HistoricoPerda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1260
            TabIndex        =   156
            Top             =   675
            Width           =   4440
         End
         Begin VB.Label Label14 
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
            Height          =   255
            Left            =   390
            TabIndex        =   155
            Top             =   705
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados da Baixa"
         Height          =   1320
         Left            =   240
         TabIndex        =   112
         Top             =   750
         Width           =   8610
         Begin VB.Label DataBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1605
            TabIndex        =   124
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label2 
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
            Left            =   1005
            TabIndex        =   123
            Top             =   420
            Width           =   480
         End
         Begin VB.Label ValorPago 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4485
            TabIndex        =   122
            Top             =   345
            Width           =   945
         End
         Begin VB.Label Juros 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7170
            TabIndex        =   121
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Multa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4485
            TabIndex        =   120
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Desconto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1605
            TabIndex        =   119
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Valor Pago:"
            BeginProperty Font 
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
            TabIndex        =   118
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Juros:"
            BeginProperty Font 
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
            TabIndex        =   117
            Top             =   900
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Multa:"
            BeginProperty Font 
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
            Left            =   3885
            TabIndex        =   116
            Top             =   900
            Width           =   540
         End
         Begin VB.Label Label8 
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
            Left            =   600
            TabIndex        =   115
            Top             =   870
            Width           =   885
         End
         Begin VB.Label ValorBaixado 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7170
            TabIndex        =   114
            Top             =   360
            Width           =   945
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Valor Baixado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   113
            Top             =   420
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Baixa"
         Height          =   1770
         Left            =   240
         TabIndex        =   70
         Top             =   2190
         Width           =   2295
         Begin VB.OptionButton Recebimento 
            Caption         =   "Cartão"
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
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1515
            Width           =   2115
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Perda"
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
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1218
            Width           =   2115
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Débito / Devolução"
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
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   922
            Width           =   2115
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Adiantamento"
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
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   611
            Width           =   2145
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Recebimento"
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
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   300
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.ComboBox Sequencial 
         Height          =   315
         Left            =   3435
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   825
      End
      Begin MSComCtl2.UpDown UpDownParcela 
         Height          =   300
         Index           =   1
         Left            =   1635
         TabIndex        =   53
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label ValorParcBaixa 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5190
         TabIndex        =   170
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label35 
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
         Left            =   4665
         TabIndex        =   169
         Top             =   300
         Width           =   510
      End
      Begin VB.Label DataVencParcBaixa 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7590
         TabIndex        =   168
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6510
         TabIndex        =   167
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Parcela 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1245
         TabIndex        =   77
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
         BeginProperty Font 
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
         TabIndex        =   76
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   69
         Top             =   285
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5130
      Index           =   4
      Left            =   210
      TabIndex        =   37
      Top             =   1500
      Visible         =   0   'False
      Width           =   9840
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4560
         TabIndex        =   220
         Tag             =   "1"
         Top             =   1920
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoImprimir 
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
         Height          =   375
         Left            =   7560
         TabIndex        =   38
         Top             =   255
         Width           =   1245
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   930
         Left            =   45
         TabIndex        =   78
         Top             =   3150
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   82
            Top             =   585
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   81
            Top             =   255
            Width           =   3720
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Left            =   1125
            TabIndex        =   80
            Top             =   270
            Width           =   570
         End
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            TabIndex        =   79
            Top             =   600
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   5340
         MaxLength       =   150
         TabIndex        =   44
         Top             =   1050
         Width           =   3495
      End
      Begin VB.CheckBox CTBAglutina 
         Enabled         =   0   'False
         Height          =   210
         Left            =   7230
         TabIndex        =   45
         Top             =   1230
         Width           =   870
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4830
         TabIndex        =   43
         Top             =   1140
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBConta 
         Bindings        =   "TituloReceber_ConsultaOcx.ctx":2F1A
         Height          =   225
         Left            =   210
         TabIndex        =   39
         Top             =   1230
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3570
         TabIndex        =   42
         Top             =   1200
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2370
         TabIndex        =   41
         Top             =   1230
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1440
         TabIndex        =   40
         Top             =   1200
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   30
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1005
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label CTBLote 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5970
         TabIndex        =   111
         Top             =   75
         Width           =   615
      End
      Begin VB.Label CTBDocumento 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3705
         TabIndex        =   110
         Top             =   75
         Width           =   705
      End
      Begin VB.Label CTBDataContabil 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   750
         TabIndex        =   109
         Top             =   442
         Width           =   1095
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5460
         TabIndex        =   95
         Top             =   135
         Width           =   450
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2640
         TabIndex        =   94
         Top             =   135
         Width           =   1035
      End
      Begin VB.Label CTBLabel8 
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
         Left            =   225
         TabIndex        =   93
         Top             =   495
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   92
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   91
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
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
         Height          =   225
         Left            =   1800
         TabIndex        =   90
         Top             =   2865
         Width           =   615
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   89
         Top             =   825
         Width           =   1140
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   2775
         TabIndex        =   88
         Top             =   465
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3705
         TabIndex        =   87
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5970
         TabIndex        =   86
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   5160
         TabIndex        =   85
         Top             =   465
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   84
         Top             =   90
         Width           =   1530
      End
      Begin VB.Label CTBLabel21 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   45
         TabIndex        =   83
         Top             =   135
         Width           =   720
      End
   End
   Begin VB.PictureBox FilialEmpresa 
      Height          =   300
      Left            =   6990
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   218
      Top             =   375
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   1035
      TabIndex        =   2
      Top             =   375
      Width           =   2400
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   45
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   7500
      ScaleHeight     =   795
      ScaleWidth      =   2565
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   120
      Width           =   2625
      Begin VB.CommandButton BotaoConsulta 
         Height          =   675
         Left            =   90
         Picture         =   "TituloReceber_ConsultaOcx.ctx":2F25
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   60
         Width           =   1275
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   675
         Left            =   1455
         Picture         =   "TituloReceber_ConsultaOcx.ctx":4CE7
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   480
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   675
         Left            =   2025
         Picture         =   "TituloReceber_ConsultaOcx.ctx":5219
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   480
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5700
      Left            =   90
      TabIndex        =   51
      Top             =   1080
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   10054
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Título"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parcelas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Baixa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
   Begin MSMask.MaskEdBox NumeroTitulo 
      Height          =   300
      Left            =   4680
      TabIndex        =   3
      Top             =   390
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "99999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   1035
      TabIndex        =   0
      Top             =   45
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LabelFilialEmpresa 
      AutoSize        =   -1  'True
      Caption         =   "FilialEmpresa:"
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
      Left            =   5790
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   219
      Top             =   420
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label16 
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
      Height          =   195
      Left            =   180
      TabIndex        =   187
      Top             =   750
      Width           =   795
   End
   Begin VB.Label DataEmissao 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1035
      TabIndex        =   186
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
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
      Height          =   165
      Left            =   4020
      TabIndex        =   160
      Top             =   765
      Width           =   615
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4680
      TabIndex        =   159
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   -15
      X2              =   10200
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label TipoDocumentoLabel 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   540
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   57
      Top             =   435
      Width           =   450
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4155
      TabIndex        =   56
      Top             =   90
      Width           =   465
   End
   Begin VB.Label NumeroLabel 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3900
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   55
      Top             =   435
      Width           =   720
   End
   Begin VB.Label ClienteEtiqueta 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   330
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   54
      Top             =   75
      Width           =   660
   End
End
Attribute VB_Name = "TituloReceber_ConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTTituloRec_Consulta
Attribute objCT.VB_VarHelpID = -1

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoConsulta_Click()
     Call objCT.BotaoConsulta_Click
End Sub

Private Sub BotaoDocOriginal_Click()
     Call objCT.BotaoDocOriginal_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub NumeroTitulo_GotFocus()
     Call objCT.NumeroTitulo_GotFocus
End Sub

Private Sub ClienteEtiqueta_Click()
     Call objCT.ClienteEtiqueta_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Function Trata_Parametros(Optional objTituloReceber As ClassTituloReceber) As Long
     Trata_Parametros = objCT.Trata_Parametros(objTituloReceber)
End Function

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub Sequencial_Click()
     Call objCT.Sequencial_Click
End Sub

Private Sub TipoDocumentoLabel_Click()
     Call objCT.TipoDocumentoLabel_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub NumeroTitulo_Validate(Cancel As Boolean)
     Call objCT.NumeroTitulo_Validate(Cancel)
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub UpDownParcela_DownClick(Index As Integer)
     Call objCT.UpDownParcela_DownClick(Index)
End Sub

Private Sub UpDownParcela_UpClick(Index As Integer)
     Call objCT.UpDownParcela_UpClick(Index)
End Sub

Private Sub Recebimento_Click(Index As Integer)
     Call objCT.Recebimento_Click(Index)
End Sub

Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
        If Cancel = False Then
            Set objCT.objUserControl = Nothing
            Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Private Sub UserControl_Initialize()
    Set objCT = New CTTituloRec_Consulta
    Set objCT.objUserControl = Me
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub



Private Sub Parcela_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Parcela(Index), Source, X, Y)
End Sub

Private Sub Parcela_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Parcela(Index), Button, Shift, X, Y)
End Sub


Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub ContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCorrente, Source, X, Y)
End Sub

Private Sub ContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCorrente, Button, Shift, X, Y)
End Sub

Private Sub ValorRecebimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorRecebimento, Source, X, Y)
End Sub

Private Sub ValorRecebimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorRecebimento, Button, Shift, X, Y)
End Sub

Private Sub Historico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Historico, Source, X, Y)
End Sub

Private Sub Historico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Historico, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub DataMovimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataMovimento, Source, X, Y)
End Sub

Private Sub DataMovimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataMovimento, Button, Shift, X, Y)
End Sub

Private Sub MeioPagtoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MeioPagtoDescricao, Source, X, Y)
End Sub

Private Sub MeioPagtoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MeioPagtoDescricao, Button, Shift, X, Y)
End Sub

Private Sub ValorPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPA, Source, X, Y)
End Sub

Private Sub ValorPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPA, Button, Shift, X, Y)
End Sub

Private Sub CCIntNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CCIntNomeReduzido, Source, X, Y)
End Sub

Private Sub CCIntNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CCIntNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label38_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label38, Source, X, Y)
End Sub

Private Sub Label38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label38, Button, Shift, X, Y)
End Sub

Private Sub Label39_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label39, Source, X, Y)
End Sub

Private Sub Label39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label39, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Private Sub DataEmissaoCred_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissaoCred, Source, X, Y)
End Sub

Private Sub DataEmissaoCred_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissaoCred, Button, Shift, X, Y)
End Sub

Private Sub SiglaDocumentoCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SiglaDocumentoCR, Source, X, Y)
End Sub

Private Sub SiglaDocumentoCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SiglaDocumentoCR, Button, Shift, X, Y)
End Sub

Private Sub ValorDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDebito, Source, X, Y)
End Sub

Private Sub ValorDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDebito, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaCR, Source, X, Y)
End Sub

Private Sub FilialEmpresaCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaCR, Button, Shift, X, Y)
End Sub

Private Sub NumTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumTitulo, Source, X, Y)
End Sub

Private Sub NumTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumTitulo, Button, Shift, X, Y)
End Sub

Private Sub SaldoDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoDebito, Source, X, Y)
End Sub

Private Sub SaldoDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoDebito, Button, Shift, X, Y)
End Sub

Private Sub HistoricoPerda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(HistoricoPerda, Source, X, Y)
End Sub

Private Sub HistoricoPerda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(HistoricoPerda, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub DataBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataBaixa, Source, X, Y)
End Sub

Private Sub DataBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub ValorPago_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPago, Source, X, Y)
End Sub

Private Sub ValorPago_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPago, Button, Shift, X, Y)
End Sub

Private Sub Juros_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Juros, Source, X, Y)
End Sub

Private Sub Juros_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Juros, Button, Shift, X, Y)
End Sub

Private Sub Multa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Multa, Source, X, Y)
End Sub

Private Sub Multa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Multa, Button, Shift, X, Y)
End Sub

Private Sub Desconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Desconto, Source, X, Y)
End Sub

Private Sub Desconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Desconto, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ValorBaixado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorBaixado, Source, X, Y)
End Sub

Private Sub ValorBaixado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorBaixado, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub ValorParcBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorParcBaixa, Source, X, Y)
End Sub

Private Sub ValorParcBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorParcBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub DataVencParcBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataVencParcBaixa, Source, X, Y)
End Sub

Private Sub DataVencParcBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataVencParcBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label42_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label42, Source, X, Y)
End Sub

Private Sub Label42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label42, Button, Shift, X, Y)
End Sub

Private Sub Cobrador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cobrador, Source, X, Y)
End Sub

Private Sub Cobrador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cobrador, Button, Shift, X, Y)
End Sub

Private Sub Carteira_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Carteira, Source, X, Y)
End Sub

Private Sub Carteira_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Carteira, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub StatusParc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(StatusParc, Source, X, Y)
End Sub

Private Sub StatusParc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(StatusParc, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub SaldoParc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoParc, Source, X, Y)
End Sub

Private Sub SaldoParc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoParc, Button, Shift, X, Y)
End Sub

Private Sub ValorParc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorParc, Source, X, Y)
End Sub

Private Sub ValorParc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorParc, Button, Shift, X, Y)
End Sub

Private Sub DataVenctParcReal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataVenctParcReal, Source, X, Y)
End Sub

Private Sub DataVenctParcReal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataVenctParcReal, Button, Shift, X, Y)
End Sub

Private Sub DataVenctoParc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataVenctoParc, Source, X, Y)
End Sub

Private Sub DataVenctoParc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataVenctoParc, Button, Shift, X, Y)
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

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub LabelTotaisComissoes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotaisComissoes, Source, X, Y)
End Sub

Private Sub LabelTotaisComissoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotaisComissoes, Button, Shift, X, Y)
End Sub

Private Sub TotalValorComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorComissao, Source, X, Y)
End Sub

Private Sub TotalValorComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorComissao, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentualComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualComissao, Source, X, Y)
End Sub

Private Sub TotalPercentualComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualComissao, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub ValorINSS_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorINSS, Source, X, Y)
End Sub

Private Sub ValorINSS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorINSS, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub ValorIRRF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorIRRF, Source, X, Y)
End Sub

Private Sub ValorIRRF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorIRRF, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagamento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagamento, Source, X, Y)
End Sub

Private Sub CondicaoPagamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagamento, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentualEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualEmissao, Source, X, Y)
End Sub

Private Sub TotalPercentualEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualEmissao, Button, Shift, X, Y)
End Sub

Private Sub TotalValorEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorEmissao, Source, X, Y)
End Sub

Private Sub TotalValorEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorEmissao, Button, Shift, X, Y)
End Sub

Private Sub LabelTotaisEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotaisEmissao, Source, X, Y)
End Sub

Private Sub LabelTotaisEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotaisEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Saldo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Saldo, Source, X, Y)
End Sub

Private Sub Saldo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Saldo, Button, Shift, X, Y)
End Sub

Private Sub JurosMensais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(JurosMensais, Source, X, Y)
End Sub

Private Sub JurosMensais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(JurosMensais, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub PercMulta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercMulta, Source, X, Y)
End Sub

Private Sub PercMulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercMulta, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub ValorTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTitulo, Source, X, Y)
End Sub

Private Sub ValorTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTitulo, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLote, Source, X, Y)
End Sub

Private Sub CTBLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLote, Button, Shift, X, Y)
End Sub

Private Sub CTBDocumento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBDocumento, Source, X, Y)
End Sub

Private Sub CTBDocumento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBDocumento, Button, Shift, X, Y)
End Sub

Private Sub CTBDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBDataContabil, Source, X, Y)
End Sub

Private Sub CTBDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBDataContabil, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub TipoDocumentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoDocumentoLabel, Source, X, Y)
End Sub

Private Sub TipoDocumentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoDocumentoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub ClienteEtiqueta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteEtiqueta, Source, X, Y)
End Sub

Private Sub ClienteEtiqueta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteEtiqueta, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub ValorOriginalParc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorOriginalParc, Source, X, Y)
End Sub

Private Sub ValorOriginalParc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorOriginalParc, Button, Shift, X, Y)
End Sub

Private Sub BotaoDif_Click()
    Call objCT.BotaoDif_Click
End Sub

'#####################################
'Inserido por Wagner 21/08/2006
Public Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub
'#####################################

'#####################################
'Inserido por Wagner 09/11/2006
Private Sub BotaoImprimirRecibo_Click()
    Call objCT.BotaoImprimirRecibo_Click
End Sub
'#####################################


Private Sub CTBGerencial_Click()
    Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
    Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
    Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
    Call objCT.CTBGerencial_Validate(Cancel)
End Sub

Private Sub BotaoCTBBaixa_Click()
    Call objCT.BotaoCTBBaixa_Click
End Sub
