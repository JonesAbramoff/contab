VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVOcorrencias 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ForeColor       =   &H00000080&
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame4 
      Caption         =   "Detalhamento"
      Height          =   3180
      Left            =   5625
      TabIndex        =   46
      Top             =   1575
      Width           =   3855
      Begin VB.ComboBox Tipo 
         Height          =   315
         ItemData        =   "TRVOcorrencias.ctx":0000
         Left            =   390
         List            =   "TRVOcorrencias.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   765
         Width           =   2115
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Left            =   2475
         TabIndex        =   48
         Top             =   750
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridDet 
         Height          =   1545
         Left            =   45
         TabIndex        =   15
         Top             =   225
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   2725
         _Version        =   393216
         Rows            =   15
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2430
         TabIndex        =   50
         Top             =   2775
         Width           =   1065
      End
      Begin VB.Label LabelValor 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor a Faturar:"
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
         Height          =   315
         Left            =   915
         TabIndex        =   49
         Top             =   2820
         Width           =   1470
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Outros"
      Height          =   1245
      Left            =   30
      TabIndex        =   41
      Top             =   4725
      Width           =   9450
      Begin VB.TextBox Observacao 
         Height          =   630
         Left            =   945
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   540
         Width           =   8415
      End
      Begin VB.ComboBox Historico 
         Height          =   315
         Left            =   945
         TabIndex        =   20
         Top             =   195
         Width           =   8430
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   43
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   42
         Top             =   555
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Voucher vinculado"
      Height          =   1020
      Left            =   30
      TabIndex        =   33
      Top             =   525
      Width           =   9450
      Begin VB.CommandButton BotaoHistOcor 
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
         Height          =   360
         Left            =   9030
         TabIndex        =   4
         ToolTipText     =   "Cancela o voucher"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton BotaoTrazerVou 
         Height          =   330
         Left            =   4200
         Picture         =   "TRVOcorrencias.ctx":0041
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Trazer Dados"
         Top             =   210
         Width           =   360
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
         Left            =   5205
         TabIndex        =   5
         Top             =   210
         Visible         =   0   'False
         Width           =   1050
      End
      Begin MSMask.MaskEdBox TipoVou 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Top             =   210
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SerieVou 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   225
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumeroVou 
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Top             =   225
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label ValorTotalTodasOcr 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   8085
         TabIndex        =   57
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Valor total de todas Ocr:"
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
         Left            =   5955
         TabIndex        =   56
         Top             =   285
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Neto:"
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
         Index           =   4
         Left            =   4365
         TabIndex        =   45
         Top             =   675
         Width           =   1020
      End
      Begin VB.Label ValorVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5400
         TabIndex        =   44
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label DataEmissaoVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   8085
         TabIndex        =   40
         Top             =   615
         Width           =   1290
      End
      Begin VB.Label ClienteVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   915
         TabIndex        =   39
         Top             =   600
         Width           =   3285
      End
      Begin VB.Label Label1 
         Caption         =   "Data de Emissão:"
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
         Index           =   5
         Left            =   6555
         TabIndex        =   38
         Top             =   675
         Width           =   1620
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
         Left            =   225
         TabIndex        =   37
         Top             =   660
         Width           =   630
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
         Left            =   2340
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   240
         Width           =   750
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
         Left            =   1350
         TabIndex        =   35
         Top             =   255
         Width           =   480
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
         Left            =   420
         TabIndex        =   34
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   3165
      Left            =   30
      TabIndex        =   27
      Top             =   1575
      Width           =   5565
      Begin VB.Frame FrameResumo 
         Height          =   1485
         Left            =   30
         TabIndex        =   58
         Top             =   1635
         Width           =   5490
         Begin MSMask.MaskEdBox ValorBrutoVouNovo 
            Height          =   315
            Left            =   1605
            TabIndex        =   16
            Top             =   795
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox ValorBrutoVouOCR 
            Height          =   315
            Left            =   1605
            TabIndex        =   18
            Top             =   1125
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox ValorCMAVouNovo 
            Height          =   315
            Left            =   2610
            TabIndex        =   17
            Top             =   795
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox ValorCMAVouOCR 
            Height          =   315
            Left            =   2610
            TabIndex        =   19
            Top             =   1125
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
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
         Begin VB.Label FatorOVER 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4695
            TabIndex        =   74
            Top             =   1125
            Width           =   765
         End
         Begin VB.Label FatorCMC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4695
            TabIndex        =   73
            Top             =   795
            Width           =   765
         End
         Begin VB.Label FatorCMR 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4695
            TabIndex        =   72
            Top             =   450
            Width           =   765
         End
         Begin VB.Label FatorCMCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4695
            TabIndex        =   71
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Fator CMR:"
            BeginProperty Font 
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
            Left            =   3735
            TabIndex        =   70
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Fator CMC:"
            BeginProperty Font 
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
            Left            =   3735
            TabIndex        =   69
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Fator OVER:"
            BeginProperty Font 
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
            Left            =   3630
            TabIndex        =   68
            Top             =   1185
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Fator CMCC:"
            BeginProperty Font 
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
            Left            =   3630
            TabIndex        =   67
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "RESUMO"
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
            Index           =   19
            Left            =   465
            TabIndex        =   66
            Top             =   180
            Width           =   930
         End
         Begin VB.Line Line6 
            X1              =   -285
            X2              =   3600
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Line Line5 
            X1              =   -285
            X2              =   3600
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Line Line4 
            X1              =   -285
            X2              =   3600
            Y1              =   435
            Y2              =   435
         End
         Begin VB.Label ValorCMAVouAtual 
            Height          =   330
            Left            =   2610
            TabIndex        =   65
            Top             =   450
            Width           =   990
         End
         Begin VB.Label ValorBrutoVouAtual 
            Height          =   330
            Left            =   1605
            TabIndex        =   64
            Top             =   450
            Width           =   990
         End
         Begin VB.Line Line3 
            X1              =   3600
            X2              =   3600
            Y1              =   105
            Y2              =   1480
         End
         Begin VB.Line Line2 
            X1              =   2595
            X2              =   2595
            Y1              =   105
            Y2              =   1480
         End
         Begin VB.Line Line1 
            X1              =   1590
            X2              =   1590
            Y1              =   105
            Y2              =   1480
         End
         Begin VB.Label Label1 
            Caption         =   "CMA"
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
            Left            =   2895
            TabIndex        =   63
            Top             =   135
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "BRUTO"
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
            Left            =   1785
            TabIndex        =   62
            Top             =   135
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Ajustes da OCR:"
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
            Index           =   16
            Left            =   180
            TabIndex        =   61
            Top             =   1185
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Vou Vlr Esperado:"
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
            Index           =   15
            Left            =   30
            TabIndex        =   60
            Top             =   855
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Vou Valor Atual:"
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
            Index           =   14
            Left            =   180
            TabIndex        =   59
            Top             =   540
            Width           =   1425
         End
      End
      Begin VB.CommandButton BotaoAbrirDoc 
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
         Left            =   4770
         TabIndex        =   14
         Top             =   1305
         Width           =   660
      End
      Begin VB.ComboBox FormaPagto 
         Height          =   315
         ItemData        =   "TRVOcorrencias.ctx":0413
         Left            =   900
         List            =   "TRVOcorrencias.ctx":041D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   3405
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1770
         Picture         =   "TRVOcorrencias.ctx":043F
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Numeração Automática"
         Top             =   240
         Width           =   300
      End
      Begin VB.ComboBox Origem 
         Height          =   315
         ItemData        =   "TRVOcorrencias.ctx":0529
         Left            =   3405
         List            =   "TRVOcorrencias.ctx":0533
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   615
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Top             =   225
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   315
         Left            =   900
         TabIndex        =   8
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataEmissao 
         Height          =   300
         Left            =   2205
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   945
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc:"
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
         Index           =   12
         Left            =   2010
         TabIndex        =   55
         Top             =   1350
         Width           =   1395
      End
      Begin VB.Label NumDocDestino 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3420
         TabIndex        =   54
         Top             =   1305
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Forma:"
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
         Height          =   315
         Index           =   8
         Left            =   150
         TabIndex        =   53
         Top             =   1365
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Index           =   13
         Left            =   2910
         TabIndex        =   52
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label LabelCliente 
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   51
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3420
         TabIndex        =   31
         Top             =   225
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   7
         Left            =   2445
         TabIndex        =   30
         Top             =   645
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   11
         Left            =   1980
         TabIndex        =   29
         Top             =   255
         Width           =   1395
      End
      Begin VB.Label LabelCodigo 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   60
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7320
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TRVOcorrencias.ctx":0564
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "TRVOcorrencias.ctx":06BE
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "TRVOcorrencias.ctx":0848
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "TRVOcorrencias.ctx":0D7A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "TRVOcorrencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoVoucher As AdmEvento
Attribute objEventoVoucher.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim lNumvouAnt As Long
Dim sSerieAnt As String
Dim sTipoAnt As String
Dim iOrigemAnt As Integer
Dim lCodigoAnt As Long
Dim dValorAnt As Double

'GridCliente
Dim objGridDet As AdmGrid
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Valor_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Ocorrências"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVOcorrencias"

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

Private Sub BotaoTrazerVou_Click()
    Call TrazerVou_Click
End Sub

Private Sub TrazerVou_Click(Optional ByVal bValidaCanc As Boolean = True)

Dim lErro As Long
Dim objVoucher As New ClassTRVVouchers
Dim objTRVVoucherInfo As New ClassTRVVoucherInfo

On Error GoTo Erro_TrazerVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVou.Text)
    objVoucher.sSerie = SerieVou.Text
    objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVoucher.sTipVou = TipoVou.Text
    
    lErro = CF("TRVVouchers_Le", objVoucher)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194421
    
    If lErro <> SUCESSO Then gError 194425
    
    If objVoucher.iStatus = STATUS_TRV_VOU_CANCELADO And bValidaCanc Then gError 194422
    
    'Preenche campo Cliente
    Cliente.Text = objVoucher.lCliente

    'Executa o Validate
    Call Cliente_Validate(bSGECancelDummy)
    
    ClienteVou.Caption = Cliente.Text
    ValorVou.Caption = Format(objVoucher.dValor, "STANDARD")
    DataEmissaoVou.Caption = Format(objVoucher.dtData, "dd/mm/yyyy")
    ValorTotalTodasOcr.Caption = Format(objVoucher.dValorOcr, "STANDARD")
    
    If objVoucher.iCartao = MARCADO Then
    
        objTRVVoucherInfo.sSerie = objVoucher.sSerie
        objTRVVoucherInfo.sTipo = objVoucher.sTipVou
        objTRVVoucherInfo.lNumVou = objVoucher.lNumVou
        
        'Lê o TRVVouchers que está sendo Passado
        lErro = CF("TRVVoucherInfoSigav_Le", objTRVVoucherInfo)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194423
        
        If lErro <> SUCESSO Then gError 194424
        
        If objTRVVoucherInfo.lCliente <> 0 Then
        
            'Preenche campo Cliente
            Cliente.Text = objTRVVoucherInfo.lCliente
        
            'Executa o Validate
            Call Cliente_Validate(bSGECancelDummy)
            
        End If
        
    End If
    
    lNumvouAnt = objVoucher.lNumVou
    sSerieAnt = objVoucher.sSerie
    sTipoAnt = objVoucher.sTipVou

    Exit Sub

Erro_TrazerVou_Click:

    Select Case gErr
    
        Case 194421, 194423
        
        Case 194422
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
        Case 194424
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_SEM_DADOS_SIGAV", gErr)
            
        Case 194425
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVoucher.lNumVou, objVoucher.sSerie, objVoucher.sTipVou)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194426)

    End Select

    Exit Sub
    
End Sub

Private Sub FormaPagto_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelNumVou_Click()
    Call BotaoVou_Click
End Sub


Private Sub NumeroVou_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumeroVou_Validate(Cancel As Boolean)
    If StrParaLong(NumeroVou.Text) <> lNumvouAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub objEventoVoucher_evSelecao(obj1 As Object)

Dim objVoucher As ClassTRVVouchers
Dim bCancel As Boolean

    Set objVoucher = obj1

    'Preenche campo Cliente
    'Cliente.Text = objVoucher.lCliente

    'Executa o Validate
    'Call Cliente_Validate(bCancel)

'    TipoVou.Caption = objVoucher.sTipVou
'    SerieVou.Caption = objVoucher.sSerie
'    NumeroVou.Caption = objVoucher.lNumVou
    TipoVou.Text = objVoucher.sTipVou
    SerieVou.Text = objVoucher.sSerie
    
    NumeroVou.PromptInclude = False
    NumeroVou.Text = objVoucher.lNumVou
    NumeroVou.PromptInclude = True
    
    Call TrazerVou_Click
    
'    ClienteVou.Caption = Cliente.Text
'    ValorVou.Caption = Format(objVoucher.dvalor, "STANDARD")
'    DataEmissaoVou.Caption = Format(objVoucher.dtdata, "dd/mm/yyyy")

    iAlterado = REGISTRO_ALTERADO

    Me.Show

    Exit Sub

End Sub

Private Sub Limpa_Vou()

    ClienteVou.Caption = ""
    ValorVou.Caption = ""
    DataEmissaoVou.Caption = ""
    ValorTotalTodasOcr.Caption = ""

End Sub

Private Sub BotaoVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRVVouchers
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVou_Click

'    objVoucher.lNumVou = StrParaLong(NumeroVou.Caption)
'    objVoucher.sSerie = SerieVou.Caption
'    objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
'    objVoucher.sTipVou = TipoVou.Caption
    
    objVoucher.lNumVou = StrParaLong(NumeroVou.Text)
    objVoucher.sSerie = SerieVou.Text
    objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVoucher.sTipVou = TipoVou.Text

    objVoucher.lCliente = LCodigo_Extrai(ClienteVou.Caption)

    'colSelecao.Add STATUS_TRV_VOU_CANCELADO

    Call Chama_Tela("VoucherRapidoLista", colSelecao, objVoucher, objEventoVoucher, "Cancelado <> 'Sim'")

    Exit Sub

Erro_BotaoVou_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190160)

    End Select

    Exit Sub

End Sub

Private Sub Origem_Click()
    If iOrigemAnt <> Codigo_Extrai(Origem.Text) Then
        Call OCR_Ajusta_Resumo
        iOrigemAnt = Codigo_Extrai(Origem.Text)
    End If
End Sub

Private Sub SerieVou_Validate(Cancel As Boolean)
    If SerieVou.Text <> sSerieAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub TipoVou_Validate(Cancel As Boolean)
    If TipoVou.Text <> sTipoAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub TipoVou_Change()
    iAlterado = REGISTRO_ALTERADO
    If Len(Trim(TipoVou.ClipText)) > 0 Then
        If SerieVou.Visible Then SerieVou.SetFocus
    End If
End Sub

Private Sub SerieVou_Change()
    iAlterado = REGISTRO_ALTERADO
    If Len(Trim(SerieVou.ClipText)) > 0 Then
        If NumeroVou.Visible Then NumeroVou.SetFocus
    End If
End Sub

Private Sub TipoVou_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SerieVou_KeyPress(KeyAscii As Integer)
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

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

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

    Set objEventoCodigo = Nothing
    Set objGridDet = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190161)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoVoucher = New AdmEvento
    Set objEventoCliente = New AdmEvento

    Set objGridDet = New AdmGrid
    
    lErro = Inicializa_Grid_Det(objGridDet)
    If lErro <> SUCESSO Then gError 190162
    
    'lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRVORIGEM_OCR, Origem)
    lErro = CF("TRVCarrega_TipoOrigemOCR", Origem)
    If lErro <> SUCESSO Then gError 190163
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRVTIPODET_OCR, Tipo)
    If lErro <> SUCESSO Then gError 190164
    
    Historico.Clear
    lErro = CF("Carrega_Combo_Historico", Historico, "TRVOcorrencias", STRING_TRV_OCR_HISTORICO)
    If lErro <> SUCESSO Then gError 190165
    
    lErro = CF("Carrega_Combo_FormaPagto", FormaPagto)
    If lErro <> SUCESSO Then gError 190745
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 190162 To 190165, 190745

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190166)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTRVOcorrencias Is Nothing) Then

        lErro = Traz_TRVOcorrencias_Tela(objTRVOcorrencias)
        If lErro <> SUCESSO Then gError 190167

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 190167

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190168)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objTRVOcrDet As ClassTRVOcorrenciaDet
Dim iLinha As Integer
Dim objVou As New ClassTRVVouchers

On Error GoTo Erro_Move_Tela_Memoria

    objcliente.sNomeReduzido = Cliente.Text

    'Lê o Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 190169
    
    objTRVOcorrencias.lCodigo = StrParaLong(Codigo.Text)
    objTRVOcorrencias.lCliente = objcliente.lCodigo
    objTRVOcorrencias.iFilialCliente = Codigo_Extrai(Filial.Text)
    objTRVOcorrencias.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objTRVOcorrencias.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objTRVOcorrencias.sObservacao = Observacao.Text
    objTRVOcorrencias.iOrigem = Codigo_Extrai(Origem.Text)
    objTRVOcorrencias.sHistorico = Historico.Text
    objTRVOcorrencias.iStatus = STATUS_TRV_OCR_BLOQUEADO
    objTRVOcorrencias.iFormaPagto = Codigo_Extrai(FormaPagto.Text)
    objTRVOcorrencias.lNumVou = StrParaLong(NumeroVou.Text)
    objTRVOcorrencias.sSerie = SerieVou.Text
    objTRVOcorrencias.sTipoDoc = TipoVou.Text
    objTRVOcorrencias.dValorOCRBrutoVou = StrParaDbl(ValorBrutoVouOCR.Text)
    objTRVOcorrencias.dValorOCRCMAVou = StrParaDbl(ValorCMAVouOCR.Text)
    
    objVou.sTipVou = objTRVOcorrencias.sTipoDoc
    objVou.sSerie = objTRVOcorrencias.sSerie
    objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVou.lNumVou = objTRVOcorrencias.lNumVou

    lErro = CF("TRVVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194222
    
    If lErro <> SUCESSO Then gError 194223
    
    For iLinha = 1 To objGridDet.iLinhasExistentes
    
        Set objTRVOcrDet = New ClassTRVOcorrenciaDet

        objTRVOcrDet.iTipo = Codigo_Extrai(GridDet.TextMatrix(iLinha, iGrid_Tipo_Col))
        objTRVOcrDet.iSeq = iLinha
        objTRVOcrDet.dValor = StrParaDbl(GridDet.TextMatrix(iLinha, iGrid_Valor_Col))
        
        objTRVOcorrencias.colDetalhamento.Add objTRVOcrDet
        
    Next
    
    If objTRVOcorrencias.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
        If objVou.iStatus <> STATUS_TRV_VOU_CANCELADO Then gError 198092
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 190169, 194222
        
        Case 194223
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)

        Case 198092
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NVL_VOU_NAO_CANC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190170)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRVOcorrencias"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", StrParaLong(Codigo.Text), 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 190171

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190172)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTRVOcorrencias As New ClassTRVOcorrencias

On Error GoTo Erro_Tela_Preenche

    objTRVOcorrencias.lCodigo = colCampoValor.Item("Codigo").vValor

    If objTRVOcorrencias.lCodigo <> 0 Then
    
        lErro = Traz_TRVOcorrencias_Tela(objTRVOcorrencias)
        If lErro <> SUCESSO Then gError 190173
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 190173

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190174)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objTRVOcorrencias As New ClassTRVOcorrencias
Dim objTRVOcrDet As ClassTRVOcorrenciaDet
Dim objTipoOcr As New ClassTRVTiposOcorrencia

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 190175
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 190176
    If Len(Trim(Origem.Text)) = 0 Then gError 190177
    If Len(Trim(Cliente.Text)) = 0 Then gError 190178
    If Len(Trim(Filial.Text)) = 0 Then gError 190179
    If Len(Trim(FormaPagto.Text)) = 0 Then gError 190180
    '#####################
    
    'Preenche o objTRVOcorrencias
    lErro = Move_Tela_Memoria(objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 190187
    
    If objTRVOcorrencias.colDetalhamento.Count = 0 Then gError 190181
       
    objTipoOcr.iCodigo = objTRVOcorrencias.iOrigem
    
    lErro = CF("TRVTiposOcorrencia_Le", objTipoOcr)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198085
    
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 198088
    
    If objTRVOcorrencias.dValorTotal > 0 Then
        If objTipoOcr.iAceitaVlrPositivo = DESMARCADO Then gError 198086
    Else
        If objTipoOcr.iAceitaVlrNegativo = DESMARCADO Then gError 198087
    End If
    
    iLinha = 0
    For Each objTRVOcrDet In objTRVOcorrencias.colDetalhamento
        iLinha = iLinha + 1
        If objTRVOcrDet.iTipo = 0 Then gError 190182
        'If objTRVOcrDet.dValor = 0 Then gError 190183
    Next

    lErro = Trata_Alteracao(objTRVOcorrencias, objTRVOcorrencias.lCodigo)
    If lErro <> SUCESSO Then gError 190184

    'Grava o/a TRVOcorrencias no Banco de Dados
    lErro = CF("TRVOcorrencias_Grava", objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 190185
    
    Historico.Clear
    lErro = CF("Carrega_Combo_Historico", Historico, "TRVOcorrencias", STRING_TRV_OCR_HISTORICO)
    If lErro <> SUCESSO Then gError 190186

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 190175 'Codigo
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 190176 'DataEmissao
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDO", gErr)
            DataEmissao.SetFocus

        Case 190177 'Origem
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA_TRV", gErr)
            Origem.SetFocus

        Case 190178 'Cliente
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            Cliente.SetFocus

        Case 190179 'Filial
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            Filial.SetFocus
            
        Case 190180 'FormaPagto
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_PREENCHIDA", gErr)
            FormaPagto.SetFocus
            
        Case 190181
            Call Rotina_Erro(vbOKOnly, "ERRO_DETALHAMENTO_NAO_PREENCHIDO", gErr)

        Case 190182
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_GRID_NAO_PREENCHIDO", gErr, iLinha)

        Case 190183
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_GRID_NAO_PREENCHIDO", gErr, iLinha)

        Case 190184 To 190187, 198085

        Case 198086
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVTIPOOCR_NAO_ACEITA_VLR_POSITIVO", gErr)

        Case 198087
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVTIPOOCR_NAO_ACEITA_VLR_NEGATIVO", gErr)
            
        Case 198088
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRVTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190188)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRVOcorrencias() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVOcorrencias

    Call Limpa_Vou

    ValorTotal.Caption = ""
    
    Call Grid_Limpa(objGridDet)
    
    FormaPagto.ListIndex = -1
    Historico.ListIndex = -1
    Historico.Text = ""
    Origem.ListIndex = -1
    Filial.Clear
    
    NumDocDestino.Caption = ""
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    ValorBrutoVouAtual.Caption = ""
    ValorCMAVouAtual.Caption = ""
    FatorCMCC.Caption = ""
    FatorCMC.Caption = ""
    FatorCMR.Caption = ""
    FatorOVER.Caption = ""
    
    lNumvouAnt = 0
    sSerieAnt = ""
    sTipoAnt = ""
    iOrigemAnt = 0
    lCodigoAnt = 0
    dValorAnt = 0
    
    FrameResumo.Enabled = True

    iAlterado = 0

    Limpa_Tela_TRVOcorrencias = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRVOcorrencias:

    Limpa_Tela_TRVOcorrencias = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190189)

    End Select

    Exit Function

End Function

Function Traz_TRVOcorrencias_Tela(ByVal objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim objTRVOcrDet As ClassTRVOcorrenciaDet
Dim iLinha As Integer
Dim objVoucher As New ClassTRVVouchers

On Error GoTo Erro_Traz_TRVOcorrencias_Tela

    Call Limpa_Tela_TRVOcorrencias

    If objTRVOcorrencias.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objTRVOcorrencias.lCodigo)
        Codigo.PromptInclude = True
    End If

    'Lê o TRVOcorrencias que está sendo Passado
    lErro = CF("TRVOcorrencias_Le", objTRVOcorrencias)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190190

    If lErro = SUCESSO Then

        If objTRVOcorrencias.dtDataEmissao <> DATA_NULA Then
            DataEmissao.PromptInclude = False
            DataEmissao.Text = Format(objTRVOcorrencias.dtDataEmissao, "dd/mm/yy")
            DataEmissao.PromptInclude = True
        End If
       
        Observacao.Text = objTRVOcorrencias.sObservacao
        
        Select Case objTRVOcorrencias.iStatus
        
            Case STATUS_TRV_OCR_BLOQUEADO
                Status.Caption = STATUS_TRV_OCR_BLOQUEADO_TEXTO
            
            Case STATUS_TRV_OCR_CANCELADO
                Status.Caption = STATUS_TRV_OCR_CANCELADO_TEXTO
        
            Case STATUS_TRV_OCR_FATURADO
                Status.Caption = STATUS_TRV_OCR_FATURADO_TEXTO
        
            Case STATUS_TRV_OCR_LIBERADO
                Status.Caption = STATUS_TRV_OCR_LIBERADO_TEXTO
        
        End Select
        
        Call Combo_Seleciona_ItemData(FormaPagto, objTRVOcorrencias.iFormaPagto)

        NumeroVou.PromptInclude = False
        NumeroVou.Text = CStr(objTRVOcorrencias.lNumVou)
        NumeroVou.PromptInclude = True
        
        SerieVou.Text = objTRVOcorrencias.sSerie
        TipoVou.Text = objTRVOcorrencias.sTipoDoc
        
        Call TrazerVou_Click(False)
        
'        objVoucher.lNumVou = objTRVOcorrencias.lNumVou
'        objVoucher.sSerie = objTRVOcorrencias.sSerie
'        objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
'        objVoucher.sTipVou = objTRVOcorrencias.sTipoDoc
'
'        lErro = CF("TRVVouchers_Le", objVoucher)
'        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190191
'
'        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 190192
'
'        If objVoucher.lCliente <> 0 Then Cliente.Text = CStr(objVoucher.lCliente)
'        Call Cliente_Validate(bSGECancelDummy)
'
'        ClienteVou.Caption = Cliente.Text
'
        If objTRVOcorrencias.lCliente <> 0 Then Cliente.Text = CStr(objTRVOcorrencias.lCliente)
        Call Cliente_Validate(bSGECancelDummy)
'
'        ValorVou.Caption = Format(objVoucher.dvalor, "STANDARD")
'        DataEmissaoVou.Caption = Format(objVoucher.dtdata, "dd/mm/yyyy")
        
        If objTRVOcorrencias.iOrigem <> 0 Then Call Combo_Seleciona_ItemData(Origem, objTRVOcorrencias.iOrigem)

        Historico.Text = objTRVOcorrencias.sHistorico
        
        NumDocDestino.Caption = objTRVOcorrencias.lNumDocDestino
        
        iLinha = 0
        For Each objTRVOcrDet In objTRVOcorrencias.colDetalhamento
            
            iLinha = iLinha + 1
            
            Call Combo_Seleciona_ItemData(Tipo, objTRVOcrDet.iTipo)
            
            GridDet.TextMatrix(iLinha, iGrid_Tipo_Col) = Tipo.Text
            GridDet.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objTRVOcrDet.dValor, "STANDARD")
            
        Next
        
        objGridDet.iLinhasExistentes = objTRVOcorrencias.colDetalhamento.Count
        
        Call ValorTotal_Calcula

    End If

    iAlterado = 0

    Traz_TRVOcorrencias_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVOcorrencias_Tela:

    Traz_TRVOcorrencias_Tela = gErr

    Select Case gErr

        Case 190190, 190191
        
        Case 190192
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVoucher.lNumVou, objVoucher.sSerie, objVoucher.sTipVou)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190193)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 190194

    'Limpa Tela
    Call Limpa_Tela_TRVOcorrencias

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 190194

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190195)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190196)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 190197

    Call Limpa_Tela_TRVOcorrencias

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 190197

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190198)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTRVOcorrencias As New ClassTRVOcorrencias
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 190199

    objTRVOcorrencias.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TRVOCORRENCIAS", objTRVOcorrencias.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("TRVOcorrencias_Exclui", objTRVOcorrencias)
        If lErro <> SUCESSO Then gError 190200

        'Limpa Tela
        Call Limpa_Tela_TRVOcorrencias

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 190199
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 190200

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190201)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 190202

    End If
    
    If lCodigoAnt <> StrParaLong(Codigo.Text) Then
        Call OCR_Ajusta_Resumo
        lCodigoAnt = StrParaLong(Codigo.Text)
    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 190202

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190203)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 190204
                   
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 190205

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 190204, 190205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190206)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objcliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 190206

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 190207

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 190208

        If lErro = 17660 Then

            'Lê o Cliente
            objcliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objcliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 190209

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objcliente.lCodigo

                gError 190210
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        Else
            
            objcliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0
            
        End If
        
    'Não encontrou a STRING
    ElseIf lErro = 6731 Then
        
        'trecho incluido por Leo em 17/04/02
        objcliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 190211
        
        If lErro = SUCESSO Then gError 190212
        
    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 190206, 190208

        Case 190207
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 190209, 190211 'tratado na rotina chamada

        Case 190210
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 190212
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190213)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmissao_DownClick

    DataEmissao.SetFocus

    If Len(DataEmissao.ClipText) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 190214

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case gErr

        Case 190214

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190215)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmissao_UpClick

    DataEmissao.SetFocus

    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 190216

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case gErr

        Case 190216

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190217)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    If Len(Trim(DataEmissao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError 190218

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 190218

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190219)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Verifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Observacao
       '#######################################

    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190220)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Origem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Origem_Validate

    'Verifica se Origem está preenchida
    If Len(Trim(Origem.Text)) <> 0 Then

       '#######################################
       'CRITICA Origem
       '#######################################

    End If

    Exit Sub

Erro_Origem_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190221)

    End Select

    Exit Sub

End Sub

Private Sub Origem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Historico_Validate

    'Verifica se Historico está preenchida
    If Len(Trim(Historico.Text)) <> 0 Then

       If Len(Historico.Text) > STRING_TRV_OCR_HISTORICO Then gError 190222

    End If

    Exit Sub

Erro_Historico_Validate:

    Cancel = True

    Select Case gErr
    
        Case 190222
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICO", gErr, STRING_TRV_OCR_HISTORICO, Len(Historico.Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190223)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTRVOcorrencias As ClassTRVOcorrencias

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTRVOcorrencias = obj1

    'Mostra os dados do TRVOcorrencias na tela
    lErro = Traz_TRVOcorrencias_Tela(objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 190224

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 190224

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190225)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTRVOcorrencias As New ClassTRVOcorrencias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTRVOcorrencias.lCodigo = Codigo.Text

    End If

    Call Chama_Tela("TRVOcorrenciaLista", colSelecao, objTRVOcorrencias, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Det(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Det

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Valor")

    'campos de edição do grid
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Valor.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Tipo_Col = 1
    iGrid_Valor_Col = 2

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDet

    'Largura da primeira coluna
    GridDet.ColWidth(0) = 250

    'Linhas do grid
    objGridInt.objGrid.Rows = 20 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Det = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Det:

    Inicializa_Grid_Det = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190227)

    End Select

    Exit Function

End Function

Private Sub GridDet_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDet, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDet, iAlterado)
    End If

End Sub

Private Sub GridDet_GotFocus()
    Call Grid_Recebe_Foco(objGridDet)
End Sub

Private Sub GridDet_EnterCell()
    Call Grid_Entrada_Celula(objGridDet, iAlterado)
End Sub

Private Sub GridDet_LeaveCell()
    Call Saida_Celula(objGridDet)
End Sub

Private Sub GridDet_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDet, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDet, iAlterado)
    End If

End Sub

Private Sub GridDet_RowColChange()
    Call Grid_RowColChange(objGridDet)
End Sub

Private Sub GridDet_Scroll()
    Call Grid_Scroll(objGridDet)
End Sub

Private Sub GridDet_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridDet_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridDet.iLinhasExistentes
    iItemAtual = GridDet.Row
    
    Call Grid_Trata_Tecla1(KeyCode, objGridDet)

    If objGridDet.iLinhasExistentes < iLinhasExistentesAnt Then

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 190228

    End If

    Exit Sub

Erro_GridDet_KeyDown:

    Select Case gErr

        Case 190228

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190229)

    End Select

    Exit Sub

End Sub

Private Sub GridDet_LostFocus()
    Call Grid_Libera_Foco(objGridDet)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridDet.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Tipo_Col

                    lErro = Saida_Celula_Tipo(objGridInt)
                    If lErro <> SUCESSO Then gError 190230

                Case iGrid_Valor_Col
                
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError 190231

            End Select
                         
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 190232

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 190230, 190231
            'erros tratatos nas rotinas chamadas
        
        Case 190232
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190233)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
    
        Case Tipo.Name
            objControl.Enabled = True
    
        Case Valor.Name
        
            If Len(Trim(GridDet.TextMatrix(iLinha, iGrid_Tipo_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
     
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190234)

    End Select

    Exit Sub

End Sub

Function ValorTotal_Calcula() As Long

Dim iLinha As Integer
Dim dValorTotal As Double

On Error GoTo Erro_ValorTotal_Calcula

    For iLinha = 1 To objGridDet.iLinhasExistentes
        dValorTotal = dValorTotal + StrParaDbl(GridDet.TextMatrix(iLinha, iGrid_Valor_Col))
    Next

    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    If Abs(dValorAnt - dValorTotal) > DELTA_VALORMONETARIO Then
        Call OCR_Ajusta_Resumo
        dValorAnt = dValorTotal
    End If

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190235)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tipo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo que está deixando de ser a corrente

Dim lErro As Long
Dim sTipo As String

On Error GoTo Erro_Saida_Celula_Tipo

    Set objGridInt.objControle = Tipo

    If Len(Trim(Tipo.Text)) > 0 Then
        If GridDet.Row - GridDet.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190236

    Saida_Celula_Tipo = SUCESSO

    Exit Function

Erro_Saida_Celula_Tipo:

    Saida_Celula_Tipo = gErr

    Select Case gErr

        Case 190236
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190237)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    If Len(Trim(Valor.Text)) > 0 Then

        lErro = Valor_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 190238

        Valor.Text = Format(Valor.Text, "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190239

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 190240

   Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 190238 To 190240
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190241)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_TRVOCORRENCIAS", "TRVOcorrencias", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 190242
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 190242

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190243)
    
    End Select

    Exit Sub
    
End Sub

Public Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Tipo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDet)
End Sub

Public Sub Tipo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDet)
End Sub

Public Sub Tipo_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDet.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGridDet)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDet)
End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDet)
End Sub

Public Sub Valor_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDet.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridDet)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then Call LabelCliente_Click
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
    
    End If
    
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub
Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub
Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub
Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub BotaoAbrirDoc_Click()

Dim lErro As Long
Dim objObjeto As Object
Dim sTela As String
Dim bExisteDestino As Boolean
Dim lNumTitulo As Long
Dim sDoc As String
Dim objTRVOcorrencias As New ClassTRVOcorrencias

On Error GoTo Erro_BotaoAbrirDoc_Click

    If Len(Trim(NumDocDestino.Caption)) <> 0 Then

        objTRVOcorrencias.lCodigo = StrParaLong(Codigo.Text)
    
        lErro = CF("TRVOcorrencias_Le", objTRVOcorrencias)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192870
    
        lErro = CF("Verifica_Existencia_Doc_TRV", objTRVOcorrencias.lNumIntDocDestino, objTRVOcorrencias.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
        If lErro <> SUCESSO Then gError 192871
        
        Select Case objTRVOcorrencias.iTipoDocDestino
        
            Case TRV_TIPO_DOC_DESTINO_CREDFORN
                sTela = TRV_TIPO_DOC_DESTINO_CREDFORN_TELA
                Set objObjeto = New ClassCreditoPagar
                
            Case TRV_TIPO_DOC_DESTINO_DEBCLI
                sTela = TRV_TIPO_DOC_DESTINO_DEBCLI_TELA
                Set objObjeto = New ClassDebitoRecCli
        
            Case TRV_TIPO_DOC_DESTINO_TITPAG
                sTela = TRV_TIPO_DOC_DESTINO_TITPAG_TELA
                Set objObjeto = New ClassTituloPagar
        
            Case TRV_TIPO_DOC_DESTINO_TITREC
                sTela = TRV_TIPO_DOC_DESTINO_TITREC_TELA
                Set objObjeto = New ClassTituloReceber
                
            Case TRV_TIPO_DOC_DESTINO_NFSPAG
                sTela = TRV_TIPO_DOC_DESTINO_NFSPAG_TELA
                Set objObjeto = New ClassNFsPag
        
        End Select
        
        If Not (objObjeto Is Nothing) Then
        
            objObjeto.lNumIntDoc = objTRVOcorrencias.lNumIntDocDestino
            
            Call Chama_Tela(sTela, objObjeto)
            
        End If
        
    End If
    
    Exit Sub

Erro_BotaoAbrirDoc_Click:

    Select Case gErr
    
        Case 192870, 192871
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192872)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHistOcor_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistOcor_Click

    colSelecao.Add StrParaLong(NumeroVou.Text)
    colSelecao.Add TipoVou.Text
    colSelecao.Add SerieVou.Text

    Call Chama_Tela("OcorrenciasHistLista", colSelecao, Nothing, Nothing)

    Exit Sub

Erro_BotaoHistOcor_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Function OCR_Ajusta_Resumo() As Long
'Ajusta a parte do resumo da alteração

Dim lErro As Long
Dim objOcr As New ClassTRVOcorrencias
Dim objTipo As New ClassTRVTiposOcorrencia
Dim objInfo As ClassTRVVoucherInfoN
Dim objVou As New ClassTRVVouchers
Dim dValorB As Double
Dim dValorA As Double
Dim dValorO As Double
Dim dFator As Double

On Error GoTo Erro_OCR_Ajusta_Resumo

    dValorO = StrParaDbl(ValorTotal.Caption)

    objOcr.lCodigo = StrParaLong(Codigo.Text)

    objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVou.sSerie = SerieVou.Text
    objVou.sTipVou = TipoVou.Text
    objVou.lNumVou = StrParaLong(NumeroVou.Text)

    lErro = CF("TRVVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200672
    
    If lErro = SUCESSO Then
    
        lErro = CF("TRVVoucherInfo_Le", objVou)
        If lErro <> SUCESSO Then gError 200673
                        
        For Each objInfo In objVou.colTRVVoucherInfo
            If objInfo.sTipoDoc = TRV_TIPODOC_BRUTO_TEXTO Then
                dValorB = dValorB + objInfo.dValor
            End If
            If objInfo.sTipoDoc = TRV_TIPODOC_CMA_TEXTO Then
                dValorA = dValorA + objInfo.dValor
            End If
        Next
        
        ValorBrutoVouAtual.Caption = Format(dValorB, "STANDARD")
        ValorCMAVouAtual.Caption = Format(dValorA, "STANDARD")
        
    End If

    lErro = CF("TRVOcorrencias_Le", objOcr)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200674
    
    If lErro = SUCESSO Then
    
        FrameResumo.Enabled = False
    
        ValorBrutoVouOCR.Text = Format(objOcr.dValorOCRBrutoVou, "STANDARD")
        ValorCMAVouOCR.Text = Format(objOcr.dValorOCRCMAVou, "STANDARD")
    
        FatorCMCC.Caption = ""
        FatorCMC.Caption = ""
        FatorCMR.Caption = ""
        FatorOVER.Caption = ""
    
    Else
        
        FrameResumo.Enabled = True
   
        objTipo.iCodigo = Codigo_Extrai(Origem.Text)
        
        If objTipo.iCodigo <> 0 Then
        
            'Le o tipo da ocorrência para dar o tratamento adequado
            lErro = CF("TRVTiposOcorrencia_Le", objTipo)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200675
        
            If objTipo.iConsideraComisInt = MARCADO Then
        
                If objTipo.iIncideSobre = TRV_TIPO_OCR_INCIDE_BRUTO Then
                    ValorBrutoVouOCR.Text = Format(dValorO / (1 - IIf(objVou.iCartao = DESMARCADO, objVou.dComissaoAg, 0)), "STANDARD")
                    ValorCMAVouOCR.Text = Format(StrParaDbl(ValorBrutoVouOCR.Text) - dValorO, "STANDARD")
                ElseIf objTipo.iIncideSobre = TRV_TIPO_OCR_INCIDE_FAT Then
                    ValorBrutoVouOCR.Text = Format(-StrParaDbl(ValorCMAVouAtual.Caption) + dValorO, "STANDARD")
                    ValorCMAVouOCR.Text = Format(-StrParaDbl(ValorCMAVouAtual.Caption), "STANDARD")
                Else
                    ValorBrutoVouOCR.Text = Format(0, "STANDARD")
                    ValorCMAVouOCR.Text = Format(-dValorO, "STANDARD")
                End If
                
               
            End If
            
            dFator = 0
            If objTipo.iIncideSobre <> TRV_TIPO_OCR_INCIDE_CMA And objVou.dValorBruto > 0 Then
                dFator = StrParaDbl(ValorBrutoVouOCR.Text) / objVou.dValorBruto
            End If
        
            If objTipo.iAlteraCMCC = MARCADO Then
                FatorCMCC.Caption = Format(dFator, "PERCENT")
            Else
                FatorCMCC.Caption = ""
            End If
            If objTipo.iAlteraCMC = MARCADO Then
                FatorCMC.Caption = Format(dFator, "PERCENT")
            Else
                FatorCMC.Caption = ""
            End If
            If objTipo.iAlteraCMR = MARCADO Then
                FatorCMR.Caption = Format(dFator, "PERCENT")
            Else
                FatorCMR.Caption = ""
            End If
            If objTipo.iAlteraOVER = MARCADO Then
                FatorOVER.Caption = Format(dFator, "PERCENT")
            Else
                FatorOVER.Caption = ""
            End If
        
        End If
        
    End If

    ValorBrutoVouNovo.Text = Format(StrParaDbl(ValorBrutoVouAtual.Caption) + StrParaDbl(ValorBrutoVouOCR.Text), "STANDARD")
    ValorCMAVouNovo.Text = Format(StrParaDbl(ValorCMAVouAtual.Caption) + StrParaDbl(ValorCMAVouOCR.Text), "STANDARD")

    OCR_Ajusta_Resumo = SUCESSO

    Exit Function

Erro_OCR_Ajusta_Resumo:

    OCR_Ajusta_Resumo = gErr

    Select Case gErr
    
        Case 200672 To 200675

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200676)

    End Select

    Exit Function

End Function

Private Sub ValorBrutoVouNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorBrutoVouNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorBrutoVouNovo_Validate

    'Veifica se ValorBrutoVouNovo está preenchida
    If Len(Trim(ValorBrutoVouNovo.Text)) <> 0 Then

       'Critica a ValorBrutoVouNovo
       lErro = Valor_NaoNegativo_Critica(ValorBrutoVouNovo.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    
    ValorBrutoVouOCR.Text = Format(-StrParaDbl(ValorBrutoVouAtual.Caption) + StrParaDbl(ValorBrutoVouNovo.Text), "STANDARD")

    Exit Sub

Erro_ValorBrutoVouNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub ValorCMAVouNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorCMAVouNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorCMAVouNovo_Validate

    'Veifica se ValorCMAVouNovo está preenchida
    If Len(Trim(ValorCMAVouNovo.Text)) <> 0 Then

       'Critica a ValorCMAVouNovo
       lErro = Valor_NaoNegativo_Critica(ValorCMAVouNovo.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    
    ValorCMAVouOCR.Text = Format(-StrParaDbl(ValorCMAVouAtual.Caption) + StrParaDbl(ValorCMAVouNovo.Text), "STANDARD")

    Exit Sub

Erro_ValorCMAVouNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub ValorBrutoVouOCR_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorBrutoVouOCR_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorBrutoVouOCR_Validate

    'Veifica se ValorBrutoVouOCR está preenchida
    If Len(Trim(ValorBrutoVouOCR.Text)) <> 0 Then

       'Critica a ValorBrutoVouOCR
       lErro = Valor_Double_Critica(ValorBrutoVouOCR.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    
    ValorBrutoVouNovo.Text = Format(StrParaDbl(ValorBrutoVouAtual.Caption) + StrParaDbl(ValorBrutoVouOCR.Text), "STANDARD")

    Exit Sub

Erro_ValorBrutoVouOCR_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub ValorCMAVouOCR_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorCMAVouOCR_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorCMAVouOCR_Validate

    'Veifica se ValorCMAVouOCR está preenchida
    If Len(Trim(ValorCMAVouOCR.Text)) <> 0 Then

       'Critica a ValorCMAVouOCR
       lErro = Valor_Double_Critica(ValorCMAVouOCR.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    
    ValorCMAVouNovo.Text = Format(StrParaDbl(ValorCMAVouAtual.Caption) + StrParaDbl(ValorCMAVouOCR.Text), "STANDARD")
    
    Exit Sub

Erro_ValorCMAVouOCR_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

