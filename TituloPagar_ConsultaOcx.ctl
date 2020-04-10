VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl TituloPagar_ConsultaOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9405
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4485
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1455
      Width           =   9105
      Begin VB.CommandButton BotaoAnexos 
         Height          =   390
         Left            =   4185
         Picture         =   "TituloPagar_ConsultaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   206
         ToolTipText     =   "Anexar Arquivos"
         Top             =   1170
         Width           =   420
      End
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
         Left            =   3495
         TabIndex        =   200
         Top             =   1215
         Width           =   495
      End
      Begin VB.ComboBox Etapa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5355
         Style           =   2  'Dropdown List
         TabIndex        =   196
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton BotaoConsultaNFPag 
         Caption         =   "NFs Associadas a Fatura"
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
         Left            =   210
         TabIndex        =   190
         Top             =   4005
         Width           =   2835
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
         Height          =   705
         Left            =   7110
         Picture         =   "TituloPagar_ConsultaOcx.ctx":0196
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   1665
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   2430
         Left            =   210
         TabIndex        =   46
         Top             =   1515
         Width           =   8730
         Begin VB.Frame Frame1 
            Caption         =   "Retenções"
            Height          =   1245
            Left            =   3870
            TabIndex        =   163
            Top             =   1125
            Width           =   4725
            Begin VB.Label ISSRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1020
               TabIndex        =   192
               Top             =   915
               Width           =   1215
            End
            Begin VB.Label Label53 
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
               Height          =   195
               Left            =   600
               TabIndex        =   191
               Top             =   960
               Width           =   375
            End
            Begin VB.Label Label46 
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
               Left            =   2595
               TabIndex        =   171
               Top             =   600
               Width           =   555
            End
            Begin VB.Label CSLLRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3195
               TabIndex        =   170
               Top             =   555
               Width           =   1215
            End
            Begin VB.Label Label44 
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
               Left            =   210
               TabIndex        =   169
               Top             =   645
               Width           =   750
            End
            Begin VB.Label COFINSRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1035
               TabIndex        =   168
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label42 
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
               Height          =   195
               Left            =   2760
               TabIndex        =   167
               Top             =   270
               Width           =   375
            End
            Begin VB.Label PISRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3180
               TabIndex        =   166
               Top             =   225
               Width           =   1215
            End
            Begin VB.Label Label16 
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
               Left            =   720
               TabIndex        =   165
               Top             =   300
               Width           =   270
            End
            Begin VB.Label ValorIRRF 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1035
               TabIndex        =   164
               Top             =   255
               Width           =   1215
            End
         End
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
            Height          =   285
            Left            =   2250
            TabIndex        =   9
            Top             =   1935
            Width           =   930
         End
         Begin VB.Frame Frame3 
            Height          =   570
            Index           =   1
            Left            =   165
            TabIndex        =   50
            Top             =   1125
            Width           =   3120
            Begin VB.CheckBox CreditoIPI 
               Caption         =   "Crédito"
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
               Left            =   2100
               TabIndex        =   8
               Top             =   255
               Width           =   930
            End
            Begin VB.Label ValorIPI 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   705
               TabIndex        =   59
               Top             =   195
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "IPI:"
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
               Left            =   315
               TabIndex        =   60
               Top             =   240
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Height          =   600
            Left            =   135
            TabIndex        =   47
            Top             =   165
            Width           =   6195
            Begin VB.CheckBox CreditoICMS 
               Caption         =   "Crédito"
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
               Height          =   300
               Left            =   5175
               TabIndex        =   7
               Top             =   210
               Width           =   930
            End
            Begin VB.Label ValorICMSSubst 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3750
               TabIndex        =   61
               Top             =   210
               Width           =   1215
            End
            Begin VB.Label ValorICMS 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   750
               TabIndex        =   62
               Top             =   210
               Width           =   1215
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "ICMS:"
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
               TabIndex        =   63
               Top             =   270
               Width           =   525
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "ICMS Subst:"
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
               Left            =   2640
               TabIndex        =   64
               Top             =   270
               Width           =   1065
            End
         End
         Begin VB.Label ValorINSS 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   840
            TabIndex        =   65
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label OutrasDespesas 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7350
            TabIndex        =   66
            Top             =   810
            Width           =   1215
         End
         Begin VB.Label ValorSeguro 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3870
            TabIndex        =   67
            Top             =   810
            Width           =   1215
         End
         Begin VB.Label ValorFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   870
            TabIndex        =   68
            Top             =   810
            Width           =   1215
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7350
            TabIndex        =   69
            Top             =   375
            Width           =   1215
         End
         Begin VB.Label Label19 
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
            Height          =   255
            Left            =   3150
            TabIndex        =   70
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label18 
            Caption         =   "Outras Despesas:"
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
            Left            =   5760
            TabIndex        =   71
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label Label20 
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
            Height          =   255
            Left            =   300
            TabIndex        =   72
            Top             =   840
            Width           =   525
         End
         Begin VB.Label Label22 
            Caption         =   "Produtos:"
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
            Left            =   6465
            TabIndex        =   73
            Top             =   405
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "INSS:"
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
            Left            =   255
            TabIndex        =   74
            Top             =   1965
            Width           =   495
         End
      End
      Begin MSMask.MaskEdBox Projeto 
         Height          =   285
         Left            =   1050
         TabIndex        =   197
         Top             =   1215
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label55 
         Caption         =   "Diferença:"
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
         Left            =   2370
         TabIndex        =   203
         Top             =   150
         Width           =   900
      End
      Begin VB.Label Diferenca 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3330
         TabIndex        =   202
         Top             =   90
         Width           =   1215
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
         TabIndex        =   199
         Top             =   1260
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
         Left            =   4710
         TabIndex        =   198
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label HistoricoTit 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   173
         Top             =   840
         Width           =   7755
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   172
         Top             =   885
         Width           =   825
      End
      Begin VB.Label Saldo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5595
         TabIndex        =   75
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label NumPC 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   76
         Top             =   465
         Width           =   1215
      End
      Begin VB.Label Label30 
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
         Left            =   4980
         TabIndex        =   77
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5595
         TabIndex        =   78
         Top             =   465
         Width           =   1215
      End
      Begin VB.Label Label14 
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
         Left            =   4935
         TabIndex        =   79
         Top             =   525
         Width           =   615
      End
      Begin VB.Label LblNumPC 
         AutoSize        =   -1  'True
         Caption         =   "Nº PC:"
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
         Left            =   450
         TabIndex        =   80
         Top             =   510
         Width           =   585
      End
      Begin VB.Label LblFilialPC 
         AutoSize        =   -1  'True
         Caption         =   "Filial PC:"
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
         Left            =   2520
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   81
         Top             =   525
         Width           =   765
      End
      Begin VB.Label Label6 
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
         Left            =   525
         TabIndex        =   82
         Top             =   150
         Width           =   510
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   83
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label ComboFilialPC 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3330
         TabIndex        =   84
         Top             =   465
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4410
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Frame Frame6 
         Caption         =   "Parcelas"
         Height          =   3795
         Left            =   90
         TabIndex        =   55
         Top             =   90
         Width           =   9000
         Begin VB.ComboBox MotivoDiferenca 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   204
            Top             =   1260
            Width           =   2445
         End
         Begin MSMask.MaskEdBox ValorOriginal 
            Height          =   225
            Left            =   0
            TabIndex        =   205
            Top             =   825
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox CodigodeBarras 
            Height          =   315
            Left            =   195
            TabIndex        =   201
            Top             =   930
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   57
            Mask            =   "#####.#####.#####.######.#####.######.#.#################"
            PromptChar      =   " "
         End
         Begin VB.ComboBox ComboCobrador 
            Height          =   315
            Left            =   4545
            TabIndex        =   58
            Top             =   2145
            Width           =   2295
         End
         Begin VB.ComboBox ComboPortador 
            Height          =   315
            Left            =   1620
            TabIndex        =   57
            Top             =   2175
            Width           =   2445
         End
         Begin VB.TextBox StatusParcela 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   7470
            TabIndex        =   17
            Top             =   450
            Width           =   645
         End
         Begin MSMask.MaskEdBox SaldoParcela 
            Height          =   225
            Left            =   2370
            TabIndex        =   13
            Top             =   480
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   225
            Left            =   1260
            TabIndex        =   12
            Top             =   510
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   3570
            TabIndex        =   14
            Top             =   480
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   210
            TabIndex        =   11
            Top             =   510
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
         Begin VB.CheckBox Suspenso 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   225
            Left            =   6630
            TabIndex        =   16
            Top             =   510
            Width           =   780
         End
         Begin VB.ComboBox TipoCobranca 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4770
            TabIndex        =   15
            Top             =   450
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1755
            Left            =   90
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3096
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4470
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   9075
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
         Height          =   405
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Abre contabilização da baixa"
         Top             =   4035
         Width           =   3210
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
         Height          =   405
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprime o Recibo"
         Top             =   4035
         Width           =   2055
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Dados do Pagamento"
         Height          =   2160
         Index           =   0
         Left            =   2610
         TabIndex        =   44
         Top             =   1815
         Width           =   6375
         Begin VB.Frame Frame10 
            Caption         =   "Meio Pagamento"
            Height          =   585
            Left            =   2625
            TabIndex        =   45
            Top             =   195
            Width           =   3648
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Dinheiro"
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
               Index           =   2
               Left            =   2532
               TabIndex        =   48
               Top             =   285
               Width           =   1035
            End
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Cheque"
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
               Index           =   0
               Left            =   120
               TabIndex        =   49
               Top             =   255
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Borderô"
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
               Height          =   204
               Index           =   1
               Left            =   1272
               TabIndex        =   51
               Top             =   288
               Width           =   996
            End
         End
         Begin VB.Label Label54 
            Caption         =   "Nominal à:"
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
            Left            =   1950
            TabIndex        =   194
            Top             =   1770
            Width           =   930
         End
         Begin VB.Label Beneficiario 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2955
            TabIndex        =   193
            Top             =   1710
            Width           =   3315
         End
         Begin VB.Label Label29 
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
            Left            =   345
            TabIndex        =   109
            Top             =   435
            Width           =   555
         End
         Begin VB.Label Label28 
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
            Left            =   105
            TabIndex        =   110
            Top             =   1380
            Width           =   810
         End
         Begin VB.Label Label27 
            Caption         =   "Local de Pagto:"
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
            Left            =   3150
            TabIndex        =   111
            Top             =   900
            Width           =   1365
         End
         Begin VB.Label Label26 
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
            Left            =   390
            TabIndex        =   112
            Top             =   915
            Width           =   495
         End
         Begin VB.Label Label32 
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
            Height          =   240
            Left            =   180
            TabIndex        =   113
            Top             =   1770
            Width           =   750
         End
         Begin VB.Label ContaCorrente 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   990
            TabIndex        =   114
            Top             =   412
            Width           =   1530
         End
         Begin VB.Label ValorPagoPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   990
            TabIndex        =   115
            Top             =   885
            Width           =   1530
         End
         Begin VB.Label Historico 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   990
            TabIndex        =   116
            Top             =   1350
            Width           =   5250
         End
         Begin VB.Label NumOuSequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   990
            TabIndex        =   117
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label Portador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4575
            TabIndex        =   118
            Top             =   870
            Width           =   1695
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Cheque Pré de Terceiros"
         Height          =   1770
         Index           =   3
         Left            =   2610
         TabIndex        =   175
         Top             =   1830
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Label Label52 
            Caption         =   "Banco:"
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
            Left            =   165
            TabIndex        =   189
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Label51 
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
            Height          =   255
            Left            =   1800
            TabIndex        =   188
            Top             =   345
            Width           =   765
         End
         Begin VB.Label Label50 
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
            Left            =   3810
            TabIndex        =   187
            Top             =   345
            Width           =   600
         End
         Begin VB.Label Label49 
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
            Left            =   60
            TabIndex        =   186
            Top             =   855
            Width           =   705
         End
         Begin VB.Label Label47 
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
            Left            =   3840
            TabIndex        =   185
            Top             =   855
            Width           =   510
         End
         Begin VB.Label Label45 
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
            Left            =   4305
            TabIndex        =   184
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label BancoCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   810
            TabIndex        =   183
            Top             =   315
            Width           =   825
         End
         Begin VB.Label ContaCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4425
            TabIndex        =   182
            Top             =   315
            Width           =   1740
         End
         Begin VB.Label ValorCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4440
            TabIndex        =   181
            Top             =   795
            Width           =   1740
         End
         Begin VB.Label FilialEmpresaCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5640
            TabIndex        =   180
            Top             =   1290
            Width           =   525
         End
         Begin VB.Label AgenciaCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2580
            TabIndex        =   179
            Top             =   315
            Width           =   1125
         End
         Begin VB.Label NumeroCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   825
            TabIndex        =   178
            Top             =   795
            Width           =   1455
         End
         Begin VB.Label ClienteCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   795
            TabIndex        =   177
            Top             =   1290
            Width           =   3255
         End
         Begin VB.Label Label43 
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
            Height          =   255
            Left            =   105
            TabIndex        =   176
            Top             =   1320
            Width           =   690
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Crédito"
         Height          =   1770
         Index           =   2
         Left            =   2610
         TabIndex        =   42
         Top             =   1830
         Visible         =   0   'False
         Width           =   6375
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
            TabIndex        =   85
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
            Left            =   4095
            TabIndex        =   86
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
            TabIndex        =   87
            Top             =   840
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
            Left            =   4035
            TabIndex        =   88
            Top             =   840
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
            TabIndex        =   89
            Top             =   1275
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
            Left            =   3300
            TabIndex        =   90
            Top             =   1275
            Width           =   1245
         End
         Begin VB.Label DataEmissaoCred 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1950
            TabIndex        =   91
            Top             =   382
            Width           =   1095
         End
         Begin VB.Label SiglaDocumentoCR 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4680
            TabIndex        =   92
            Top             =   382
            Width           =   1080
         End
         Begin VB.Label ValorCredito 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4680
            TabIndex        =   93
            Top             =   810
            Width           =   1080
         End
         Begin VB.Label FilialEmpresaCR 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4680
            TabIndex        =   94
            Top             =   1245
            Width           =   525
         End
         Begin VB.Label NumTitulo 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1950
            TabIndex        =   95
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label SaldoCredito 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1950
            TabIndex        =   96
            Top             =   1245
            Width           =   1095
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Adiantamento à Fornecedor"
         Height          =   1770
         Index           =   1
         Left            =   2610
         TabIndex        =   43
         Top             =   1830
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Label NumeroMP 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4635
            TabIndex        =   97
            Top             =   825
            Width           =   720
         End
         Begin VB.Label CCIntNomeReduzido 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4635
            TabIndex        =   98
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label FilialEmpresaPA 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4635
            TabIndex        =   99
            Top             =   1275
            Width           =   525
         End
         Begin VB.Label ValorPA 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1740
            TabIndex        =   100
            Top             =   1275
            Width           =   1095
         End
         Begin VB.Label MeioPagtoDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1740
            TabIndex        =   101
            Top             =   825
            Width           =   1095
         End
         Begin VB.Label DataMovimento 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1740
            TabIndex        =   102
            Top             =   390
            Width           =   1095
         End
         Begin VB.Label Label33 
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
            Left            =   3300
            TabIndex        =   103
            Top             =   1305
            Width           =   1245
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
            Left            =   1125
            TabIndex        =   104
            Top             =   1305
            Width           =   510
         End
         Begin VB.Label Label24 
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
            Left            =   3840
            TabIndex        =   105
            Top             =   855
            Width           =   705
         End
         Begin VB.Label Label23 
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
            Left            =   600
            TabIndex        =   106
            Top             =   855
            Width           =   1035
         End
         Begin VB.Label Label21 
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
            Left            =   3195
            TabIndex        =   107
            Top             =   420
            Width           =   1350
         End
         Begin VB.Label Label15 
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
            Left            =   390
            TabIndex        =   108
            Top             =   420
            Width           =   1245
         End
      End
      Begin VB.ComboBox Sequencial 
         Height          =   315
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   210
         Width           =   825
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados da Baixa"
         Height          =   1170
         Left            =   105
         TabIndex        =   52
         Top             =   600
         Width           =   8940
         Begin VB.Label DataBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1470
            TabIndex        =   119
            Top             =   322
            Width           =   1125
         End
         Begin VB.Label Label17 
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
            Left            =   900
            TabIndex        =   120
            Top             =   375
            Width           =   480
         End
         Begin VB.Label ValorPago 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4260
            TabIndex        =   121
            Top             =   322
            Width           =   945
         End
         Begin VB.Label Juros 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6930
            TabIndex        =   122
            Top             =   742
            Width           =   945
         End
         Begin VB.Label Multa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4260
            TabIndex        =   123
            Top             =   742
            Width           =   945
         End
         Begin VB.Label Desconto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1470
            TabIndex        =   124
            Top             =   742
            Width           =   1125
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
            Left            =   3195
            TabIndex        =   125
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label Label5 
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
            Left            =   6360
            TabIndex        =   126
            Top             =   795
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
            Left            =   3660
            TabIndex        =   127
            Top             =   795
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
            Left            =   480
            TabIndex        =   128
            Top             =   795
            Width           =   885
         End
         Begin VB.Label ValorBaixado 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6930
            TabIndex        =   129
            Top             =   322
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
            Left            =   5610
            TabIndex        =   130
            Top             =   375
            Width           =   1245
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Baixa"
         Height          =   2160
         Left            =   105
         TabIndex        =   53
         Top             =   1815
         Width           =   2385
         Begin VB.OptionButton Pagamento 
            Caption         =   "Cheque de Terceiros"
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
            Index           =   3
            Left            =   120
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   1800
            Value           =   -1  'True
            Width           =   2115
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Crédito/Devolução"
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
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1395
            Width           =   1950
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Adiantamento à Forn."
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
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   930
            Width           =   2145
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Pagamento"
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   465
            Width           =   1290
         End
      End
      Begin MSComCtl2.UpDown UpDownParcela 
         Height          =   300
         Index           =   1
         Left            =   1650
         TabIndex        =   41
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         OrigLeft        =   1650
         OrigTop         =   217
         OrigRight       =   1890
         OrigBottom      =   517
         Enabled         =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6480
         TabIndex        =   131
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label DataVencParcBaixa 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7650
         TabIndex        =   132
         Top             =   217
         Width           =   1245
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
         Left            =   4590
         TabIndex        =   133
         Top             =   270
         Width           =   510
      End
      Begin VB.Label ValorParcBaixa 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5190
         TabIndex        =   134
         Top             =   217
         Width           =   945
      End
      Begin VB.Label Parcela 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   1260
         TabIndex        =   135
         Top             =   217
         Width           =   375
      End
      Begin VB.Label Label13 
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
         Left            =   2370
         TabIndex        =   136
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label7 
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
         Left            =   510
         TabIndex        =   137
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   4
      Left            =   90
      TabIndex        =   26
      Top             =   1470
      Visible         =   0   'False
      Width           =   9165
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4920
         TabIndex        =   195
         Tag             =   "1"
         Top             =   1920
         Width           =   870
      End
      Begin VB.CheckBox CTBAglutina 
         Enabled         =   0   'False
         Height          =   210
         Left            =   8040
         TabIndex        =   34
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4500
         MaxLength       =   150
         TabIndex        =   33
         Top             =   1260
         Width           =   3495
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   930
         Left            =   45
         TabIndex        =   54
         Top             =   3090
         Width           =   5895
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
            TabIndex        =   138
            Top             =   600
            Visible         =   0   'False
            Width           =   1440
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
            TabIndex        =   139
            Top             =   270
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   140
            Top             =   255
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   141
            Top             =   585
            Visible         =   0   'False
            Width           =   3720
         End
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
         TabIndex        =   27
         Top             =   255
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4140
         TabIndex        =   32
         Top             =   1260
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
         Bindings        =   "TituloPagar_ConsultaOcx.ctx":30AC
         Height          =   225
         Left            =   60
         TabIndex        =   28
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
         Left            =   2970
         TabIndex        =   31
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   1800
         TabIndex        =   30
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1050
         TabIndex        =   29
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   945
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
         TabIndex        =   142
         Top             =   135
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   143
         Top             =   90
         Width           =   1530
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
         TabIndex        =   144
         Top             =   465
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5970
         TabIndex        =   145
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3705
         TabIndex        =   146
         Top             =   450
         Width           =   1185
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
         TabIndex        =   147
         Top             =   465
         Width           =   870
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
         TabIndex        =   148
         Top             =   765
         Width           =   1140
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
         TabIndex        =   149
         Top             =   2805
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   150
         Top             =   2790
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   151
         Top             =   2790
         Width           =   1155
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
         TabIndex        =   152
         Top             =   495
         Width           =   480
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
         TabIndex        =   153
         Top             =   135
         Width           =   1035
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
         TabIndex        =   154
         Top             =   135
         Width           =   450
      End
      Begin VB.Label CTBDataContabil 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   750
         TabIndex        =   155
         Top             =   442
         Width           =   1095
      End
      Begin VB.Label CTBDocumento 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3705
         TabIndex        =   156
         Top             =   75
         Width           =   705
      End
      Begin VB.Label CTBLote 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5970
         TabIndex        =   157
         Top             =   75
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6660
      ScaleHeight     =   795
      ScaleWidth      =   2565
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   45
      Width           =   2625
      Begin VB.CommandButton BotaoFechar 
         Height          =   675
         Left            =   2025
         Picture         =   "TituloPagar_ConsultaOcx.ctx":30B7
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   480
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   675
         Left            =   1455
         Picture         =   "TituloPagar_ConsultaOcx.ctx":3235
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   480
      End
      Begin VB.CommandButton BotaoConsulta 
         Height          =   675
         Left            =   90
         Picture         =   "TituloPagar_ConsultaOcx.ctx":3767
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   60
         Width           =   1275
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4710
      TabIndex        =   1
      Top             =   30
      Width           =   1815
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   1245
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   375
      Width           =   2190
   End
   Begin MSMask.MaskEdBox NumeroTitulo 
      Height          =   300
      Left            =   4710
      TabIndex        =   3
      Top             =   390
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "999999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   300
      Left            =   1245
      TabIndex        =   0
      Top             =   30
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4860
      Left            =   60
      TabIndex        =   40
      Top             =   1110
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8573
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
   Begin MSComCtl2.UpDown UpDownEmissao 
      Height          =   300
      Left            =   2295
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   735
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataEmissao 
      Height          =   300
      Left            =   1245
      TabIndex        =   4
      Top             =   735
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
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
      Height          =   195
      Left            =   420
      TabIndex        =   158
      Top             =   780
      Width           =   765
   End
   Begin VB.Label FornecedorLabel 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
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
      Left            =   150
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   159
      Top             =   90
      Width           =   1035
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
      Left            =   3945
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   160
      Top             =   420
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   " Filial:"
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
      Left            =   4140
      TabIndex        =   161
      Top             =   90
      Width           =   525
   End
   Begin VB.Label LabelTipo 
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
      Left            =   735
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   162
      Top             =   420
      Width           =   450
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9420
      Y1              =   1065
      Y2              =   1065
   End
End
Attribute VB_Name = "TituloPagar_ConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTTituloPagar_Consulta
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoAnexos_Click()
    Call objCT.BotaoAnexos_Click
End Sub

Private Sub BotaoConsulta_Click()
     Call objCT.BotaoConsulta_Click
End Sub

Private Sub BotaoDocOriginal_Click()
     Call objCT.BotaoDocOriginal_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoImprimirRecibo_Click()
    Call objCT.BotaoImprimirRecibo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub LabelTipo_Click()
     Call objCT.LabelTipo_Click
End Sub

Private Sub NumeroTitulo_GotFocus()
     Call objCT.NumeroTitulo_GotFocus
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub FornecedorLabel_Click()
     Call objCT.FornecedorLabel_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long
     Trata_Parametros = objCT.Trata_Parametros(objTituloPagar)
End Function

Private Sub UpDownParcela_DownClick(Index As Integer)
     Call objCT.UpDownParcela_DownClick(Index)
End Sub

Private Sub UpDownParcela_UpClick(Index As Integer)
     Call objCT.UpDownParcela_UpClick(Index)
End Sub

Private Sub Sequencial_Click()
     Call objCT.Sequencial_Click
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub NumeroTitulo_Validate(Cancel As Boolean)
     Call objCT.NumeroTitulo_Validate(Cancel)
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub Pagamento_Click(Index As Integer)
     Call objCT.Pagamento_Click(Index)
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
    Set objCT = New CTTituloPagar_Consulta
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


Private Sub ValorIPI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorIPI, Source, X, Y)
End Sub

Private Sub ValorIPI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorIPI, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub ValorICMSSubst_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorICMSSubst, Source, X, Y)
End Sub

Private Sub ValorICMSSubst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorICMSSubst, Button, Shift, X, Y)
End Sub

Private Sub ValorICMS_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorICMS, Source, X, Y)
End Sub

Private Sub ValorICMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorICMS, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub ValorINSS_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorINSS, Source, X, Y)
End Sub

Private Sub ValorINSS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorINSS, Button, Shift, X, Y)
End Sub

Private Sub ValorIRRF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorIRRF, Source, X, Y)
End Sub

Private Sub ValorIRRF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorIRRF, Button, Shift, X, Y)
End Sub

Private Sub OutrasDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OutrasDespesas, Source, X, Y)
End Sub

Private Sub OutrasDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OutrasDespesas, Button, Shift, X, Y)
End Sub

Private Sub ValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorSeguro, Source, X, Y)
End Sub

Private Sub ValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorSeguro, Button, Shift, X, Y)
End Sub

Private Sub ValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorFrete, Source, X, Y)
End Sub

Private Sub ValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorFrete, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Saldo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Saldo, Source, X, Y)
End Sub

Private Sub Saldo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Saldo, Button, Shift, X, Y)
End Sub

Private Sub NumPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumPC, Source, X, Y)
End Sub

Private Sub NumPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumPC, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LblNumPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNumPC, Source, X, Y)
End Sub

Private Sub LblNumPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNumPC, Button, Shift, X, Y)
End Sub

Private Sub LblFilialPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilialPC, Source, X, Y)
End Sub

Private Sub LblFilialPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilialPC, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub ComboFilialPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ComboFilialPC, Source, X, Y)
End Sub

Private Sub ComboFilialPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ComboFilialPC, Button, Shift, X, Y)
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

Private Sub ValorCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorCredito, Source, X, Y)
End Sub

Private Sub ValorCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorCredito, Button, Shift, X, Y)
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

Private Sub SaldoCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoCredito, Source, X, Y)
End Sub

Private Sub SaldoCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoCredito, Button, Shift, X, Y)
End Sub

Private Sub NumeroMP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroMP, Source, X, Y)
End Sub

Private Sub NumeroMP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroMP, Button, Shift, X, Y)
End Sub

Private Sub CCIntNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CCIntNomeReduzido, Source, X, Y)
End Sub

Private Sub CCIntNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CCIntNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaPA, Source, X, Y)
End Sub

Private Sub FilialEmpresaPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaPA, Button, Shift, X, Y)
End Sub

Private Sub ValorPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPA, Source, X, Y)
End Sub

Private Sub ValorPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPA, Button, Shift, X, Y)
End Sub

Private Sub MeioPagtoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MeioPagtoDescricao, Source, X, Y)
End Sub

Private Sub MeioPagtoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MeioPagtoDescricao, Button, Shift, X, Y)
End Sub

Private Sub DataMovimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataMovimento, Source, X, Y)
End Sub

Private Sub DataMovimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataMovimento, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub ContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCorrente, Source, X, Y)
End Sub

Private Sub ContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCorrente, Button, Shift, X, Y)
End Sub

Private Sub ValorPagoPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPagoPagto, Source, X, Y)
End Sub

Private Sub ValorPagoPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPagoPagto, Button, Shift, X, Y)
End Sub

Private Sub Historico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Historico, Source, X, Y)
End Sub

Private Sub Historico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Historico, Button, Shift, X, Y)
End Sub

Private Sub NumOuSequencial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumOuSequencial, Source, X, Y)
End Sub

Private Sub NumOuSequencial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumOuSequencial, Button, Shift, X, Y)
End Sub

Private Sub Portador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Portador, Source, X, Y)
End Sub

Private Sub Portador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Portador, Button, Shift, X, Y)
End Sub

Private Sub DataBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataBaixa, Source, X, Y)
End Sub

Private Sub DataBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub DataVencParcBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataVencParcBaixa, Source, X, Y)
End Sub

Private Sub DataVencParcBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataVencParcBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub ValorParcBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorParcBaixa, Source, X, Y)
End Sub

Private Sub ValorParcBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorParcBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub CTBDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBDataContabil, Source, X, Y)
End Sub

Private Sub CTBDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBDataContabil, Button, Shift, X, Y)
End Sub

Private Sub CTBDocumento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBDocumento, Source, X, Y)
End Sub

Private Sub CTBDocumento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBDocumento, Button, Shift, X, Y)
End Sub

Private Sub CTBLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLote, Source, X, Y)
End Sub

Private Sub CTBLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLote, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

'#####################################
'Inserido por Wagner 21/08/2006
Private Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub
'#####################################

'#################################
'Inserido por Wagner 08/03/2006
Private Sub BotaoConsultaNFPag_Click()
     Call objCT.BotaoConsultaNFPag_Click
End Sub
'#################################


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
