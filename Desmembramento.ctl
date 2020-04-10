VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Desmembramento 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4785
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Width           =   9255
      Begin VB.CommandButton BotaoEstoque 
         Caption         =   "Estoque"
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
         Left            =   2445
         TabIndex        =   71
         Top             =   4350
         Width           =   1335
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   1245
         Width           =   660
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   1905
         MaxLength       =   50
         TabIndex        =   64
         Top             =   1275
         Width           =   2600
      End
      Begin MSMask.MaskEdBox ContaContabilProducao 
         Height          =   270
         Left            =   4920
         TabIndex        =   70
         Top             =   2040
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox ContaContabilEst 
         Height          =   270
         Left            =   2715
         TabIndex        =   69
         Top             =   2055
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   270
         Left            =   5490
         TabIndex        =   66
         Top             =   1275
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   476
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   270
         Left            =   180
         TabIndex        =   63
         Top             =   1290
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   270
         Left            =   2205
         TabIndex        =   67
         Top             =   1650
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   270
         Left            =   6360
         TabIndex        =   68
         Top             =   1650
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
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
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   3690
         Picture         =   "Desmembramento.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Numeração Automática"
         Top             =   150
         Width           =   300
      End
      Begin VB.CommandButton BotaoPlanoConta 
         Caption         =   "Plano de Contas"
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
         Left            =   7155
         TabIndex        =   9
         Top             =   4350
         Width           =   1815
      End
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
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
         Left            =   150
         TabIndex        =   7
         Top             =   4350
         Width           =   1380
      End
      Begin VB.CommandButton BotaoCcls 
         Caption         =   "Centros de Custo"
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
         Left            =   4665
         TabIndex        =   8
         Top             =   4350
         Width           =   1815
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   6030
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   150
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CclPadrao 
         Height          =   300
         Left            =   2925
         TabIndex        =   6
         Top             =   585
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   4965
         TabIndex        =   2
         Top             =   150
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   300
         Left            =   2910
         TabIndex        =   1
         Top             =   135
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
         Height          =   2295
         Left            =   90
         TabIndex        =   4
         Top             =   1140
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox Hora 
         Height          =   300
         Left            =   7425
         TabIndex        =   3
         Top             =   150
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   6885
         TabIndex        =   72
         Top             =   195
         Width           =   480
      End
      Begin VB.Label CodigoLabel 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2205
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   165
         Width           =   660
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4410
         TabIndex        =   38
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Material Produzido"
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
         TabIndex        =   37
         Top             =   915
         Width           =   1590
      End
      Begin VB.Label CclPadraoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo/Lucro Padrão:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   645
         Width           =   2670
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4725
      Index           =   2
      Left            =   210
      TabIndex        =   10
      Top             =   855
      Visible         =   0   'False
      Width           =   9165
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4080
         TabIndex        =   74
         Tag             =   "1"
         Top             =   1395
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
         Height          =   270
         Left            =   7815
         TabIndex        =   15
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   900
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
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
         Left            =   6360
         TabIndex        =   14
         Top             =   30
         Width           =   1245
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
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
         Left            =   6360
         TabIndex        =   16
         Top             =   345
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4920
         TabIndex        =   24
         Top             =   1755
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   26
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   25
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6360
         TabIndex        =   28
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   40
         Top             =   3630
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
            TabIndex        =   44
            Top             =   660
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
            TabIndex        =   43
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   42
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   41
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
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
         Left            =   3435
         TabIndex        =   19
         Top             =   1050
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   20
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   23
         Top             =   1890
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   22
         Top             =   1830
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   21
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   585
         TabIndex        =   13
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5595
         TabIndex        =   12
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3795
         TabIndex        =   11
         Top             =   120
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   27
         Top             =   1320
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   3180
         Left            =   6375
         TabIndex        =   29
         Top             =   1575
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5609
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   3180
         Left            =   6360
         TabIndex        =   30
         Top             =   1560
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5609
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   6420
         TabIndex        =   17
         Top             =   660
         Width           =   690
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
         TabIndex        =   61
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   60
         Top             =   120
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
         Left            =   4230
         TabIndex        =   59
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   58
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   57
         Top             =   555
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
         Left            =   1995
         TabIndex        =   56
         Top             =   585
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
         TabIndex        =   55
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   54
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   6345
         TabIndex        =   53
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   52
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
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
         TabIndex        =   51
         Top             =   3180
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   50
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   49
         Top             =   3165
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
         Left            =   45
         TabIndex        =   48
         Top             =   555
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   47
         Top             =   165
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   46
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7260
      ScaleHeight     =   495
      ScaleWidth      =   2100
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Width           =   2160
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Desmembramento.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Desmembramento.ctx":0274
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Desmembramento.ctx":03F2
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Desmembramento.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5190
      Left            =   105
      TabIndex        =   62
      Top             =   480
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   9155
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
End
Attribute VB_Name = "Desmembramento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'inicio contabilidade
Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Public gobjMovEst As ClassMovEstoque

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'mnemonicos
Private Const CODIGO1 As String = "Codigo"
Private Const DATA1 As String = "Data"
Private Const ESTORNO1 As String = "Estorno"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const CCL1 As String = "Ccl"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const CONTACONTABILEST1 As String = "ContaContabilEst"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"

'Declaração das Variáveis Globais
Public iAlterado As Integer
Dim iCodigoAlterado As Integer
Dim iFrameAtual As Integer
Dim lCodigoAntigo As Long

'GRID
Dim objGrid As AdmGrid
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_ContaContabilEst_Col As Integer
Dim iGrid_ContaContabilProducao_Col As Integer

'BROWSERS
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento '
Attribute objEventoEstoque.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Movimentos = 1
Private Const TAB_Contabilizacao = 2

Private Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridMovimentos.Row = 0 Then gError 132650

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 132651

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        
        objEstoqueProduto.sAlmoxarifadoNomeReduzido = Almoxarifado.Text

        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        gError 132652
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 132651
        
        Case 132650
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 132652
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179344)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 132653

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    '''03/09/01 - Marcelo inclusao da pergunta se deseja excluir a Entrada de Producao
    'Envia aviso perguntando se realmente deseja excluir a Entrada de Producao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_DESMEMBRAMENTO", objMovEstoque.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a producao
        lErro = CF("MovEstoque_Exclui", objMovEstoque, objContabil)
        If lErro <> SUCESSO Then gError 132654

        Call Limpa_Tela_ProducaoEntrada
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 132653
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 132654
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179345)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 132655

    'Verifica se o produto está preenchido e se a linha corrente é diferente da linha fixa
    If iProdutoPreenchido = PRODUTO_PREENCHIDO And GridMovimentos.Row <> 0 Then

        'Preenche o Nome do Almoxarifado
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        'Preenche a conta contabil de estoque depois que o produto e o Almoxarifado já estão preenchidos
        lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
        If lErro <> SUCESSO Then gError 132656

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 132655, 132656

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179346)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("MovEstoque_Automatico", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then gError 132657

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    lCodigoAntigo = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 132657
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179347)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovimentos.Row = 0 Then gError 132658
    
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = "" Then gError 132659

    sConta = String(STRING_CONTA, 0)

    If GridMovimentos.Col = iGrid_ContaContabilEst_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEst.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 132660
    
    ElseIf GridMovimentos.Col = iGrid_ContaContabilProducao_Col Then

        lErro = CF("Conta_Formata", ContaContabilProducao.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 132661

    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 132658
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 132659
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 132660, 132661

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179348)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If GridMovimentos.Col <> iGrid_ContaContabilEst_Col And GridMovimentos.Col <> iGrid_ContaContabilProducao_Col Then
        Me.Show
        Exit Sub
    End If
        
    If objPlanoConta.sConta <> "" Then
   
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 132662
               
        If GridMovimentos.Col = iGrid_ContaContabilEst_Col Then
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
        
            GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = ContaContabilEst.Text
        Else
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaEnxuta
            ContaContabilProducao.PromptInclude = True
        
            GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = ContaContabilProducao.Text
        End If
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 132662
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179349)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    'se o codigo foi trocado
    If lCodigoAntigo <> StrParaLong(Trim(Codigo.Text)) Then
    
        If Len(Trim(Codigo.ClipText)) > 0 Then
        
            objMovEstoque.lCodigo = Codigo.Text
            
            'Le o Movimento de Estoque e Verifica se ele já foi estornado
            lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_PRODUCAO)
            If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 132663
            
            'Se todos os Itens do Movimento foram estornados
            If lErro = 78885 Then gError 132664
            
            If lErro = SUCESSO Then
            
                If objMovEstoque.iTipoMov <> MOV_EST_PRODUCAO Then gError 132665
                
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")
                
                If vbMsg = vbNo Then gError 132666
                
                lErro = Preenche_Tela(objMovEstoque)
                If lErro <> SUCESSO Then gError 132667
                      
            End If
        
        End If
      
        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))
      
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
            
        Case 132665
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_PENTRADA", gErr, objMovEstoque.lCodigo)
            lCodigoAntigo = 0
            
        Case 132663, 132667
        
        Case 132666
            
        Case 132664
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179350)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ContaContabilEst_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabilEst_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ContaContabilEst_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ContaContabilEst_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabilEst
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ContaContabilProducao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabilProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ContaContabilProducao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ContaContabilProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabilProducao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCclPadrao As String

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    'Inicializa todos os objeventos
    Set gobjMovEst = New ClassMovEstoque
    
    Set objEventoCodigo = New AdmEvento
    Set objEventoCclPadrao = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    
    'Mostra a Data Atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 132668
    
    'Inicializa mascara de contaContabilEst
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEst)
    If lErro <> SUCESSO Then gError 132669
    
    'Inicializa a mascara de ContaContabilProducao
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilProducao)
    If lErro <> SUCESSO Then gError 132670

    'Inicializa Máscara para CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 132671

    Ccl.Mask = sMascaraCclPadrao
    CclPadrao.Mask = sMascaraCclPadrao

    'Formata Quantidade
    Quantidade.Format = FORMATO_ESTOQUE
       
    'Inicialização do GridMovimentos
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridMovimentos(objGrid)
    If lErro <> SUCESSO Then gError 132672
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE)
    If lErro <> SUCESSO Then gError 132673
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 132668, 132671, 132672, 132673, 132669, 132670

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179351)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_GridMovimentos(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil de Estoque")
    objGridInt.colColuna.Add ("Conta Contabil de Produção")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ContaContabilEst.Name)
    objGridInt.colCampo.Add (ContaContabilProducao.Name)
    
    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_Almoxarifado_Col = 5
    iGrid_Ccl_Col = 6
    iGrid_ContaContabilEst_Col = 7
    iGrid_ContaContabilProducao_Col = 8
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridMovimentos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMovimentos = SUCESSO

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no Banco de Dados
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objTipoMovEst As ClassTipoMovEst
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Extrai

    sTabela = "MovEstProd"
    
    'Lê os atributos de objMovEstoque que aparecem na Tela
    If Len(Trim(Codigo.ClipText)) <> 0 Then objMovEstoque.lCodigo = StrParaLong(Codigo.Text)

    If Len(Trim(Data.ClipText)) <> 0 Then
        objMovEstoque.dtData = StrParaDate(Data.Text)
    End If

    If Len(Trim(Hora.ClipText)) > 0 Then
        objMovEstoque.dtHora = StrParaDate(Hora.Text)
    Else
        objMovEstoque.dtHora = 0
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objMovEstoque.lCodigo, 0, "Codigo"
    colCampoValor.Add "Data", objMovEstoque.dtData, 0, "Data"
    colCampoValor.Add "Hora", CDbl(objMovEstoque.dtHora), 0, "Hora"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    'colSelecao.Add "TipoMov", OP_IGUAL, MOV_EST_PRODUCAO
    colSelecao.Add "NumIntDocEst", OP_IGUAL, 0

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179352)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do Banco de Dados
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objMovEstoque.lCodigo = colCampoValor.Item("Codigo").vValor
    objMovEstoque.dtData = colCampoValor.Item("Data").vValor
    objMovEstoque.dtHora = colCampoValor.Item("Hora").vValor
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 132674

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 132674

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179353)

    End Select

    Exit Sub

End Sub

Function Preenche_Tela(objMovEstoque As ClassMovEstoque) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Preenche_Tela

    'Limpa a tela sem Fechar o Comando de setas
    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)
    
    'Se o grid permite excluir e incluir Linhas
    If objGrid.iProibidoIncluir <> GRID_PROIBIDO_INCLUIR And objGrid.iProibidoExcluir <> GRID_PROIBIDO_EXCLUIR Then
        'prepara o Grid para não permitir inserir e excluir Linhas
        objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
        objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
        Call Grid_Inicializa(objGrid)
    End If
    
    Set gobjMovEst = objMovEstoque
    Set objMovEstoque.colItens = New ColItensMovEstoque

    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovEstoqueItens_Le1", objMovEstoque, MOV_EST_DESMEMBRAMENTO_SAIDA)
    If lErro <> SUCESSO And lErro <> 55387 Then gError 132675

    If lErro = 55387 Then gError 132676

    'Coloca os Dados na Tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objMovEstoque.lCodigo)
    Codigo.PromptInclude = True

    Call DateParaMasked(Data, objMovEstoque.dtData)

    Hora.PromptInclude = False
    'este teste está correto
    If objMovEstoque.dtData <> DATA_NULA Then Hora.Text = Format(objMovEstoque.dtHora, "hh:mm:ss")
    Hora.PromptInclude = True

    lErro = Preenche_GridMovimentos(objMovEstoque.colItens)
    If lErro <> SUCESSO Then gError 132677
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovEstoque.colItens(1).lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 132678

    iAlterado = 0
    lCodigoAntigo = objMovEstoque.lCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 132675, 132677, 132678

        Case 132676
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_PRODUCAO", gErr, objMovEstoque.lCodigo)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179354)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMovimentos(colItens As ColItensMovEstoque) As Long

Dim iIndice As Integer
Dim sProdutoMascarado As String, sCclMascarado As String
Dim lErro As Long
Dim objTipoMovEst As ClassTipoMovEst
Dim objItemMovEstoque As ClassItemMovEstoque
Dim sContaEnxutaEst As String
Dim sContaEnxutaProducao As String
Dim objFilialEmpresa As New AdmFiliais
Dim objProduto As New ClassProduto
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_Preenche_GridMovimentos

    'Preenche GridMovimentos
    For Each objItemMovEstoque In colItens

        iIndice = iIndice + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 132679

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilEst <> "" Then
        
            sContaEnxutaEst = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilEst, sContaEnxutaEst)
            If lErro <> SUCESSO Then gError 132680
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxutaEst
            ContaContabilEst.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
            
        End If
        
         'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then
        
            sContaEnxutaProducao = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaProducao)
            If lErro <> SUCESSO Then gError 132681
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaEnxutaProducao
            ContaContabilProducao.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducao_Col) = ContaContabilProducao.Text
            
        End If
        
        GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridMovimentos.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemMovEstoque.sProdutoDesc
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemMovEstoque.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemMovEstoque.dQuantidade)
        GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemMovEstoque.sAlmoxarifadoNomeRed

        If objItemMovEstoque.sCCL <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemMovEstoque.sCCL, sCclMascarado)
            If lErro <> SUCESSO Then gError 132682

        Else

            sCclMascarado = ""

        End If

        GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado
                      
        objProduto.sCodigo = objItemMovEstoque.sProduto
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 132683

        objItemMovEstoque.sSiglaUMEst = objProduto.sSiglaUMEstoque
        
    Next

    objGrid.iLinhasExistentes = colItens.Count

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 132684

    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 132682, 132684, 132680, 132681, 132683

        Case 132679
            Call Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objItemMovEstoque.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179355)

    End Select

    Exit Function

End Function

Private Sub CodigoLabel_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection

    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    colSelecao.Add MOV_EST_DESMEMBRAMENTO_SAIDA
    colSelecao.Add MOV_EST_DESMEMBRAMENTO_SAIDA
    colSelecao.Add MOV_EST_DESMEMBRAMENTO_SAIDA
    
    Call Chama_Tela("MovEstoqueLista1", colSelecao, objMovEstoque, objEventoCodigo)
   
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objMovEstoque As ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_objCodigoEvento_evSelecao

    Set objMovEstoque = obj1

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 132685

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show
    
    Exit Sub

Erro_objCodigoEvento_evSelecao:

    Select Case gErr

        Case 132685

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179356)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual
        
            Case TAB_Movimentos
                Parent.HelpContextID = IDH_ENTRADA_MATERIAL_PRODUZIDO_MOVIMENTOS
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ENTRADA_MATERIAL_PRODUZIDO_CONTABILIZACAO
                        
        End Select

    End If

End Sub

Private Sub CclPadraoLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCclPadrao_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCCL, sCclMascarado)
    If lErro <> SUCESSO Then gError 132686

    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 132686
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCCL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179357)

    End Select

    Exit Sub

End Sub
Private Sub BotaoCcls_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) <> 0 And GridMovimentos.Row > 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCCL, sCclMascarado)
        If lErro <> SUCESSO Then gError 132687

        'Coloca o valor do Ccl na coluna correspondente
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 132687

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179358)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoProdutos_Click

    If GridMovimentos.Row = 0 Then gError 132688
    
    'Lista de produtos produzidos e inventariados
    Call Chama_Tela("ProdutoProduz_EstoqLista", colSelecao, objProduto, objEventoProduto)
    
   Exit Sub
   
Erro_BotaoProdutos_Click:

    Select Case gErr
        
        Case 132688
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179359)
        
    End Select
    
    Exit Sub
    
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objCTBConfig As New ClassCTBConfig
Dim objItemMovEst As New ClassItemMovEstoque
    
On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 132689

    If GridMovimentos.Row = 0 Then gError 132690
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 132691

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 132692

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 132693

    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 132694

    If lErro = 28030 Then gError 132695
    
    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True

    If Not (Me.ActiveControl Is Produto) Then

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = sProdutoMascarado
        
        'se a ContaContabilProdução não estiver preenchida em Produto procurar em TipoProduto
        If Trim(objProduto.sContaContabilProducao) = "" Then
            
            objTipoDeProduto.iTipo = objProduto.iTipo
            
            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then gError 132696
            
            If lErro = 22531 Then gError 132697
            
            objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
            
            'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
            If Trim(objProduto.sContaContabilProducao) = "" Then
                                
                objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
                objCTBConfig.iFilialEmpresa = giFilialEmpresa
                        
                lErro = CF("CTBConfig_Le", objCTBConfig)
                If lErro <> SUCESSO And lErro <> 9755 Then gError 132698
                
                If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
                
            End If
        
        End If
    
        'Preenche a Linha do Grid
        lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
        If lErro <> SUCESSO Then gError 132699
        
        'Preenche a conta contabil de estoque depois que o produto e o Almoxarifado já estão preenchidos
        lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
        If lErro <> SUCESSO Then gError 132700

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 132690
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 132691, 132694, 132699, 132700, 132696, 132698
        
        Case 132692
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, GridMovimentos.Row)
        
        Case 132693
            Call Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objProduto.sCodigo)
    
        Case 132695
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
                
        Case 132697
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)

        Case 132689
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, Produto.Text)
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179360)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objMovEstoque Is Nothing Then

        'Lê MovEstoque no Banco de Dados
'hora

        objMovEstoque.iFilialEmpresa = giFilialEmpresa

        lErro = CF("MovEstoque_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30128 Then gError 132701

        'Se o movimento existe
        If lErro = SUCESSO Then

            If objMovEstoque.iTipoMov <> MOV_EST_DESMEMBRAMENTO_SAIDA Or objMovEstoque.iTipoMov <> MOV_EST_DESMEMBRAMENTO_ENTRADA Then gError 132702

            lErro = Preenche_Tela(objMovEstoque)
            If lErro <> SUCESSO Then gError 132703

        Else
            'Se ele não existe exibe apenas o código
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objMovEstoque.lCodigo)
            Codigo.PromptInclude = True

            lCodigoAntigo = objMovEstoque.lCodigo

        End If

    Else

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 132701, 132703

        Case 132702
            Call Rotina_Erro(vbOKOnly, "ERRO_MOV_EST_NAO_DESMEMBRAMENTO", gErr, objMovEstoque.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179361)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub CclPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclPadrao_Validate

    'Verifica se o CclPadrao foi Preenchida
    If Len(Trim(CclPadrao.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", CclPadrao.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 132704

        'se o ccl não estiver cadastrado
        If lErro = 5703 Then gError 132705

    End If

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 132704

        Case 132705
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179362)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 132706

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 132706

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179363)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 132707

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 132707

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179364)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 132708

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 132708

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179365)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer, lNumIntDoc As Long
Dim sCodProduto2 As String
Dim sProdutoFormatado2 As String
Dim iProdutoPreenchido2 As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then

        'Verifica se produto está preenchido
        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
    
        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 132709
    
        If gobjMovEst.colItens.Count >= GridMovimentos.Row And GridMovimentos.Row > 0 Then
            lNumIntDoc = gobjMovEst.colItens(GridMovimentos.Row).lNumIntDoc
        Else
            lNumIntDoc = 0
        End If
        
        lErro = CF("Produto_Formata", sCodProduto2, sProdutoFormatado2, iProdutoPreenchido2)
        If lErro <> SUCESSO Then gError 132710
    
        If objControl.Name = "Produto" Then
    
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
    
            End If
    
        ElseIf objControl.Name = "UnidadeMed" Then
    
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Or lNumIntDoc <> 0 Or Left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) = "#" Then
                
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
    
                objProduto.sCodigo = sProdutoFormatado
    
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 132711
    
                If lErro = 28030 Then gError 132712
    
                objClasseUM.iClasse = objProduto.iClasseUM
    
                'Preenche a List da Combo UnidadeMed com as UM's do Produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 132713
    
                'Guardo o valor da Unidade de Medida da Linha
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
    
                'Limpar as Unidades utilizadas anteriormente
                UnidadeMed.Clear
    
                For Each objUnidadeDeMedida In colSiglas
                    UnidadeMed.AddItem objUnidadeDeMedida.sSigla
    
                Next
    
                'Tento selecionar na Combo a Unidade anterior
                If UnidadeMed.ListCount <> 0 Then
    
                    For iIndice = 0 To UnidadeMed.ListCount - 1
    
                        If UnidadeMed.List(iIndice) = sUnidadeMed Then
                            UnidadeMed.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
                
            End If
    
        ElseIf objControl.Name = "Quantidade" Or objControl.Name = "Almoxarifado" Then
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 And Left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        ElseIf objControl.Name = "Ccl" Or objControl.Name = "ContaContabilEst" Or objControl.Name = "ContaContabilProducao" Then
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
                                 
        End If

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 132709, 132711, 132712, 132713

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179366)

    End Select

    Exit Sub

End Sub

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer, lNumIntDoc As Long
Dim iLinhaAnterior As Integer
Dim iLinhasExistentes As Integer 'm

On Error GoTo Erro_GridMovimentos_KeyDown

    If gobjMovEst.colItens.Count >= GridMovimentos.Row Then
        lNumIntDoc = gobjMovEst.colItens(GridMovimentos.Row).lNumIntDoc
    Else
        lNumIntDoc = 0
    End If

    If lNumIntDoc = 0 Then

        'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then
        
            'Guarda iLinhasExistentes
            iLinhasExistentesAnterior = objGrid.iLinhasExistentes
    
            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridMovimentos.Row
    
        End If

        Call Grid_Trata_Tecla1(KeyCode, objGrid)

        'Verifica se a Linha foi realmente excluída
        If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then
            
            'Exclui de colItens o Item correspondente, se houver
            gobjMovEst.colItens.Remove iLinhaAnterior

            For iLinhasExistentes = 1 To objGrid.iLinhasExistentes 'm
                If gobjMovEst.colItens(iLinhasExistentes).iPossuiGrade = MARCADO Then
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
                Else
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
                End If
                
            Next

            GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes

        End If

    End If
    
    Exit Sub
    
Erro_GridMovimentos_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179367)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objCTBConfig As New ClassCTBConfig
Dim colItensRomaneioGrade As New Collection
Dim objItemMovEst As New ClassItemMovEstoque
Dim objRomaneioGrade As New ClassRomaneioGrade

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 132714

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'se é um produto gerencial e não é pai de grade ==> erro
            If lErro = 25043 Then gError 132715
            
            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 132716

            'se o conta de produção não estiver associada ao produto
            If Trim(objProduto.sContaContabilProducao) = "" Then
                    
                objTipoDeProduto.iTipo = objProduto.iTipo
                
                'pesquisa a conta junto ao tipo do produto
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError 132717
                
                If lErro = 22531 Then gError 132718
                
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
                                        
                'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
                If Trim(objProduto.sContaContabilProducao) = "" Then
                                    
                    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
                    objCTBConfig.iFilialEmpresa = giFilialEmpresa
                            
                    lErro = CF("CTBConfig_Le", objCTBConfig)
                    If lErro <> SUCESSO And lErro <> 9755 Then gError 132719
                    
                    If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
                    
                End If
                
            End If
                
            If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 132720
                
            If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 132721
                
            ' preenche a linha do produto
            lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
            If lErro <> SUCESSO Then gError 132722
            
            'Preenche a conta contabil de estoque depois que o produto e o Almoxarifado já estão preenchidos
            lErro = Preenche_ContaContabilEst(Produto.Text)
            If lErro <> SUCESSO Then gError 132723
        
        End If

    Else
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Descricao_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEst_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilProducao_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132724

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 132724, 132722, 132723, 132717, 132719, 132714
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 132716
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 132720
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 132721
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case 132718
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case 117645, 132715 'Alterado por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179368)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto, objItemMovEst As ClassItemMovEstoque) As Long

Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sContaEnxuta As String
Dim sAlmoxarifadoPadrao As String
Dim objItemOP As New ClassItemOP
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_ProdutoLinha_Preenche
         
    'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
    lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
    If lErro <> SUCESSO Then gError 52224
    
    'preenche o grid
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao

    If Trim(objProduto.sContaContabilProducao) <> "" Then
    
        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaEnxuta)
        If lErro <> SUCESSO Then gError 132725
    
        'preenche  a ContaContabilProducao
        ContaContabilEst.PromptInclude = False
        ContaContabilEst.Text = sContaEnxuta
        ContaContabilEst.PromptInclude = True
    
        'preenche Conta De Producao
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilProducao_Col) = ContaContabilEst.Text
    
    End If
    
    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Descricao_Col) = objProduto.sDescricao
    
    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError 132726

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = CclPadrao.Text


    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
        
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        
        gobjMovEst.colItens.Add1 objItemMovEst
    
        objItemMovEst.iPossuiGrade = DESMARCADO
                    
        objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
        objItemMovEst.sProduto = objProduto.sCodigo
    
    End If
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 132726, 132725

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179369)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim vbMsg As VbMsgBoxResult
Dim sPercentual As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 132727

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 132728

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132729

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 132727, 132729, 132728
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179370)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado está preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 132730

        'Valida o ALmoxarifado
        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 132731
        
        'Se não for encontrado --> Erro
        If lErro = 25157 Then gError 132732
        If lErro = 25162 Then gError 132733

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132734

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 132734, 132730, 132731
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 132732
            'Pergunta de deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)
            'Se a resposta for sim
            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                'Chama a Tela Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 132733

            'Pergunta se deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(Almoxarifado.Text))
            'Se a resposta for positiva
            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = Codigo_Extrai(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179371)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Trim(Ccl.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 132735

        If lErro = 5703 Then gError 132736

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132737

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 132735, 132737
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 132736
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsg = vbYes Then
            
                objCcl.sCCL = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179372)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 132738

    Call Limpa_Tela_ProducaoEntrada

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 132738

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179373)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsgRes As VbMsgBoxResult
Dim sCodigoOP As String
Dim bRegravando As Boolean

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 132739

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 132740

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridMovimentos
    If objGrid.iLinhasExistentes = 0 Then gError 132741

    'Para cada MovEstoque
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 132742
    
    Next

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30128 Then gError 132743
    
    bRegravando = False
    
    If lErro = SUCESSO Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS")
        If vbMsgRes = vbNo Then gError 132744
        
        lErro = CF("MovEstoqueItens_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30116 Then gError 123800
    
        If lErro = SUCESSO Then bRegravando = True
    
    End If
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 132745
           
    'Move os dados para o objMovimentoEstoque
    lErro = Move_Tela_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError 132746
    
    'Vai usar os itens que já estão no BD para ter o NumIntDoc e não
    'Desmembrar de forma diferente (Possível alteração do Kit)
    If Not bRegravando Then
        lErro = Move_Desmembramento_Memoria(objMovEstoque)
        If lErro <> SUCESSO Then gError 132749
    End If
    
    'Grava no BD(inclusive os dados contabeis)
    lErro = CF("MovEstoque_Grava_Generico", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 132747

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
       
        Case 132739
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 132740
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 132741
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 132742
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 132746, 132747, 132743, 132749
        
        Case 132744, 132745
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179374)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objMovEstoque As ClassMovEstoque) As Long
'Preenche objMovEstoque (inclusive colItens)

Dim iIndice As Integer
Dim lCodigo As Long
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    'Carrega o Código
    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)
    
    'Carrega a Data
    objMovEstoque.dtData = StrParaDate(Data.Text)
    
'hora
    If Len(Trim(Hora.ClipText)) > 0 Then
        objMovEstoque.dtHora = StrParaDate(Hora.Text)
    Else
        objMovEstoque.dtHora = Time
    End If
    
    'A Filial Empresa
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    'Varre o Grid de Itens de Movimentos
    For iIndice = 1 To objGrid.iLinhasExistentes
                
        'Pega todos os Itens mesmo que seja extorno
        lErro = Move_Itens_Memoria(iIndice, objMovEstoque)
        If lErro <> SUCESSO Then gError 132748

    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 132748

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179375)

    End Select

    Exit Function

End Function

Function Move_Itens_Memoria(iIndice As Integer, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sCclFormatada As String, sCCL As String
Dim iCclPreenchida As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaFormatadaEst As String
Dim iContaPreenchida As Integer
Dim sContaFormatadaProducao As String
Dim colRastreamento As New Collection
Dim iTipoMovEstoque As Integer
Dim objItemMovEst As ClassItemMovEstoque
Dim colApropriacaoInsumos As New Collection

On Error GoTo Erro_Move_Itens_Memoria

    With GridMovimentos
         
        iTipoMovEstoque = MOV_EST_DESMEMBRAMENTO_SAIDA
        
        objAlmoxarifado.sNomeReduzido = .TextMatrix(iIndice, iGrid_Almoxarifado_Col)

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 132750

        If lErro = 25060 Then gError 132751
             
        sCCL = .TextMatrix(iIndice, iGrid_Ccl_Col)
        
        If Len(Trim(sCCL)) <> 0 Then
        
            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", .TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 132752
        
        Else
        
            sCclFormatada = ""

        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilEst_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilEst_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 132753
        
        Else
            sContaFormatadaEst = ""
        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col), sContaFormatadaProducao, iContaPreenchida)
            If lErro <> SUCESSO Then gError 132754
        
        Else
            sContaFormatadaProducao = ""
        End If
        
        'Formata o Produto para BD
        sProdutoFormatado = ""
        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 132755
        
        Set colRastreamento = New Collection
                        
        Set objItemMovEst = objMovEstoque.colItens.Add(gobjMovEst.colItens(iIndice).lNumIntDoc, iTipoMovEstoque, 0, 0, sProdutoFormatado, .TextMatrix(iIndice, iGrid_Descricao_Col), .TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(.TextMatrix(iIndice, iGrid_Quantidade_Col)), objAlmoxarifado.iCodigo, .TextMatrix(iIndice, iGrid_Almoxarifado_Col), 0, sCclFormatada, 0, "", "", sContaFormatadaProducao, sContaFormatadaEst, 0, colRastreamento, colApropriacaoInsumos, DATA_NULA)
                
    End With

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case 132750, 132752, 132755, 132753, 132754
        
        Case 132751
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179376)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 132756

    Call Limpa_Tela_ProducaoEntrada

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 132756

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179377)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ProducaoEntrada()

Dim lErro As Long
Dim lCodigo As Long
On Error GoTo Erro_Limpa_Tela_ProducaoEntrada

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'prepara o Grid para permitir inserir e excluir Linhas
    objGrid.iProibidoIncluir = 0
    objGrid.iProibidoExcluir = 0
    Call Grid_Inicializa(objGrid)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'Remove os ítens de colItensNumIntDoc
'    Set colItensNumIntDoc = New Collection

    Set gobjMovEst = New ClassMovEstoque

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True

    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    lCodigoAntigo = 0

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_ProducaoEntrada:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179378)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set gobjMovEst = Nothing
    
    Set objEventoCodigo = Nothing
    Set objEventoCclPadrao = Nothing
    Set objEventoCcl = Nothing
    Set objEventoProduto = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoEstoque = Nothing
    
    'eventos associados a contabilidade
    Set objEventoDoc = Nothing

    Set objGrid = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
   
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Private Sub CclPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_Click()

    If UnidadeMed.ListIndex = -1 Then Exit Sub

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132757

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 132757
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 179379)

    End Select

    Exit Function

End Function

Private Sub GridMovimentos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Private Sub GridMovimentos_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridMovimentos_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridMovimentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If


End Sub

Private Sub GridMovimentos_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 132758
        
        If objGridInt.objGrid Is GridMovimentos Then
        
            Select Case GridMovimentos.Col
    
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 132759
    
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 132760
                
                Case iGrid_Almoxarifado_Col
    
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 132761
                
                Case iGrid_Ccl_Col
    
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 132762

                Case iGrid_UnidadeMed_Col
    
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 132763
                    
                Case iGrid_ContaContabilEst_Col
                    lErro = Saida_Celula_ContaContabilEst(objGridInt)
                    If lErro <> SUCESSO Then gError 132764
                    
                Case iGrid_ContaContabilProducao_Col
                    lErro = Saida_Celula_ContaContabilProducao(objGridInt)
                    If lErro <> SUCESSO Then gError 132765

                Case Else
                    lErro = Saida_Celula_Grid(objGridInt)
                    If lErro <> SUCESSO Then gError 132766
    
            End Select
        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 132767

        iAlterado = REGISTRO_ALTERADO

    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:
    
    Saida_Celula = gErr
    
    Select Case gErr

        Case 132759, 132760, 132762, 132763, 132766, 132764, 132765, 132761, 132758

        Case 132767
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179380)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilEst(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilEst

    Set objGrid.objControle = ContaContabilEst

    If Len(Trim(ContaContabilEst.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilEst.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 132768
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 132769
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaMascarada
            ContaContabilEst.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEst.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 132770
    
            'conta não cadastrada
            If lErro = 5700 Then gError 132771
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132772
    
    Saida_Celula_ContaContabilEst = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEst:

    Saida_Celula_ContaContabilEst = gErr

    Select Case gErr

        Case 132768, 132770, 132772
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 132769
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 132771
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEst.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179381)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilProducao(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilProducao

    Set objGrid.objControle = ContaContabilProducao

    If Len(Trim(ContaContabilProducao.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilProducao.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 132773
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 132774
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaMascarada
            ContaContabilProducao.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 132775
    
            'conta não cadastrada
            If lErro = 5700 Then gError 132776
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132777
    
    Saida_Celula_ContaContabilProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilProducao:

    Saida_Celula_ContaContabilProducao = gErr

    Select Case gErr

        Case 132773, 132775, 132777
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 132774
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 132776
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilProducao.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179382)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Grid(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Grid

    Select Case GridMovimentos.Col

        Case iGrid_Produto_Col

            Set objGridInt.objControle = Produto

        Case iGrid_UnidadeMed_Col

            Set objGridInt.objControle = UnidadeMed

        Case iGrid_Quantidade_Col

            Set objGridInt.objControle = Quantidade

        Case iGrid_Almoxarifado_Col

            Set objGridInt.objControle = Almoxarifado

        Case iGrid_Ccl_Col

            Set objGridInt.objControle = Ccl
            
        Case iGrid_ContaContabilEst_Col
        
            Set objGridInt.objControle = ContaContabilEst

    End Select

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132778

    Saida_Celula_Grid = SUCESSO

    Exit Function

Erro_Saida_Celula_Grid:

    Saida_Celula_Grid = gErr

    Select Case gErr

        Case 132778

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179383)

    End Select

    Exit Function

End Function

Private Sub GridMovimentos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub GridMovimentos_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
        
        Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'inicio contabilidade

Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

End Sub

Private Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Private Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub

Private Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Private Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Private Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Private Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Private Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Private Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Private Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Private Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Private Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Private Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Private Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Private Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Private Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Private Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Private Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Private Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Private Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Private Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Private Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Private Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Private Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Private Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Private Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Private Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Private Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Private Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Private Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Private Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim dQuantidadeConvertida As Double
Dim dQuantidade As Double
Dim sProduto As String
Dim sUM As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case QUANT_ESTOQUE
            For iLinha = 1 To objGrid.iLinhasExistentes
            
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    lErro = CF("UMEstoque_Conversao", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 132779

                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case CODIGO1
            If Len(Codigo.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Codigo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case DATA1
            If Len(Data.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
                    
        Case CCL1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case ALMOXARIFADO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
                            
        Case PRODUTO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case UNIDADE_MED
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case DESCRICAO_ITEM
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
                    
        Case QUANTIDADE1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)
                    sUM = GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 132780
                    
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case CONTACONTABILEST1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEst_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEst_Col)
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case Else
            gError 132781

        End Select

        Calcula_Mnemonico = SUCESSO

        Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 132781
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 132779
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179384)

    End Select

    Exit Function

End Function

Private Function Preenche_ContaContabilEst(sProduto As String) As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabilEst
        
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) > 0 And Len(Trim(sProduto)) > 0 Then
    
        'preenche o objEstoqueProduto
        objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
        
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 132782
        
        If lErro = 25060 Then gError 132783
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 132784
        
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 49991 Then gError 132785
        
        If lErro = SUCESSO Then
        
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 132786
        
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
            
        End If
        
    End If
    
    Preenche_ContaContabilEst = SUCESSO
    
    Exit Function
    
Erro_Preenche_ContaContabilEst:

    Preenche_ContaContabilEst = gErr
    
    Select Case gErr
        
        Case 132782, 132784, 132785
        
        Case 132786
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
             
        Case 132783
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179385)
    
    End Select
    
    Exit Function
        
End Function

Private Function Preenche_Almoxarifado(iFilialEmpresa As Integer, sOPCodigo As String, sProduto As String) As Long
'preenche o almoxarifado no grid a partir do item da OP

Dim objItemOP As New ClassItemOP
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long

On Error GoTo Erro_Preenche_Almoxarifado

    objItemOP.iFilialEmpresa = giFilialEmpresa
    objItemOP.sCodigo = sOPCodigo
    objItemOP.sProduto = sProduto

    lErro = CF("ItemOP_Le", objItemOP)
    If lErro <> SUCESSO And lErro <> 34711 Then gError 132787

    If lErro = 34711 Then gError 132788
    
    objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
    
    'le o nome reduzido do almoxarifado associado ao itemop
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 132789
    
    If lErro = 25056 Then gError 132790
    
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

    Preenche_Almoxarifado = SUCESSO

    Exit Function

Erro_Preenche_Almoxarifado:

    Preenche_Almoxarifado = gErr
    
    Select Case gErr
    
        Case 132788
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 132787, 132789

        Case 132790
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objItemOP.iAlmoxarifado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179386)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ENTRADA_MATERIAL_PRODUZIDO_MOVIMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Desmembramento de Material"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Desmembramento"
    
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

'**** fim do trecho a ser copiado *****

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is CclPadrao Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is ContaContabilEst Or Me.ActiveControl Is ContaContabilProducao Then
            Call BotaoPlanoConta_Click
        End If
    
    End If

End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
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

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
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

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Function Move_Desmembramento_Memoria(ByVal objMovEst As ClassMovEstoque) As Long

Dim lErro As Long
Dim objKit As ClassKit
Dim objProdutoKit As ClassProdutoKit
Dim dFatorMultiplicacao As Double
Dim dFatorConversao As Double
Dim objItemMovEstSaida As ClassItemMovEstoque
Dim objItemMovEstEntrada As ClassItemMovEstoque
Dim colItemMovEst As New Collection
Dim iAlmoxarifado As Integer
Dim objProduto As New ClassProduto
Dim colApropriacaoInsumos As New Collection
Dim colRastreamento As New Collection

On Error GoTo Erro_Move_Desmembramento_Memoria
    
    'Para Cada objItemMovDestoque a ser desmembrado (de saída do estoque)
    For Each objItemMovEstSaida In objMovEst.colItens
        
        'Obtém o Kit Padrão com os componentes
        Set objKit = New ClassKit
        
        objKit.sProdutoRaiz = objItemMovEstSaida.sProduto
        
        lErro = CF("Kit_Le_Padrao", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then gError 132791

        If lErro = 106304 Then gError 132792
        
        'leitura dos componentes do kit
        lErro = CF("Kit_Le_Componentes", objKit)
        If lErro <> SUCESSO And lErro <> 21831 Then gError 132793
    
        'Cada filho no Kit vira um objItemMovEstoque de entrada
        For Each objProdutoKit In objKit.colComponentes

            Set objProduto = New ClassProduto

            objProduto.sCodigo = objProdutoKit.sProduto

            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 132795
        
            If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
            
                'Converte para UM da tela
                lErro = CF("UM_Conversao", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objItemMovEstSaida.sSiglaUM, dFatorConversao)
                If lErro <> SUCESSO Then gError 132796
            
                'Acha o Fator de multimicação da quantidade da tela para o Kit
                dFatorMultiplicacao = objItemMovEstSaida.dQuantidade / (objProdutoKit.dQuantidade * dFatorConversao)
               
            Else

                Set objItemMovEstEntrada = New ClassItemMovEstoque
                                
                'Preenche o Item de entrada no estoque
                If objProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_FIXA Then
                    objItemMovEstEntrada.dQuantidade = objProdutoKit.dQuantidade
                Else
                    objItemMovEstEntrada.dQuantidade = objProdutoKit.dQuantidade * dFatorMultiplicacao
                End If
                
                objItemMovEstEntrada.sProduto = objProdutoKit.sProduto
                objItemMovEstEntrada.sSiglaUM = objProdutoKit.sUnidadeMed
                objItemMovEstEntrada.iTipoMov = MOV_EST_DESMEMBRAMENTO_ENTRADA
                objItemMovEstEntrada.lCodigo = objMovEst.lCodigo
                objItemMovEstEntrada.sCCL = objItemMovEstSaida.sCCL
                objItemMovEstEntrada.iPossuiGrade = DESMARCADO
                objItemMovEstEntrada.sContaContabilEst = objProduto.sContaContabilProducao
                objItemMovEstEntrada.sProdutoDesc = objProduto.sDescricao
               
                lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 23796 Then gError 132794
            
                If lErro = SUCESSO And iAlmoxarifado <> 0 Then
                    objItemMovEstEntrada.iAlmoxarifado = iAlmoxarifado
                End If
               
                colItemMovEst.Add objItemMovEstEntrada
    
            End If
                   
        Next
        
    Next
    
    'Adiciona os Itens de Entrada no Movimento
    For Each objItemMovEstEntrada In colItemMovEst
        With objItemMovEstEntrada
            objMovEst.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, "", 0, .sCCL, 0, "", "", "", .sContaContabilEst, 0, colRastreamento, colApropriacaoInsumos, DATA_NULA
        End With
    Next
    
    Move_Desmembramento_Memoria = SUCESSO
        
    Exit Function
    
Erro_Move_Desmembramento_Memoria:
        
    Move_Desmembramento_Memoria = gErr
        
    Select Case gErr
    
        Case 132791, 132793, 132795, 132796, 132794
        
        Case 132792
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz, objKit.sVersao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179387)
        
    End Select
    
    Exit Function
        
End Function
