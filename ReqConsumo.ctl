VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ReqConsumo 
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5235
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   705
      Width           =   9210
      Begin VB.CommandButton BotaoLote 
         Caption         =   "Lote"
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
         Left            =   1479
         TabIndex        =   84
         Top             =   4755
         Width           =   1395
      End
      Begin VB.CommandButton BotaoSerie 
         Caption         =   "Séries"
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
         Left            =   30
         TabIndex        =   83
         Top             =   4755
         Width           =   1395
      End
      Begin VB.ComboBox Requisitante 
         Height          =   315
         Left            =   2985
         TabIndex        =   8
         Top             =   1045
         Width           =   2295
      End
      Begin VB.ComboBox FilialOP 
         Height          =   315
         Left            =   6450
         TabIndex        =   77
         Top             =   1980
         Width           =   2160
      End
      Begin MSMask.MaskEdBox Lote 
         Height          =   270
         Left            =   5415
         TabIndex        =   76
         Top             =   2040
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1980
         Picture         =   "ReqConsumo.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   150
         Width           =   300
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1785
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1965
         Width           =   855
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   6180
         MaxLength       =   50
         TabIndex        =   25
         Top             =   2670
         Width           =   2600
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
         Height          =   345
         Left            =   5826
         TabIndex        =   12
         Top             =   4755
         Width           =   1605
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
         Height          =   345
         Left            =   2928
         TabIndex        =   10
         Top             =   4755
         Width           =   1395
      End
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
         Height          =   345
         Left            =   4377
         TabIndex        =   11
         Top             =   4755
         Width           =   1395
      End
      Begin VB.CheckBox Estorno 
         Height          =   210
         Left            =   7800
         TabIndex        =   26
         Top             =   2340
         Width           =   870
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
         Height          =   345
         Left            =   7485
         TabIndex        =   13
         Top             =   4755
         Width           =   1605
      End
      Begin MSMask.MaskEdBox ContaContabilEst 
         Height          =   300
         Left            =   4275
         TabIndex        =   24
         Top             =   2970
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox ContaContabilAplic 
         Height          =   300
         Left            =   4305
         TabIndex        =   23
         Top             =   2535
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   529
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
         Height          =   315
         Left            =   2700
         TabIndex        =   19
         Top             =   1980
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   285
         TabIndex        =   22
         Top             =   2385
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   315
         Left            =   3660
         TabIndex        =   20
         Top             =   2055
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   300
         Left            =   5010
         TabIndex        =   21
         Top             =   2055
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
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
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   4065
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   142
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   2985
         TabIndex        =   3
         Top             =   142
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
         Height          =   2325
         Left            =   270
         TabIndex        =   9
         Top             =   1830
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   4101
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   945
         TabIndex        =   1
         Top             =   135
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Hora 
         Height          =   300
         Left            =   6705
         TabIndex        =   5
         Top             =   142
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclPadrao 
         Height          =   315
         Left            =   2985
         TabIndex        =   6
         Top             =   590
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AlmoxPadrao 
         Height          =   315
         Left            =   6705
         TabIndex        =   7
         Top             =   590
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Requisitante:"
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
         Left            =   1770
         TabIndex        =   82
         Top             =   1110
         Width           =   1140
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   81
         Top             =   650
         Width           =   2670
      End
      Begin VB.Label AlmoxPadraoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado Padrão:"
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
         Left            =   4845
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   80
         Top             =   650
         Width           =   1815
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
         Left            =   6180
         TabIndex        =   78
         Top             =   195
         Width           =   480
      End
      Begin VB.Label QuantDisponivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2340
         TabIndex        =   52
         Top             =   4290
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Disponível:"
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
         Left            =   255
         TabIndex        =   51
         Top             =   4335
         Width           =   2025
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Consumo de Material"
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
         Left            =   240
         TabIndex        =   50
         Top             =   1560
         Width           =   1785
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
         Left            =   2430
         TabIndex        =   49
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   48
         Top             =   195
         Width           =   660
      End
   End
   Begin VB.CheckBox ImprimeAoGravar 
      Caption         =   "Imprimir ao Gravar"
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
      Left            =   4665
      TabIndex        =   87
      Top             =   165
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5145
      Index           =   2
      Left            =   240
      TabIndex        =   27
      Top             =   780
      Visible         =   0   'False
      Width           =   9090
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4800
         TabIndex        =   85
         Tag             =   "1"
         Top             =   2520
         Width           =   870
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
         Left            =   6390
         TabIndex        =   33
         Top             =   345
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
         Left            =   6390
         TabIndex        =   31
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   840
         Width           =   2700
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
         Left            =   7845
         TabIndex        =   32
         Top             =   30
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   5160
         TabIndex        =   41
         Top             =   1320
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
         TabIndex        =   43
         Top             =   2025
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   42
         Top             =   1635
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2595
         Left            =   6360
         TabIndex        =   45
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   53
         Top             =   3495
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   55
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   54
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
         Left            =   3465
         TabIndex        =   36
         Top             =   930
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   37
         Top             =   1320
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
         TabIndex        =   40
         Top             =   1350
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
         TabIndex        =   39
         Top             =   1290
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
         TabIndex        =   38
         Top             =   1335
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
         Left            =   1635
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   30
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
         Left            =   5580
         TabIndex        =   29
         Top             =   150
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
         TabIndex        =   28
         Top             =   135
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
         TabIndex        =   44
         Top             =   1185
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
         Height          =   2985
         Left            =   6360
         TabIndex        =   46
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6360
         TabIndex        =   47
         Top             =   1560
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
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
         Left            =   6450
         TabIndex        =   34
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
         TabIndex        =   74
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   73
         Top             =   135
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
         TabIndex        =   72
         Top             =   615
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   71
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   70
         Top             =   570
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
         TabIndex        =   69
         Top             =   600
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
         TabIndex        =   68
         Top             =   960
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
         Top             =   3060
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   63
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   62
         Top             =   3045
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
         TabIndex        =   61
         Top             =   570
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
         TabIndex        =   60
         Top             =   180
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
         TabIndex        =   59
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6840
      ScaleHeight     =   495
      ScaleWidth      =   2520
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   90
      Width           =   2580
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   75
         Picture         =   "ReqConsumo.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   525
         Picture         =   "ReqConsumo.ctx":01EC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1530
         Picture         =   "ReqConsumo.ctx":0346
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2040
         Picture         =   "ReqConsumo.ctx":0878
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1035
         Picture         =   "ReqConsumo.ctx":09F6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5670
      Left            =   150
      TabIndex        =   75
      Top             =   405
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10001
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
Attribute VB_Name = "ReqConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjAnotacao As ClassAnotacoes

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoRastroLote As AdmEvento 'Inserido por Wagner
Attribute objEventoRastroLote.VB_VarHelpID = -1

'mnemonicos
Private Const CODIGO1 As String = "Codigo"
Private Const DATA1 As String = "Data"
'###########GridMovimentos###########'
Private Const ESTORNO1 As String = "Estorno"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const CCL1 As String = "Ccl"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const CONTACONTABILEST1 As String = "ContaContabilEst"
Private Const CONTACONTABILAPLIC1 As String = "ContaContabilAplic"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"

Dim gcolcolRastreamentoSerie As Collection 'Inserido por Wagner
Dim iTipoMovtoAnt As Integer

Dim iLinhaAntiga As Integer
Dim colItensNumIntDoc As Collection

Dim objGrid As AdmGrid

Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Estorno_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_ContaContabilEst_Col As Integer
Dim iGrid_ContaContabilAplic_Col As Integer
Dim iGrid_Lote_Col As Integer
Dim iGrid_FilialOP_Col As Integer

Private WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxPadrao As AdmEvento
Attribute objEventoAlmoxPadrao.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1

Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim lCodigoAntigo As Long

'Constantes públicas dos tabs
Private Const TAB_Movimentos = 1
Private Const TAB_Contabilizacao = 2

Private Sub BotaoExcluir_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 89921

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    'Exclui a requisição de consumo
    lErro = CF("MovimentoEstoque_Trata_Exclusao", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 89922

    'Limpa Tela
    Call Limpa_Tela_Consumo

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 89921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 89922
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173935)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long, lCodigo As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 30217
    
    lCodigo = CLng(Codigo.Text)

    lErro = objRelatorio.ExecutarDireto("Requisições Para Consumo", "MovEstCod = " & CStr(lCodigo), 0, "", "NMOVESTCOD", CStr(lCodigo), "TPRODINIC", "", "TPRODFIM", "", "TCCLINIC", "", "TCCLFIM", "", "DINIC", Forprint_ConvData(DATA_NULA), "DFIM", Forprint_ConvData(DATA_NULA))
    If lErro <> SUCESSO Then gError 30225
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 30217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173936)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("MovEstoque_Automatico", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then gError 57525

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    lCodigoAntigo = lCodigo
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 57525
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173936)
    
    End Select

    Exit Sub

End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

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

Private Sub AlmoxPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AlmoxPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadrao_Validate

    If Len(Trim(AlmoxPadrao.Text)) = 0 Then Exit Sub

    lErro = TP_Almoxarifado_Filial_Le(AlmoxPadrao, objAlmoxarifado, 0)
    If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 30216

    If lErro = 25136 Then gError 22909
    
    If lErro = 25143 Then gError 22910
     
    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 30216
            
        Case 22909, 22910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173937)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxPadraoLabel_Click()

Dim colSelecao As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxPadrao)

End Sub

Private Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovimentos.Row = 0 Then gError 43766
    
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = "" Then gError 43767

    sConta = String(STRING_CONTA, 0)

    'Verifica através da coluna que está preenchida
    If GridMovimentos.Col = iGrid_ContaContabilEst_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEst.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 43768
    
    ElseIf GridMovimentos.Col = iGrid_ContaContabilAplic_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilAplic.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 43769
    
    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 43766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 43767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 43768, 43769

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173938)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub ContaContabilAplic_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabilAplic_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ContaContabilAplic_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ContaContabilAplic_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabilAplic
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

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

Private Sub CclPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclPadrao_Validate

    If Len(Trim(CclPadrao.Text)) = 0 Then Exit Sub

    lErro = CF("Ccl_Critica", CclPadrao.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError 30214

    If lErro = 5703 Then gError 30215

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 30214

        Case 30215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173939)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_colItensNumIntDoc(colItensNumIntDoc As Collection)

Dim lErro As Long
Dim iCount As Integer
Dim iIndice As Integer

On Error GoTo Erro_Limpa_colItensNumIntDoc

    iCount = colItensNumIntDoc.Count
    Set colItensNumIntDoc = New Collection
    
    For iIndice = 1 To iCount
        
        colItensNumIntDoc.Add 0
        GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "0"
        
    Next
    
    Exit Sub
    
Erro_Limpa_colItensNumIntDoc:

    Select Case gErr
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173940)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 57761
    
    End If
    
    'se o codigo foi trocado
    If lCodigoAntigo <> StrParaLong(Trim(Codigo.Text)) Then

        If Len(Trim(Codigo.ClipText)) > 0 Then
      
            Call Limpa_colItensNumIntDoc(colItensNumIntDoc)
            
            objMovEstoque.lCodigo = Codigo.Text
            
            'Le o Movimento de Estoque e Verifica se ele já foi estornado
            lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_CONSUMO)
            If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 34894
            
            'Se todos os Itens do Movimento foram estornados
            If lErro = 78885 Then gError 78887
            
            If lErro = SUCESSO Then
    
                If objMovEstoque.iTipoMov <> MOV_EST_CONSUMO Then gError 34897
                
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")
                
                If vbMsg = vbNo Then gError 34895
                
                lErro = Preenche_Tela(objMovEstoque)
                If lErro <> SUCESSO Then gError 34896
          
            End If
      
        End If
      
        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))
    
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
            
        Case 34894, 34896
        
        Case 34895, 57761
            
        Case 34897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_CONSUMO", gErr, objMovEstoque.lCodigo)
            lCodigoAntigo = 0
        
        Case 78887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173941)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 30200

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 30200

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173942)

    End Select

    Exit Sub

End Sub

'hora
Public Sub Hora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Hora, iAlterado)

End Sub

'hora
Public Sub Hora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'hora
Public Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se a hora foi digitada
    If Len(Trim(Hora.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(Hora.Text)
    If lErro <> SUCESSO Then gError 89807

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 89807

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173943)

    End Select

    Exit Sub

End Sub

Private Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Estorno_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Estorno_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Estorno_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Estorno
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Lote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub FilialOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = FilialOP
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub objEventoAlmoxPadrao_evSelecao(obj1 As Object)

Dim objAlmoxarifado As New ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    'Preenche o Almoxarifado Padrao
    AlmoxPadrao.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

End Sub

Private Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    If GridMovimentos.Row = 0 Then gError 43775

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 43776
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub
    
Erro_BotaoCcls_Click:
    
    Select Case gErr
    
        Case 43775
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43776
             lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173944)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridMovimentos.Row = 0 Then gError 43779
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 43780
    
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 43781

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado

        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)

    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr
    
        Case 43779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 43781

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173945)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If GridMovimentos.Col <> iGrid_ContaContabilEst_Col And GridMovimentos.Col <> iGrid_ContaContabilAplic_Col Then
        Me.Show
        Exit Sub
    End If
        
    If objPlanoConta.sConta <> "" Then
   
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 43761
        
        If GridMovimentos.Col = iGrid_ContaContabilEst_Col Then
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
        Else
            ContaContabilAplic.PromptInclude = False
            ContaContabilAplic.Text = sContaEnxuta
            ContaContabilAplic.PromptInclude = True
        End If

        GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = objGrid.objControle.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 43761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173946)

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
    If lErro <> SUCESSO Then gError 30405

    'Verifica se o produto está preenchido e se a linha corrente é diferente da linha fixa
    If iProdutoPreenchido = PRODUTO_PREENCHIDO And GridMovimentos.Row <> 0 Then

        'Preenche o Nome do Almoxarifado
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        
        'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
        lErro = Preenche_ContaContabilEst()
        If lErro <> SUCESSO Then gError 49681
        
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
            'Calcula a Quantidade Disponível nesse Almoxarifado
            lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
            If lErro <> SUCESSO Then gError 30124
        Else
            lErro = QuantDisponivelLote_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 78850
        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 30124, 30405, 49681, 78850

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173947)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 30186

    'Limpa Tela
    Call Limpa_Tela_Consumo

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 30186

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173948)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 30188

    Call Limpa_Tela_Consumo

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 30188

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173949)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoProdutos_Click

    If GridMovimentos.Row = 0 Then gError 43777

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
        
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 22923
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProduto
    End If
    
    sSelecao = "ControleEstoque<>?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    
    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_BotaoProdutos_Click:

     Select Case gErr
     
        Case 22923
     
        Case 43777
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
     
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173950)
     
     End Select

    Exit Sub

End Sub

Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclPadraoLabel_Click()

Dim colSelecao As New Collection
Dim objCcl As ClassCcl

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoEnxuto As String
Dim iProdutoPreenchido As Integer
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30122

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoEnxuto = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 30388

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 30387

            If lErro = 28030 Then gError 30406
            
            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is Produto) Then
            
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
    
                If Trim(objProduto.sContaContabil) = "" Then
                    
                    objTipoDeProduto.iTipo = objProduto.iTipo
                    
                    lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                    If lErro <> SUCESSO And lErro <> 22531 Then gError 49999
                    
                    If lErro = 22531 Then gError 52000
                    
                    objProduto.sContaContabil = objTipoDeProduto.sContaContabil
                                
                End If
                
                'preenche a linha do produtos com seus dados padrões
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 30123
                
                'se estiver preenchido o Almoxarifado e o produto então preenche a conta de estoque
                lErro = Preenche_ContaContabilEst()
                If lErro <> SUCESSO Then gError 52228
                
                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
                    lErro = QuantDisponivel_Calcula1(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
                    If lErro <> SUCESSO Then gError 30363
                Else
                    lErro = QuantDisponivelLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78852
                End If
                
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 30122, 30123, 30363, 30387, 30388, 49999, 52228

        Case 30406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 52000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173951)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)

Dim objCcl As New ClassCcl
Dim lErro As Long
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclPadrao_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 30240

    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 30240

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173952)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) <> 0 And GridMovimentos.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 30241

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

        Case 30241

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173953)

    End Select

    Exit Sub

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

Private Sub Estorno_Click()
  
Dim lErro As Long
  
On Error GoTo Erro_Estorno_Click
  
    iAlterado = REGISTRO_ALTERADO
       
    '############################################
    'Inserido por Wagner 15/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 177297
    '############################################
    
    Exit Sub
    
Erro_Estorno_Click:

    Select Case gErr
    
        Case 177297
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177298)

    End Select
    
    Exit Sub
       
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCclPadrao As String
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner
    
    Set colItensNumIntDoc = New Collection

    Set objEventoCclPadrao = New AdmEvento
    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCodigo = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    
    Set objEventoRastroLote = New AdmEvento

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 30213

    'Inicializa Máscara de CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)
    
    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 30190

    CclPadrao.Mask = sMascaraCclPadrao
    Ccl.Mask = sMascaraCclPadrao

    Quantidade.Format = FORMATO_ESTOQUE
        
    'Inicializa mascara de contaContabilEst
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEst)
    If lErro <> SUCESSO Then gError 49664
    
    'Inicializa mascara de contaContabilAplic
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilAplic)
    If lErro <> SUCESSO Then gError 49665
    
    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    '################################
    'Inserido por Wagner
    Call Carrega_Requisitante
    '################################

    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 78296
    
    'Inicialização do GridMovimentos
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridMovimentos(objGrid)
    If lErro <> SUCESSO Then gError 30109

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE)
    If lErro <> SUCESSO Then gError 39629
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 30109, 30190, 30213, 39629, 49664, 49665, 78296

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173954)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 78293

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOP.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOP.ItemData(FilialOP.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 78293
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173955)

    End Select

    Exit Function

End Function

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
    objGridInt.colColuna.Add ("Lote/Serie Ini.")
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil Estoque")
    objGridInt.colColuna.Add ("Conta Contábil Aplicação")
    objGridInt.colColuna.Add ("Estorno")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ContaContabilEst.Name)
    objGridInt.colCampo.Add (ContaContabilAplic.Name)
    objGridInt.colCampo.Add (Estorno.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_Almoxarifado_Col = 5
    iGrid_Lote_Col = 6
    iGrid_FilialOP_Col = 7
    iGrid_Ccl_Col = 8
    iGrid_ContaContabilEst_Col = 9
    iGrid_ContaContabilAplic_Col = 10
    iGrid_Estorno_Col = 11
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

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

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MovimentoEstoque"

    'Lê os atributos de objMovEstoque que aparecem na Tela
    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    If Len(Data.ClipText) <> 0 Then
        objMovEstoque.dtData = CDate(Data.Text)

    Else
        objMovEstoque.dtData = DATA_NULA

    End If

    If Len(Trim(Hora.ClipText)) > 0 Then
        objMovEstoque.dtHora = CDate(Hora.Text)
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
    colSelecao.Add "TipoMov", OP_IGUAL, MOV_EST_CONSUMO
    colSelecao.Add "NumIntDocEst", OP_IGUAL, 0

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173956)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objMovEstoque.lCodigo = colCampoValor.Item("Codigo").vValor
    objMovEstoque.dtData = colCampoValor.Item("Data").vValor
    objMovEstoque.dtHora = colCampoValor.Item("Hora").vValor

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 30111

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 30111

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173957)

    End Select

    Exit Sub

End Sub

Function Preenche_Tela(objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Tela

    '##########################################
    'Alterado por Wagner
    'Limpa a Tela sem fechar o comando de Setas
    'Função genérica para Limpar a Tela
    'Call Limpa_Tela(Me)

    'Limpa o Grid
    'Call Grid_Limpa(objGrid)
    
    Call Limpa_Tela_Consumo
    '############################################

    'Se o grid permite excluir e incluir Linhas
    If objGrid.iProibidoIncluir <> GRID_PROIBIDO_INCLUIR And objGrid.iProibidoExcluir <> GRID_PROIBIDO_EXCLUIR Then
        'prepara o Grid para não permitir inserir e excluir Linhas
        objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
        objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
        Call Grid_Inicializa(objGrid)
    End If
    
    'Limpa o Label QuantDisponivel
    QuantDisponivel.Caption = ""
    
    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection
    Set objMovEstoque.colItens = New ColItensMovEstoque
    
    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovEstoqueItens_Le1", objMovEstoque, MOV_EST_CONSUMO)
    If lErro <> SUCESSO And lErro <> 55387 Then gError 30117

    If lErro = 55387 Then gError 55401

    'Passa as Informações de NumIntDoc de colItens para colItensNumIntDoc
    For iIndice = 1 To objMovEstoque.colItens.Count
    
        colItensNumIntDoc.Add objMovEstoque.colItens.Item(iIndice).lNumIntDoc

    Next

    'Coloca os Dados na Tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objMovEstoque.lCodigo)
    Codigo.PromptInclude = True

    If objMovEstoque.dtData <> DATA_NULA Then
        Data.PromptInclude = False
        Data.Text = Format(objMovEstoque.dtData, "dd/mm/yy")
        Data.PromptInclude = True

    Else
        Data.PromptInclude = False
        Data.Text = ""
        Data.PromptInclude = True

    End If

'hora
    Hora.PromptInclude = False
    'este teste está correto
    If objMovEstoque.dtData <> DATA_NULA Then Hora.Text = Format(objMovEstoque.dtHora, "hh:mm:ss")
    Hora.PromptInclude = True

    '#############################################################
    'Inserido por Wagner
    For iIndice = 0 To Requisitante.ListCount - 1
        If Requisitante.ItemData(iIndice) = objMovEstoque.lRequisitante Then
            Requisitante.ListIndex = iIndice
            Exit For
        End If
    Next
    '#############################################################


    lErro = Preenche_GridMovimentos(objMovEstoque.colItens)
    If lErro <> SUCESSO Then gError 30118
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovEstoque.colItens.Item(1).lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39630

    iAlterado = 0
    lCodigoAntigo = objMovEstoque.lCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 30117, 30118, 39630

        Case 55401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_REQCONSUMO", gErr, objMovEstoque.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173958)

    End Select

    Exit Function

End Function

Private Sub CodigoLabel_Click()

Dim colSelecao As New Collection
Dim objMovEstoque As New ClassMovEstoque

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    'Adiciona filtro
    colSelecao.Add MOV_EST_CONSUMO

    Call Chama_Tela("MovEstoqueLista", colSelecao, objMovEstoque, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objMovEstoque = obj1

    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO Then gError 30120

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 30191

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 30120, 30191

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173959)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set colItensNumIntDoc = Nothing
    
    Set objEventoCclPadrao = Nothing
    Set objEventoAlmoxPadrao = Nothing
    Set objEventoCcl = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoRastroLote = Nothing 'Inserido por Wagner
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    Set gcolcolRastreamentoSerie = Nothing 'Inserido por Wagner
    
    Set gobjAnotacao = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Function Trata_Parametros(Optional objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objMovEstoque Is Nothing Then

        objMovEstoque.iFilialEmpresa = giFilialEmpresa
        
        'Lê MovEstoque no Banco de Dados
        lErro = CF("MovEstoque_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30128 Then gError 30129

        If lErro <> 30128 Then 'Se ele existe

            If objMovEstoque.iTipoMov <> MOV_EST_CONSUMO Then gError 30130

            lErro = Preenche_Tela(objMovEstoque)
            If lErro <> SUCESSO Then gError 30131

        Else
            'Se ele não existe exibe apenas o código
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objMovEstoque.lCodigo)
            Codigo.PromptInclude = True

            lCodigoAntigo = objMovEstoque.lCodigo

        End If

    End If

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 30129, 30131

        Case 30130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_REQCONSUMO", gErr, objMovEstoque.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173960)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer, lNumIntDoc As Long
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30143
    
    If colItensNumIntDoc.Count >= GridMovimentos.Row Then
        lNumIntDoc = colItensNumIntDoc.Item(GridMovimentos.Row)
    Else
        lNumIntDoc = 0
    End If
    
    If objControl.Name = "Produto" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
           objControl.Enabled = False

        Else
            objControl.Enabled = True

        End If

    ElseIf objControl.Name = "UnidadeMed" Then

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
            
            objControl.Enabled = False

        Else
            
            objControl.Enabled = True

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 30145

            If lErro = 28030 Then gError 30146

            objClasseUM.iClasse = objProduto.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 30147

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)

            'Limpar as Unidades utilizadas anteriormente
            UnidadeMed.Clear

            For Each objUnidadeDeMedida In colSiglas
                UnidadeMed.AddItem objUnidadeDeMedida.sSigla
            Next

            UnidadeMed.AddItem ""

            'Tento selecionar na Combo a Unidade anterior
            If UnidadeMed.ListCount <> 0 Then

                For iIndice = 0 To UnidadeMed.ListCount - 1

                    If UnidadeMed.List(iIndice) = sUnidadeMed Then
                        UnidadeMed.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If

            If lNumIntDoc = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        End If

    ElseIf objControl.Name = "Quantidade" Or objControl.Name = "Almoxarifado" Or objControl.Name = "Ccl" Or objControl.Name = "ContaContabilEst" Or objControl.Name = "ContaContabilAplic" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
                
    ElseIf objControl.Name = "Lote" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78211
    
            If lErro = 28030 Then gError 78212
        
            If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Else
            objControl.Enabled = False
        End If
                
    ElseIf objControl.Name = "FilialOP" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78321
    
            If lErro = 28030 Then gError 78322
        
            If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        Else
            objControl.Enabled = False
        End If
    
    ElseIf objControl.Name = "Estorno" Then
    
        If lNumIntDoc = 0 Then
        
            objControl.Enabled = False

        Else
            objControl.Enabled = True

        End If
    
    End If
                    
    If gobjCRFAT.iUsaBloqAcessoPorTelaControle = MARCADO Then
        lErro = CF("Rotina_Grid_Enable_BloqueiaAcesso", Me.Name, objControl)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 30143, 30145, 30146, 30147, 78211, 78212, 78321, 78322
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173961)

    End Select

    Exit Sub

End Sub

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

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer, lNumIntDoc As Long

    If colItensNumIntDoc.Count >= GridMovimentos.Row Then
        lNumIntDoc = colItensNumIntDoc.Item(GridMovimentos.Row)
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
    
            'Exclui de colItensNumIntDoc o Item correspondente, se houver
            colItensNumIntDoc.Remove iLinhaAnterior
            gcolcolRastreamentoSerie.Remove iLinhaAnterior
    
        End If
    
    End If
    
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

Private Sub GridMovimentos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridMovimentos_RowColChange()

Dim lErro As Long

'#################################
'Inserido por Wagner 15/03/2006
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
'#################################

On Error GoTo Erro_GridMovimentos_RowColChange

    Call Grid_RowColChange(objGrid)

    If (GridMovimentos.Row <> iLinhaAntiga) Then

        'Guarda a Linha corrente
        iLinhaAntiga = GridMovimentos.Row
        
        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141946
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141947
        '###########################################################
        
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 30151
        Else
            lErro = QuantDisponivelLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78853
        End If
    
    End If

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case 30151, 78853, 141946

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173962)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

'#################################
'Inserido por Wagner 15/03/2006
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
'#################################

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    If Len(UnidadeMed.Text) > 0 Then

        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141949
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141950
        '###########################################################

        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 55397
        Else
            lErro = QuantDisponivelLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78854
        End If
    
        If colItensNumIntDoc.Item(GridMovimentos.Row) = 0 Then
    
            'Se a quantidade está preenchida e não se trata de estorno
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    
                'Testa a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError 55398
    
            End If
    
        End If

    Else
    
        QuantDisponivel.Caption = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30203

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 30203, 55397, 55398, 78854, 141950
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173963)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilAplic(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilAplic
    
    Set objGrid.objControle = ContaContabilAplic
    
    If Len(Trim(ContaContabilAplic.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilAplic.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49632
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49633
            
            ContaContabilAplic.PromptInclude = False
            ContaContabilAplic.Text = sContaMascarada
            ContaContabilAplic.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilAplic.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49634
    
            'conta não cadastrada
            If lErro = 5700 Then gError 49635
             
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49647
    
    Saida_Celula_ContaContabilAplic = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilAplic:

    Saida_Celula_ContaContabilAplic = gErr

    Select Case gErr

        Case 49632, 49634, 49647
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 49635
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilAplic.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173964)

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
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49628
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49629
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaMascarada
            ContaContabilEst.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEst.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49630
    
            'conta não cadastrada
            If lErro = 5700 Then gError 49631
             
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49643
    
   Saida_Celula_ContaContabilEst = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEst:

    Saida_Celula_ContaContabilEst = gErr

    Select Case gErr

        Case 49628, 49630, 49643
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 49631
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEst.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173965)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Produto.ClipText) <> 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199349

        Produto.Text = sProduto

        lErro = CF("Produto_Critica_Estoque", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 30160

        If lErro = 25077 Then gError 30161

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            If Trim(objProduto.sContaContabil) = "" Then
                
                objTipoDeProduto.iTipo = objProduto.iTipo
                
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError 49997
                
                If lErro = 22531 Then gError 49998
                
                objProduto.sContaContabil = objTipoDeProduto.sContaContabil
                            
            End If
    
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 30162
    
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
                'Calcula a Quantidade Disponível
                lErro = QuantDisponivel_Calcula1(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
                If lErro <> SUCESSO Then gError 30163
            Else
                'Calcula a Quantidade Disponível do lote
                lErro = QuantDisponivelLote_Calcula1(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)), objProduto)
                If lErro <> SUCESSO Then gError 78855
            End If
    
        End If
    
        If objProduto.iRastro = PRODUTO_RASTRO_OP Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30164
    
    Call Preenche_ContaContabilEst
        
    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 30160, 30162, 30163, 49997, 78855, 199349
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30161
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30164
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49998
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173966)

    End Select

    Exit Function

End Function

Private Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

Dim lErro As Long

    Call Grid_Campo_Recebe_Foco(objGrid)

    If gobjEST.iInventarioCodBarrAuto = 1 Then

        If objGrid.lErroSaidaCelula = 0 Then

            lErro = Trata_CodigoBarras1

            objGrid.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE
            
            Call Grid_Entrada_Celula(objGrid, iAlterado)

            objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

            If lErro <> SUCESSO Then
    
                objGrid.lErroSaidaCelula = 1
            End If

        Else
    
            objGrid.lErroSaidaCelula = 0
    
        End If
        
    End If
    
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

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

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
                Parent.HelpContextID = IDH_REQUISICAO_MATERIAL_CONSUMO_MOVIMENTOS
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_REQUISICAO_MATERIAL_CONSUMO_CONTABILIZACAO
                        
        End Select
    
    End If

End Sub

Private Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.Text)) <> 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 30169

        If colItensNumIntDoc.Item(GridMovimentos.Row) = 0 Then
            
            'se nao for estorno e QuantDisponivel estiver preenchida verificar se é maior
            If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" And Len(Trim(QuantDisponivel.Caption)) <> 0 Then
    
                lErro = Testa_QuantRequisitada(CDbl(Quantidade.Text))
                If lErro <> SUCESSO Then gError 30214
    
            End If
        
        End If

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    '##############################################
    'Inserido por Wagner 15/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(Quantidade.Text), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141911
    '##############################################

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30170

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 30169, 30214, 30170, 141911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173967)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim objProduto As New ClassProduto 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.Text)) <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30503

        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 30171

        If lErro = 25157 Then gError 30172

        If lErro = 25162 Then gError 30173

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141948
        '###########################################################

        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then   'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 30174
        Else
            'Calcula a Quantidade Disponível do lote
            lErro = QuantDisponivelLote_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 78856
        End If

    Else
    
        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30175

    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    Call Preenche_ContaContabilEst
    
    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    'Limpa a Quantidade Disponível da Tela
    QuantDisponivel.Caption = ""
    
    Select Case gErr

        Case 30171, 30174, 30175, 30503, 78856, 141948
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30172

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30173

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173968)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Ccl.Text) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 30183

        If lErro = 5703 Then gError 30184

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30185

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 30183, 30185
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30184
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173969)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote
    
    If Len(Trim(Lote.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78537
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78538
            
        If lErro = 28030 Then gError 78539
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78540
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78541
                
                'Preenche a Quantidade do Lote
                lErro = QuantDisponivelLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78867
                                    
            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                
                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))) > 0 Then
'??? lixo ?
''                    objOrdemProducao.iFilialEmpresa = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))
''                    objOrdemProducao.sCodigo = Lote.Text
''
''                    'Verifica se existe a OP
''                    lErro = CF("OrdemProducao_Le",objOrdemProducao)
''                    If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 78542
''
''                    If lErro = 30368 Then gError 78543
''
''                    If lErro = 55316 Then gError 78544
''
                    objRastroLote.sCodigo = Lote.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))
                    
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 78545
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 78546
                    
                    'Preenche a Quantidade do Lote
                    lErro = QuantDisponivelLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78868
                
                End If
                
            '###############################################################
            'Inserido por Wagner 15/03/2006
            'Se for rastro por série
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                'Preenche a Quantidade do Lote
                lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError 78869
            '###############################################################
                
            End If
        
        End If
    
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 78869

    End If
            
    'Se a quantidade está preenchida e não se trata de estorno
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" Then

        dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))

        'Testa a Quantidade requisitada
        lErro = Testa_QuantRequisitada(dQuantidade)
        If lErro <> SUCESSO Then gError 78872

    End If
        
    '############################################
    'Inserido por Wagner 15/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141912
    '############################################
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78547

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 78537, 78538, 78540, 78542, 78545, 78547, 78867, 78868, 78869, 78872, 141912
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78541, 78546
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 78543
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173970)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long
'Faz a saida de celula da Filial da Ordem de Produção

Dim lErro As Long
Dim objFilialOP As New AdmFiliais
Dim iCodigo As Integer
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP

    If Len(Trim(FilialOP.Text)) <> 0 Then
            
        'Verifica se é uma FilialOP selecionada
        If FilialOP.Text <> FilialOP.List(FilialOP.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOP, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 78252
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 78253
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 78254
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
            
            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 78255
                    
        End If
        
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) > 0 Then
'??? lixo ?
''            objOrdemProducao.iFilialEmpresa = Codigo_Extrai(FilialOP.Text)
''            objOrdemProducao.sCodigo = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)
''
''            'Verifica se existe a OP
''            lErro = CF("OrdemProducao_Le",objOrdemProducao)
''            If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 78581
''
''            If lErro = 30368 Then gError 78582
''
''            If lErro = 55316 Then gError 78583
''
            lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 78584
                                
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objRastroLote.sCodigo = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)
                objRastroLote.sProduto = sProdutoFormatado
                objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
            
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78585
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78586
                            
                'Preenche a Quantidade do Lote
                lErro = QuantDisponivelLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(FilialOP.Text))
                If lErro <> SUCESSO Then gError 78870
                
            End If
            
        End If
        
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 78871
        
    End If

    'Se a quantidade está preenchida e não se trata de estorno
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" Then

        dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))

        'Testa a Quantidade requisitada
        lErro = Testa_QuantRequisitada(dQuantidade)
        If lErro <> SUCESSO Then gError 78873

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78256

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 78252, 78253, 78256, 78581, 78584, 78585, 78870, 78873
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78582
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78586
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173971)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Consumo()

Dim lErro As Long, lCodigo As Long
On Error GoTo Erro_Limpa_Tela_Consumo

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)
    
    '#######################
    'Inserido por Wagner
    Requisitante.ListIndex = -1
    '#######################

    'Limpa o Label QuantDisponivel
    QuantDisponivel.Caption = ""

    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    If objGrid.iProibidoIncluir <> 0 And objGrid.iProibidoExcluir <> 0 Then
        'prepara o Grid para permitir inserir e excluir Linhas
        objGrid.iProibidoIncluir = 0
        objGrid.iProibidoExcluir = 0
        Call Grid_Inicializa(objGrid)
    End If
    
    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection

    'Coloca código na Tela
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True

    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

   Set gcolcolRastreamentoSerie = New Collection

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0
    lCodigoAntigo = 0
    
    Set gobjAnotacao = Nothing

    Exit Sub
    
Erro_Limpa_Tela_Consumo:

    Select Case gErr
             
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173972)
     
    End Select
    
    Exit Sub
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39686

        If objGridInt.objGrid Is GridMovimentos Then

            Select Case objGridInt.objGrid.Col

                Case iGrid_Almoxarifado_Col
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 30193

                Case iGrid_Ccl_Col
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 30194

                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 30196

                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 30197

                Case iGrid_Estorno_Col
                    lErro = Saida_Celula_Estorno(objGridInt)
                    If lErro <> SUCESSO Then gError 30206

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 30207
                    
                Case iGrid_ContaContabilEst_Col
                    lErro = Saida_Celula_ContaContabilEst(objGridInt)
                    If lErro <> SUCESSO Then gError 49719
                    
                Case iGrid_ContaContabilAplic_Col
                    lErro = Saida_Celula_ContaContabilAplic(objGridInt)
                    If lErro <> SUCESSO Then gError 49720
                        
                Case iGrid_Lote_Col
                    lErro = Saida_Celula_Lote(objGridInt)
                    If lErro <> SUCESSO Then gError 78208

                Case iGrid_FilialOP_Col
                    lErro = Saida_Celula_FilialOP(objGridInt)
                    If lErro <> SUCESSO Then gError 78275
            
            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 30198
    
    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 30193, 30194, 30196, 30197, 30206, 30207, 49719, 49720, 78208, 78275

        Case 30198
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39686
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173973)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMovimentos(colItens As ColItensMovEstoque) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sProdutoMascarado As String, sCclMascarado As String
Dim objItemMovEstoque As ClassItemMovEstoque
Dim sContaEnxutaEst As String
Dim sContaEnxutaAplic As String
Dim colRatreamentoMovto As Collection
Dim objRatreamentoMovto As New ClassRastreamentoMovto
Dim objFilialOP As New AdmFiliais
Dim colRastreamentoSerie As Collection 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_Preenche_GridMovimentos

    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 15/03/2006

    'Preenche GridMovimentos
    For Each objItemMovEstoque In colItens

        iIndice = iIndice + 1
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProduto, sProdutoMascarado) 'Alterado por Wagner REPLICAR_ACERTO
        If lErro <> SUCESSO Then gError 30885

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        If objItemMovEstoque.sContaContabilEst <> "" Then
        
            sContaEnxutaEst = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilEst, sContaEnxutaEst)
            If lErro <> SUCESSO Then gError 49715
        
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxutaEst
            ContaContabilEst.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
            
        End If
        
        If objItemMovEstoque.sContaContabilAplic <> "" Then
        
            sContaEnxutaAplic = String(STRING_CONTA, 0)
        
            lErro = Mascara_MascararConta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaAplic)
            If lErro <> SUCESSO Then gError 49716
            
            ContaContabilAplic.PromptInclude = False
            ContaContabilAplic.Text = sContaEnxutaAplic
            ContaContabilAplic.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilAplic_Col) = ContaContabilAplic.Text
        
        End If
        
        GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridMovimentos.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemMovEstoque.sProdutoDesc
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemMovEstoque.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemMovEstoque.dQuantidade)
        GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemMovEstoque.sAlmoxarifadoNomeRed
        
        If objItemMovEstoque.sCcl <> "" Then
        
            sCclMascarado = String(STRING_CCL, 0)
        
            lErro = Mascara_MascararCcl(objItemMovEstoque.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 22911
            
        Else
        
            sCclMascarado = ""
            
        End If
        
        GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        Set colRatreamentoMovto = New Collection
        
        'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
        lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRatreamentoMovto)
        If lErro <> SUCESSO And lErro <> 78414 Then gError 78611
        
        'Se existe rastreamento
        If colRatreamentoMovto.Count > 0 Then
                        
            'Seta o primeiro Lote
            Set objRatreamentoMovto = colRatreamentoMovto(1)
            
            gcolcolRastreamentoSerie.Add objRatreamentoMovto.colRastreamentoSerie 'Inserido por Wagner 15/03/2006
            
            If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridMovimentos.TextMatrix(iIndice, iGrid_Lote_Col) = objRatreamentoMovto.sLote
            
            If objRatreamentoMovto.iFilialOP > 0 Then
            
                objFilialOP.iCodFilial = objRatreamentoMovto.iFilialOP

                'Le a Filial Empresa da OP para pegar a descrição
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO Then gError 78860

                GridMovimentos.TextMatrix(iIndice, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
            
            End If
        
        '#####################################################
        'Inserido por Wagner 15/03/2006
        Else
            Set colRastreamentoSerie = New Collection
            gcolcolRastreamentoSerie.Add colRastreamentoSerie
        '#####################################################
        
        End If
        
    Next

    objGrid.iLinhasExistentes = colItens.Count

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 30234

    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 22911, 30234, 49715, 49716, 78611, 78860

        Case 30885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemMovEstoque.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173974)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iAlmoxarifadoPadrao As Integer
Dim sContaEnxutaAplic As String
Dim sAlmoxarifadoPadrao As String
Dim colRastreamentoSerie As New Collection 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_ProdutoLinha_Preenche

    'Preenche Linha Corrente
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    lErro = Preenche_ContaContabilEst()
    If lErro <> SUCESSO Then gError 49679
    
    If Len(Trim(objProduto.sContaContabil)) > 0 Then
    
        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabil, sContaEnxutaAplic)
        If lErro <> SUCESSO Then gError 49680
        
        ContaContabilAplic.PromptInclude = False
        ContaContabilAplic.Text = sContaEnxutaAplic
        ContaContabilAplic.PromptInclude = True
    
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilAplic_Col) = ContaContabilAplic.Text
        
    End If
    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Descricao_Col) = objProduto.sDescricao
    
    'Almoxarifado
    '(Utiliza Almoxarifado Padrão caso esteja preenchido)
    If Len(Trim(AlmoxPadrao.ClipText)) > 0 Then 'And Len(Trim(Almoxarifado.ClipText)) = 0 Then
        lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadrao)
        If lErro = SUCESSO Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
        Else
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""
        End If
    Else

        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 52227

        'preenche o grid
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao


    End If

    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError 30168

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = CclPadrao.Text

    'Preenche Estorno com Valor 0 (Checked = False)
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) = "0"

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        colItensNumIntDoc.Add 0
        gcolcolRastreamentoSerie.Add colRastreamentoSerie 'Inserido por Wagner 15/03/2006
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 30168, 30403, 46679, 49680, 52227

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173975)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer, vbMsg As VbMsgBoxResult
Dim iMovimento As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsgRes As VbMsgBoxResult
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 30217

    'Verifica se a Data foi preenchida
    If Len(Data.ClipText) = 0 Then gError 30218

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridMovimentos
    If objGrid.iLinhasExistentes = 0 Then gError 30219

    'Para cada MovEstoque
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 30220

        'Verifica se o Almoxarifado foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 30221

        'Verifica se a Unidade de Medida foi preenchida
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 55422

    Next

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30128 Then gError 30882
    
    If lErro = SUCESSO Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS2")
        If vbMsgRes = vbNo Then gError 78865
    
    End If

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 92030

    lErro = Move_Tela_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError 30224
    
    'Grava os dados no BD (inclusive os dados Contábeis)
    lErro = CF("MovEstoque_Grava_Generico", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 30225
    
    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
            gobjAnotacao.sID = CStr(objMovEstoque.iFilialEmpresa) & "," & CStr(objMovEstoque.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 30225
            
        End If
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    'Se a opcao de imprimir o Relatorio estiver marcada
    If ImprimeAoGravar.Value = MARCADO Then
    
        lErro = objRelatorio.ExecutarDireto("Requisições Para Consumo", "MovEstCod = " & CStr(objMovEstoque.lCodigo), 0, "", "NMOVESTCOD", CStr(objMovEstoque.lCodigo), "TPRODINIC", "", "TPRODFIM", "", "TCCLINIC", "", "TCCLFIM", "", "DINIC", Forprint_ConvData(DATA_NULA), "DFIM", Forprint_ConvData(DATA_NULA))
        If lErro <> SUCESSO Then gError 30225
    
    End If
        
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 30179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_MOVTO_ESTOQUE_NAO_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 30180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_ITEM_NAO_CADASTRADO", gErr, iIndice)

        Case 30217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 30218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 30219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 30220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 30221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 30224, 30225, 30882, 30834, 78865, 92030

        Case 30883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVTO_ESTOQUE_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 55422
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173976)

    End Select

    Exit Function

End Function

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

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 30201

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 30201

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173977)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 30202

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 30202

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173978)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Estorno(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Estorno

    Set objGridInt.objControle = Estorno
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30205

    Saida_Celula_Estorno = SUCESSO

    Exit Function

Erro_Saida_Celula_Estorno:

    Saida_Celula_Estorno = gErr

    Select Case gErr

        Case 30205
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173979)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objMovEstoque As ClassMovEstoque) As Long
'Preenche objMovEstoque (inclusive colItens)

Dim iIndice As Integer
Dim lCodigo As Long
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objMovEstoque.iTipoMov = 0
        
    If Len(Trim(Codigo.Text)) <> 0 Then
        objMovEstoque.lCodigo = CLng(Codigo.Text)
    Else
        objMovEstoque.lCodigo = 0
    End If
    
    If Len(Trim(Data.Text)) <> 0 Then
        objMovEstoque.dtData = CDate(Data.Text)
    Else
        objMovEstoque.dtData = DATA_NULA
    End If

'hora
    If Len(Trim(Hora.ClipText)) > 0 Then
        objMovEstoque.dtHora = CDate(Hora.Text)
    Else
        objMovEstoque.dtHora = Time
    End If

    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    '############################
    'Inserido por Wagner
    objMovEstoque.lRequisitante = Codigo_Extrai(Requisitante.Text)
    '############################
    
    For iIndice = 1 To objGrid.iLinhasExistentes

        lErro = Move_Itens_Memoria(iIndice, objMovEstoque)
        If lErro <> SUCESSO Then gError 30222

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 30222

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173980)

    End Select

    Exit Function

End Function

Function Move_Itens_Memoria(iIndice As Integer, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sCcl As String, sCclFormatada As String
Dim iCclPreenchida As Integer
Dim iTipoMov As Integer
Dim objAlmoxarifado As ClassAlmoxarifado
Dim sContaFormatadaEst As String
Dim sContaFormatadaAplic As String
Dim iContaPreenchida As Integer
Dim colRateamentoMovto As New Collection

On Error GoTo Erro_Move_Itens_Memoria

    With GridMovimentos

        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = .TextMatrix(iIndice, iGrid_Almoxarifado_Col)

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 30226

        If lErro = 25056 Then gError 30227

        sProdutoFormatado = ""
        
        sCcl = .TextMatrix(iIndice, iGrid_Ccl_Col)
        
        If Len(Trim(sCcl)) <> 0 Then
        
            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 30235
        
        Else
        
            sCclFormatada = ""

        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilEst_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilEst_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 49661
        
        Else
            sContaFormatadaEst = ""
        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilAplic_Col) <> "" Then
        
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilAplic_Col), sContaFormatadaAplic, iContaPreenchida)
            If lErro <> SUCESSO Then gError 49662
        
        Else
            sContaFormatadaAplic = ""
        End If

        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30236

        'Verifica se a Requisição foi Estornada
        If .TextMatrix(iIndice, iGrid_Estorno_Col) = "1" Then
            iTipoMov = MOV_EST_ESTORNO_CONSUMO

        Else
            iTipoMov = MOV_EST_CONSUMO
        End If
        
        'Move os dados do rastreamento para a Memória
        lErro = Move_RastroEstoque_Memoria(iIndice, colRateamentoMovto)
        If lErro <> SUCESSO Then gError 78244
        
        objMovEstoque.colItens.Add colItensNumIntDoc(iIndice), iTipoMov, 0, 0, sProdutoFormatado, .TextMatrix(iIndice, iGrid_Descricao_Col), .TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(.TextMatrix(iIndice, iGrid_Quantidade_Col)), objAlmoxarifado.iCodigo, .TextMatrix(iIndice, iGrid_Almoxarifado_Col), 0, sCclFormatada, CLng(.TextMatrix(iIndice, iGrid_Estorno_Col)), "", "", sContaFormatadaAplic, sContaFormatadaEst, 0, colRateamentoMovto, Nothing, DATA_NULA

    End With

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case 30226, 30235, 30236, 49661, 49662, 78244

        Case 30227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173981)

    End Select

    Exit Function

End Function

Function Move_RastroEstoque_Memoria(iLinha As Integer, colRastreamentoMovto As Collection) As Long
'Move o Rastro dos Itens de Movimento

Dim objProduto As New ClassProduto, lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRastreamentoMovto As New ClassRastreamentoMovto

On Error GoTo Erro_Move_RastroEstoque_Memoria
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 78230
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 78231

    If lErro = 28030 Then gError 78232
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            
            'Se colocou o Número do Lote
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
                objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col)
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            'se o lote está preenchido e a filial não ==> erro
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
                If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))) = 0 Then gError 78339
                
                objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col)
                objRastreamentoMovto.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))
                
            End If
                
            'se a filial está preenchida e o lote não ==> erro
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))) <> 0 Then
                If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) = 0 Then gError 78233
            End If
            
        '##################################################
        'Inserido por Wagner 15/03/2006
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
           
            For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(iLinha)
                colRastreamentoMovto.Add objRastreamentoMovto
                objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
            Next
        '##################################################
        
        End If

        If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))) > 0 Then objRastreamentoMovto.dQuantidade = CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))
            objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
            objRastreamentoMovto.sProduto = sProdutoFormatado
            
            '######################################################
            'Alterado por Wagner 15/03/2006
            If objProduto.iRastro <> PRODUTO_RASTRO_NUM_SERIE Then
                colRastreamentoMovto.Add objRastreamentoMovto
            End If
            '######################################################
        
        End If
        
    End If
    
    Move_RastroEstoque_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroEstoque_Memoria:

    Move_RastroEstoque_Memoria = gErr
    
    Select Case gErr
        
        Case 78230, 78231
        
        Case 78232
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 78233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 78339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173982)
    
    End Select
    
    Exit Function
    
End Function

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
                    
                    If Len(GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    
                        lErro = CF("UMEstoque_Conversao", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                        If lErro <> SUCESSO Then gError 64204

                        objMnemonicoValor.colValor.Add dQuantidadeConvertida
                    
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                    
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
            
        Case ESTORNO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                objMnemonicoValor.colValor.Add CInt(GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col))
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
                    If lErro <> SUCESSO Then gError 79943
                    
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
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILAPLIC1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilAplic_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilAplic_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
    
            
        Case Else
            Error 39652

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39652
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173983)

    End Select

    Exit Function

End Function

Private Function Preenche_ContaContabilEst() As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabilEst
        
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) > 0 And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
    
        'preenche o objEstoqueProduto
        objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
        
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 49694
        
        If lErro = 25060 Then gError 52003
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 49695
        
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        'Le a conta contabil na tabela de EstoqueProduto se nao encontrar procura na tabela de Almoxarifado
        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 49991 Then gError 49696
            
        If lErro = SUCESSO Then
            
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 49697
                
            'Preenche a Conta Contabil de Estoque
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
                
            'Preenche o Grid
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
        
        End If
    
    End If
    
    Preenche_ContaContabilEst = SUCESSO
    
    Exit Function
    
Erro_Preenche_ContaContabilEst:

    Preenche_ContaContabilEst = gErr
    
    Select Case gErr
        
        Case 49694, 49695, 49696
        
        Case 49697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
        
        Case 52003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173984)
    
    End Select
    
    Exit Function
        
End Function

Private Function QuantDisponivel_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantDisponivel_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado)
        If lErro <> SUCESSO Then gError 55417
        
    Else
    
        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado, objProduto)
        If lErro <> SUCESSO Then gError 55418

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError 55419

    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = gErr

    Select Case gErr

        Case 55417, 55418, 55419

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173985)

    End Select

    Exit Function

End Function

Private Function QuantDisponivel_Calcula1(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long
'descobre a quantidade disponivel e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_QuantDisponivel_Calcula1

    QuantDisponivel.Caption = ""
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30152
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If objProduto Is Nothing Then
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 30153
    
                If lErro = 28030 Then gError 30154
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 30155
    
            If lErro = 25056 Then gError 30156
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 30157
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 30158
        
                QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel * dFator)
    
            End If
    
        Else
    
            'Limpa a Quantidade Disponível da Tela
            QuantDisponivel.Caption = ""
    
        End If

    End If
    
    QuantDisponivel_Calcula1 = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula1:

    QuantDisponivel_Calcula1 = gErr

    Select Case gErr

        Case 30152, 30153, 30155, 30157, 30158

        Case 30154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 30156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173986)

    End Select

    Exit Function

End Function

Private Function QuantDisponivelLote_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantDisponivelLote_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantDisponivelLote_Calcula1(sProduto, sAlmoxarifado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78840
        
    Else
    
        lErro = QuantDisponivelLote_Calcula1(sProduto, sAlmoxarifado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)), objProduto)
        If lErro <> SUCESSO Then gError 78841

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError 78842

    QuantDisponivelLote_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivelLote_Calcula:

    QuantDisponivelLote_Calcula = gErr

    Select Case gErr

        Case 78840, 78841, 78842

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173987)

    End Select

    Exit Function

End Function

Private Function QuantDisponivelLote_Calcula1(sProduto As String, sAlmoxarifado As String, sLote As String, iFilialOP As Integer, Optional objProduto As ClassProduto) As Long
'descobre a quantidade disponivel e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoLoteSaldo As New ClassRastreamentoLoteSaldo

On Error GoTo Erro_QuantDisponivelLote_Calcula1

    QuantDisponivel.Caption = ""
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78843
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If objProduto Is Nothing Then
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 78844
    
                If lErro = 28030 Then gError 78845
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 78846
    
            If lErro = 25056 Then gError 78847
    
            objRastreamentoLoteSaldo.iAlmoxarifado = objAlmoxarifado.iCodigo
            objRastreamentoLoteSaldo.sProduto = sProdutoFormatado
            objRastreamentoLoteSaldo.sLote = sLote
            objRastreamentoLoteSaldo.iFilialOP = iFilialOP
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("RastreamentoLoteSaldo_Le", objRastreamentoLoteSaldo)
            If lErro <> SUCESSO And lErro <> 78633 Then gError 78848
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 78633 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 78849
        
                QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantDispNossa * dFator)
    
            End If
    
        Else
    
            'Limpa a Quantidade Disponível da Tela
            QuantDisponivel.Caption = ""
    
        End If

    End If
    
    QuantDisponivelLote_Calcula1 = SUCESSO

    Exit Function

Erro_QuantDisponivelLote_Calcula1:

    QuantDisponivelLote_Calcula1 = gErr

    Select Case gErr

        Case 78843, 78844, 78846, 78848, 78849

        Case 78845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 78847
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173988)

    End Select

    Exit Function

End Function

Private Function Testa_Quantidade() As Long

Dim dQuantidade As Double
Dim lErro As Long

On Error GoTo Erro_Testa_Quantidade

    If GridMovimentos.Row >= GridMovimentos.FixedRows Then

        If colItensNumIntDoc.Item(GridMovimentos.Row) = 0 Then
    
            'Se a quantidade está preenchida e não se trata de linha estornada
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    
                'Calcula a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError 30212
    
            End If
    
        End If

    End If

    Testa_Quantidade = SUCESSO

    Exit Function

Erro_Testa_Quantidade:

    Testa_Quantidade = gErr

    Select Case gErr

        Case 30212

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173989)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_QuantRequisitada(ByVal dQuantAtual As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoAtual As String
Dim sAlmoxarifado As String
Dim sAlmoxarifadoAtual As String
Dim sUnidadeAtual As String
Dim sUnidadeProd As String
Dim dQuantidadeProd As String
Dim dFator As Double
Dim objProduto As New ClassProduto
Dim dQuantTotal As Double
Dim sLote As String

On Error GoTo Erro_Testa_QuantRequisitada

    If gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then

        sProdutoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
        sAlmoxarifadoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
        sUnidadeAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        sLote = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)
    
        If Len(sProdutoAtual) > 0 And Len(sAlmoxarifadoAtual) > 0 And Len(sUnidadeAtual) > 0 Then
    
            lErro = CF("Produto_Formata", sProdutoAtual, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 55392
    
            objProduto.sCodigo = sProdutoFormatado
    
            'Lê o produto para saber qual é a sua ClasseUM
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 55393
        
            If lErro = 28030 Then gError 55394
        
            For iIndice = 1 To objGrid.iLinhasExistentes
        
                'Não pode somar a Linha atual
                If GridMovimentos.Row <> iIndice Then
        
                    sCodProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)
                    sAlmoxarifado = GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
        
                    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 55395
        
                    'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                    If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) And UCase(GridMovimentos.TextMatrix(iIndice, iGrid_Lote_Col)) = UCase(sLote) Then
        
                        'Verifica se há alguma QuanTidade informada
                        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) <> 0 Then
        
                            sUnidadeProd = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        
                            dQuantidadeProd = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
                            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                            If lErro <> SUCESSO Then gError 55396
        
                            dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
        
                        End If
        
                    End If
        
                End If
        
            Next
        
            dQuantTotal = dQuantTotal + dQuantAtual
    
            If dQuantTotal > CDbl(QuantDisponivel.Caption) Then gError 55397
    
        End If

    End If

    Testa_QuantRequisitada = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada:

    Testa_QuantRequisitada = gErr

    Select Case gErr

        Case 55392, 55393, 55395, 55396

        Case 55394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case 55397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173990)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_REQUISICAO_MATERIAL_CONSUMO_MOVIMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Requisição de Material para Consumo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ReqConsumo"
    
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
        ElseIf Me.ActiveControl Is AlmoxPadrao Then
            Call AlmoxPadraoLabel_Click
        ElseIf Me.ActiveControl Is CclPadrao Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is ContaContabilEst Or Me.ActiveControl Is ContaContabilAplic Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Lote Then 'Inserido por Wagner
            Call BotaoLote_Click
        End If
    
    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        Call Trata_CodigoBarras1
    
    End If
    
End Sub

Private Sub QuantDisponivel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivel, Source, X, Y)
End Sub

Private Sub QuantDisponivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub AlmoxPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxPadraoLabel, Source, X, Y)
End Sub

Private Sub AlmoxPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
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

Private Function Carrega_Requisitante() As Long
'Carrega a combo de Historico
'Inserido por Wagner

Dim lErro As Long

On Error GoTo Erro_Carrega_Requisitante

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_REQUISITANTE, Requisitante)
    If lErro <> SUCESSO Then gError 131870

    Carrega_Requisitante = SUCESSO

    Exit Function

Erro_Carrega_Requisitante:

    Carrega_Requisitante = gErr

    Select Case gErr
    
        Case 131870

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173991)

    End Select

    Exit Function

End Function

'################################################################
'Inserido por Wagner 04/10/2005
Private Sub BotaoLote_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim sLote As String
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_BotaoLote_Click

    If (GridMovimentos.Row = 0) Then gError 140223

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
    sLote = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 140224

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 140225
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) = 0 Then gError 177296
    
    Set objAlmoxarifado = New ClassAlmoxarifado
    
    objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 177297
    
    If Len(Trim(sLote)) > 0 Then
        objRastroLoteSaldo.sLote = sLote
    End If

    colSelecao.Add sProdutoFormatado
    colSelecao.Add objAlmoxarifado.iCodigo

    Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoRastroLote, "Produto = ? AND Almoxarifado = ?")

    Exit Sub

Erro_BotaoLote_Click:

    Select Case gErr

        Case 140224, 177297
        
        Case 140223
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 140225
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 177296
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, GridMovimentos.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165459)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRastroLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoRastroLote_evSelecao

    Set objRastroLoteSaldo = obj1

    If (GridMovimentos.Row > 0) Then
        Lote.Text = objRastroLoteSaldo.sLote
        
        'Carrega as séries na coleção global
        lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
        If lErro <> SUCESSO Then gError 141913
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col) = objRastroLoteSaldo.sLote
    End If

    objProduto.sCodigo = objRastroLoteSaldo.sProduto
            
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 140226
    
    'Se for rastro por série
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 140228
    Else
   
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivelLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 140228
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoRastroLote_evSelecao:

    Select Case gErr
    
        Case 140226 To 140228, 141913
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165460)

    End Select

    Exit Sub

End Sub
'######################################################################

'#####################################################
'Inserido por Wagner 13/03/2006
Public Sub BotaoSerie_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim bPodeAlterarQtd As Boolean

On Error GoTo Erro_BotaoSerie_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 141914
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 141915
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then gError 177303
    If StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) = 0 Then gError 177304
        
    'Formata o produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141916
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141917
       
    objItemMovEstoque.dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    objItemMovEstoque.iItemNF = GridMovimentos.Row
    objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
    objItemMovEstoque.sProduto = sProdutoFormatado
    objItemMovEstoque.sSiglaUM = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
    'Verifica se a Requisição foi Estornada
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) = "1" Then
        objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_CONSUMO
    Else
        objItemMovEstoque.iTipoMov = MOV_EST_CONSUMO
    End If
                
    bPodeAlterarQtd = True
    If colItensNumIntDoc.Count >= GridMovimentos.Row Then
        If colItensNumIntDoc.Item(GridMovimentos.Row) <> 0 Then
            bPodeAlterarQtd = False
        End If
    End If
    
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela_Modal("RastreamentoSerie", gcolcolRastreamentoSerie.Item(GridMovimentos.Row), objItemMovEstoque, Me.Name, bPodeAlterarQtd)
                    
    lErro = Acerta_Quantidade_Rastreada(GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141918
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
    
        Case 141914
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 141915 To 141918
        
        Case 177303
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, GridMovimentos.Row)
        
        Case 177304
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, GridMovimentos.Row)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141919)
    
    End Select
    
    Exit Sub

End Sub

Public Function Carrega_Series(colRastreamentoMovto As Collection, ByVal dQuantidade As Double, ByVal sLoteIni As String, ByVal dQuantidadeAnterior As Double, ByVal sLoteIniAnterior As String, ByVal iLinha As Integer)
'Gera as séries a partir da série inicial e quantidade

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoSerie As ClassRastreamentoLote
Dim objRastreamentoSerieIni As ClassRastreamentoLote
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objAlmoxarifado As ClassAlmoxarifado
Dim vbResult As VbMsgBoxResult
Dim colRastreamentoMovtoAux As New Collection
Dim iTipoMovto As Integer

On Error GoTo Erro_Carrega_Series

    'Formata o produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141921
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141922
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 141923

    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
        'Verifica se a Requisição foi Estornada
        If GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col) = "1" Then
            iTipoMovto = MOV_EST_ESTORNO_CONSUMO
        Else
            iTipoMovto = MOV_EST_CONSUMO
        End If
                
        If dQuantidadeAnterior <> 0 And Len(Trim(sLoteIniAnterior)) <> 0 And iTipoMovtoAnt = iTipoMovto Then
            
            If Abs(dQuantidade - dQuantidadeAnterior) > QTDE_ESTOQUE_DELTA Or sLoteIni <> sLoteIniAnterior Then
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_MODIFICACAO_SERIES")
                If vbResult = vbNo Then gError 141920
            Else
                vbResult = vbNo
            End If
            
        Else
            vbResult = vbYes
            iTipoMovtoAnt = iTipoMovto
        End If
                    
        If vbResult = vbYes Then
                                
            If Len(Trim(sLoteIni)) <> 0 Then
            
                If Not IsNumeric(right(sLoteIni, objProduto.iSerieParteNum)) Then gError 141924
                
                Set objRastreamentoSerieIni = New ClassRastreamentoLote
                
                objRastreamentoSerieIni.sProduto = objProduto.sCodigo
                objRastreamentoSerieIni.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))
                objRastreamentoSerieIni.sCodigo = sLoteIni
                
                lErro = CF("RastreamentoLote_Le", objRastreamentoSerieIni)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 141925
                
                If lErro <> SUCESSO Then gError 141926
            
                Set objItemMovEstoque = New ClassItemMovEstoque
                
                objItemMovEstoque.dQuantidade = Fix(dQuantidade)
                objItemMovEstoque.iItemNF = GridMovimentos.Row
                objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                objItemMovEstoque.sProduto = sProdutoFormatado
                objItemMovEstoque.sSiglaUM = GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                objItemMovEstoque.iTipoMov = iTipoMovto
       
                Set objAlmoxarifado = New ClassAlmoxarifado
                
                objAlmoxarifado.sNomeReduzido = objItemMovEstoque.sAlmoxarifadoNomeRed
        
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 177237
                
                objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                
                lErro = CF("Rastreamento_Serie_Gera", objItemMovEstoque, objProduto, sLoteIni, colRastreamentoMovtoAux)
                If lErro <> SUCESSO Then gError 177240

            End If

            'Remove os dados anteriores
            For iIndice = colRastreamentoMovto.Count To 1 Step -1
                colRastreamentoMovto.Remove iIndice
            Next
            
            'Coloca os novos dados
            For Each objRastreamentoMovto In colRastreamentoMovtoAux
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
            
        End If

    End If

    Carrega_Series = SUCESSO
    
    Exit Function

Erro_Carrega_Series:

    Carrega_Series = gErr

    Select Case gErr
    
        Case 141921, 141922, 141925, 141927, 141920, 141929, 177240
        
        Case 141923
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 141924
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, right(sLoteIni, objProduto.iSerieParteNum))
        
        Case 141926
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerieIni.sProduto, objRastreamentoSerieIni.sCodigo, objRastreamentoSerieIni.iFilialOP)
        
        Case 141928
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerie.sProduto, objRastreamentoSerie.sCodigo, objRastreamentoSerie.iFilialOP)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141930)

    End Select

    Exit Function

End Function

Public Function Acerta_Quantidade_Rastreada(ByVal iLinha As Integer)
'Acerta a quantidade do grid com base na quantidadse da coleção global de movimentos de séries

Dim lErro As Long
Dim dQuantidade As Double
Dim objRastreamentoSerie As ClassRastreamentoMovto

On Error GoTo Erro_Acerta_Quantidade_Rastreada

    For Each objRastreamentoSerie In gcolcolRastreamentoSerie.Item(iLinha)
    
        dQuantidade = dQuantidade + objRastreamentoSerie.dQuantidade
    
    Next
                
    If colItensNumIntDoc.Item(iLinha) = 0 Then

        If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col) <> "1" Then

            lErro = Testa_QuantRequisitada(dQuantidade)
            If lErro <> SUCESSO Then gError 141932

        End If

    End If
        
    GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    Acerta_Quantidade_Rastreada = SUCESSO
    
    Exit Function

Erro_Acerta_Quantidade_Rastreada:

    Acerta_Quantidade_Rastreada = gErr

    Select Case gErr
    
        Case 141932

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141931)

    End Select

    Exit Function

End Function
'#####################################################

Private Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Private Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = CStr(giFilialEmpresa) & "," & Codigo.Text
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158093)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158094)
     
    End Select
     
    Exit Sub

End Sub

Public Function Trata_CodigoBarras1() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoEnxuto As String
Dim sCodBarras As String
Dim sCodBarrasOriginal As String
Dim dCusto As Double

On Error GoTo Erro_Trata_CodigoBarras1

    If objGrid.iLinhasExistentes + 1 = GridMovimentos.Row Then
    
        'Verifica se o Produto está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then
            
            If Me.ActiveControl Is Produto Then
                    
                    Set objGrid.objControle = Produto
            
                    lErro = Grid_Abandona_Celula(objGrid)
                    If lErro <> SUCESSO Then gError 210829
                    
            End If
            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 210830
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 210831
    
                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 210832
    
                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 210833
        
                Me.Show
        
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
                
                gError 210867
'
'                If Not Me.ActiveControl Is Produto Then
'                    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
'
'                    'Preenche a Linha do Grid
'                    lErro = ProdutoLinha_Preenche(objProduto)
'                    If lErro <> SUCESSO Then gError 210834
'
'                End If
    
            Else
            
                gError 210835
    
    
            End If
            
            GridMovimentos.SetFocus
            GridMovimentos.FocusRect = flexFocusHeavy
    
        End If
    
    End If

    Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_Trata_CodigoBarras1:

    Trata_CodigoBarras1 = gErr

'    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 210829 To 210831, 210834, 210835, 210867

        Case 210832
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 210833
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210836)

    End Select

    Exit Function

End Function


