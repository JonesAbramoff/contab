VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Transfer 
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4455
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   825
      Width           =   9255
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
         Height          =   360
         Left            =   75
         TabIndex        =   84
         Top             =   4035
         Width           =   1665
      End
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
         Height          =   360
         Left            =   1935
         TabIndex        =   83
         Top             =   4035
         Width           =   1665
      End
      Begin VB.ComboBox FilialOP 
         Height          =   315
         Left            =   6630
         TabIndex        =   76
         Top             =   1710
         Width           =   2160
      End
      Begin MSMask.MaskEdBox Lote 
         Height          =   270
         Left            =   7320
         TabIndex        =   75
         Top             =   1380
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
         Left            =   1845
         Picture         =   "Transfer.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   375
         Width           =   300
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1350
         Width           =   690
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2385
         Width           =   2600
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
         Height          =   360
         Left            =   3795
         TabIndex        =   19
         Top             =   4035
         Width           =   1665
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
         Height          =   360
         Left            =   5655
         TabIndex        =   20
         Top             =   4035
         Width           =   1665
      End
      Begin VB.Frame Frame2 
         Caption         =   "Almoxarifados Padrão"
         Height          =   1095
         Left            =   5925
         TabIndex        =   43
         Top             =   0
         Width           =   3240
         Begin MSMask.MaskEdBox AlmoxPadraoOrigem 
            Height          =   315
            Left            =   975
            TabIndex        =   5
            Top             =   255
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxPadraoDestino 
            Height          =   315
            Left            =   975
            TabIndex        =   6
            Top             =   660
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label AlmoxOrigemLabel 
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
            Height          =   195
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   300
            Width           =   660
         End
         Begin VB.Label AlmoxDestinoLabel 
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
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   49
            Top             =   720
            Width           =   720
         End
      End
      Begin VB.ComboBox TipoOrigem 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Transfer.ctx":00EA
         Left            =   5250
         List            =   "Transfer.ctx":00EC
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1350
         Width           =   2010
      End
      Begin VB.ComboBox TipoDestino 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Transfer.ctx":00EE
         Left            =   4590
         List            =   "Transfer.ctx":00F0
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1710
         Width           =   2010
      End
      Begin VB.CheckBox Estorno 
         Height          =   210
         Left            =   7440
         TabIndex        =   16
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
         Height          =   360
         Left            =   7515
         TabIndex        =   21
         Top             =   4035
         Width           =   1665
      End
      Begin MSMask.MaskEdBox ContaContabilEstSaida 
         Height          =   225
         Left            =   5880
         TabIndex        =   17
         Top             =   2715
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   397
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
      Begin MSMask.MaskEdBox ContaContabilEstEntrada 
         Height          =   225
         Left            =   3780
         TabIndex        =   13
         Top             =   2970
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   397
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
      Begin MSMask.MaskEdBox AlmoxDestino 
         Height          =   225
         Left            =   3240
         TabIndex        =   11
         Top             =   1770
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   2970
         TabIndex        =   9
         Top             =   1380
         Width           =   990
         _ExtentX        =   1746
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   375
         TabIndex        =   7
         Top             =   1425
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AlmoxOrigem 
         Height          =   225
         Left            =   3990
         TabIndex        =   10
         Top             =   1365
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   315
         Left            =   3840
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   2775
         TabIndex        =   3
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
         Height          =   2055
         Left            =   30
         TabIndex        =   18
         Top             =   1170
         Width           =   9210
         _ExtentX        =   16245
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   11
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   870
         TabIndex        =   1
         Top             =   360
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Hora 
         Height          =   300
         Left            =   4815
         TabIndex        =   4
         Top             =   360
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
         Left            =   4290
         TabIndex        =   77
         Top             =   405
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Movimentos de Estoque"
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
         TabIndex        =   50
         Top             =   900
         Width           =   2040
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
         Height          =   180
         Left            =   2220
         TabIndex        =   51
         Top             =   405
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
         Height          =   210
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   52
         Top             =   405
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade na Origem:"
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
         Left            =   5595
         TabIndex        =   53
         Top             =   3705
         Width           =   1965
      End
      Begin VB.Label QuantOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7620
         TabIndex        =   54
         Top             =   3645
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   2
      Left            =   75
      TabIndex        =   22
      Top             =   825
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   3960
         TabIndex        =   85
         Tag             =   "1"
         Top             =   1440
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
         Left            =   6420
         TabIndex        =   28
         Top             =   375
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
         Left            =   6420
         TabIndex        =   26
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   900
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
         Left            =   7875
         TabIndex        =   27
         Top             =   60
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4920
         TabIndex        =   36
         Top             =   1560
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
         TabIndex        =   38
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   37
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2595
         Left            =   6360
         TabIndex        =   40
         Top             =   1500
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   44
         Top             =   3345
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
            TabIndex        =   55
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
            TabIndex        =   57
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   58
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
         Left            =   3450
         TabIndex        =   31
         Top             =   900
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   32
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
         Left            =   2325
         TabIndex        =   34
         Top             =   1845
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
         Left            =   3495
         TabIndex        =   35
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   33
         Top             =   1905
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
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   495
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   25
         Top             =   510
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
         TabIndex        =   24
         Top             =   120
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
         TabIndex        =   23
         Top             =   105
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
         Left            =   15
         TabIndex        =   39
         Top             =   1155
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
         Height          =   2790
         Left            =   6360
         TabIndex        =   41
         Top             =   1500
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4921
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2790
         Left            =   6360
         TabIndex        =   42
         Top             =   1500
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4921
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
         Left            =   6480
         TabIndex        =   29
         Top             =   690
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
         TabIndex        =   59
         Top             =   150
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   60
         Top             =   105
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
         TabIndex        =   61
         Top             =   585
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   62
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   63
         Top             =   540
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
         TabIndex        =   64
         Top             =   570
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
         TabIndex        =   65
         Top             =   930
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   68
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
         TabIndex        =   69
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   70
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   71
         Top             =   3015
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
         TabIndex        =   72
         Top             =   540
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
         TabIndex        =   73
         Top             =   150
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
         TabIndex        =   74
         Top             =   150
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   2100
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   105
      Width           =   2160
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Transfer.ctx":00F2
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Transfer.ctx":027C
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Transfer.ctx":03FA
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Transfer.ctx":092C
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4860
      Left            =   30
      TabIndex        =   47
      Top             =   465
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   8573
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
Attribute VB_Name = "Transfer"
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
Private Const ESTORNO1 As String = "Estorno"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const ALMOX_ORIGEM As String = "Almox_Origem"
Private Const ALMOX_DESTINO As String = "Almox_Destino"
Private Const TIPO_ORIGEM As String = "Tipo_Origem"
Private Const TIPO_DESTINO As String = "Tipo_Destino"
Private Const CONTACONTABILESTENTRADA1 As String = "CtaAlmoxDestino"
Private Const CONTACONTABILESTSAIDA1 As String = "CtaAlmoxOrigem"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
Private Const QUANT_DISPONIVEL As String = "Quant_Disponivel"
Private Const QUANT_CONSIGNADA1 As String = "Quant_Consignada"
Private Const QUANT_CONSIGNADADETERC1 As String = "Quant_ConsigDeTerc"

Dim gcolcolRastreamentoSerie As Collection 'Inserido por Wagner
Dim iTipoMovtoAnt As Integer

Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim lCodigoAntigo As Long

Dim colItensNumIntDoc As Collection

Dim objGrid As AdmGrid
Dim iLinhaAntiga As Integer
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_AlmoxOrigem_Col As Integer
Public iGrid_TipoOrigem_Col As Integer
Dim iGrid_AlmoxDestino_Col As Integer
Public iGrid_TipoDestino_Col As Integer
Dim iGrid_Estorno_Col As Integer
Dim iGrid_ContaContabilEstEntrada_Col As Integer
Dim iGrid_ContaContabilEstSaida_Col As Integer
Dim iGrid_Lote_Col As Integer
Dim iGrid_FilialOP_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxOrigem As AdmEvento
Attribute objEventoAlmoxOrigem.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxDestino As AdmEvento
Attribute objEventoAlmoxDestino.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Movimentos = 1
Private Const TAB_Contabilizacao = 2

Private Sub BotaoExcluir_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 89923

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    'Exclui a transferência
    lErro = CF("MovimentoEstoque_Trata_Exclusao", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 89924

    Call Limpa_Tela_Transfer

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 89923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 89924
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175456)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("MovEstoque_Automatico", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then gError 57527

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True
    
    lCodigoAntigo = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 57527
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175457)
    
    End Select

    Exit Sub

End Sub

Private Sub AlmoxDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AlmoxDestinoLabel_Click()

Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colSelecao As New Collection

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxDestino)

End Sub

Private Sub AlmoxOrigem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AlmoxOrigemLabel_Click()

Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colSelecao As New Collection

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxOrigem)

End Sub

Private Sub AlmoxPadraoDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AlmoxPadraoDestino_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadraoDestino_Validate

    If Len(Trim(AlmoxPadraoDestino.Text)) <> 0 Then

        lErro = TP_Almoxarifado_Filial_Le(AlmoxPadraoDestino, objAlmoxarifado, 0)
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 30800
    
        If lErro = 25136 Then gError 22912
        
        If lErro = 25143 Then gError 22913

    End If

    Exit Sub

Erro_AlmoxPadraoDestino_Validate:

    Cancel = True


    Select Case gErr

        Case 30800
        
        Case 22912, 22913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadraoDestino.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175458)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxPadraoOrigem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AlmoxPadraoOrigem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadraoOrigem_Validate

    If Len(Trim(AlmoxPadraoOrigem.Text)) <> 0 Then

        lErro = TP_Almoxarifado_Filial_Le(AlmoxPadraoOrigem, objAlmoxarifado, 0)
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 30890
    
        If lErro = 25136 Then gError 22914
        
        If lErro = 25143 Then gError 22915

    End If

    Exit Sub

Erro_AlmoxPadraoOrigem_Validate:

    Cancel = True


    Select Case gErr

        Case 30890
        
        Case 22914, 22915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadraoOrigem.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175459)

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

    If GridMovimentos.Row = 0 Then gError 43770
    
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = "" Then gError 43771

    sConta = String(STRING_CONTA, 0)
    
    'Verifica através da coluna que está preenchida
    If GridMovimentos.Col = iGrid_ContaContabilEstEntrada_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEstEntrada.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 43772
    
    ElseIf GridMovimentos.Col = iGrid_ContaContabilEstSaida_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEstSaida.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 43773
    
    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 43770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 43771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 43772, 43773

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175460)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub ContaContabilEstEntrada_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabilEstEntrada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ContaContabilEstEntrada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ContaContabilEstEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabilEstEntrada
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ContaContabilEstSaida_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabilEstSaida_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ContaContabilEstSaida_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ContaContabilEstSaida_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabilEstSaida
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ContaContabilEstEntrada(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilEstEntrada
    
    Set objGrid.objControle = ContaContabilEstEntrada
    
    If Len(Trim(ContaContabilEstEntrada.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilEstEntrada.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49632
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49633
            
            ContaContabilEstEntrada.PromptInclude = False
            ContaContabilEstEntrada.Text = sContaMascarada
            ContaContabilEstEntrada.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEstEntrada.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49634
    
            'conta não cadastrada
            If lErro = 5700 Then gError 49635
             
        End If
    
    Else
        
        ContaContabilEstEntrada.PromptInclude = False
        ContaContabilEstEntrada.Text = ""
        ContaContabilEstEntrada.PromptInclude = True
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49647
    
    Saida_Celula_ContaContabilEstEntrada = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEstEntrada:

    Saida_Celula_ContaContabilEstEntrada = gErr

    Select Case gErr

        Case 49632, 49634, 49647
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 49635
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEstEntrada.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175461)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilEstSaida(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilEstSaida

    Set objGrid.objControle = ContaContabilEstSaida

    If Len(Trim(ContaContabilEstSaida.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilEstSaida.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49628
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49629
            
            ContaContabilEstSaida.PromptInclude = False
            ContaContabilEstSaida.Text = sContaMascarada
            ContaContabilEstSaida.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEstSaida.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49630
    
            'conta não cadastrada
            If lErro = 5700 Then gError 49631
             
        End If
    
    Else
        
        ContaContabilEstSaida.PromptInclude = False
        ContaContabilEstSaida.Text = ""
        ContaContabilEstSaida.PromptInclude = True
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49643
    
   Saida_Celula_ContaContabilEstSaida = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEstSaida:

    Saida_Celula_ContaContabilEstSaida = gErr

    Select Case gErr

        Case 49628, 49630, 49643
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 49631
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEstSaida.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175462)

    End Select

    Exit Function

End Function

Private Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridMovimentos.Row = 0 Then gError 43783

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 43784
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30789

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado

        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)

    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 30789

        Case 43783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175463)

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
    If lErro <> SUCESSO Then gError 30859
    
    Call Limpa_Tela_Transfer

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 30859

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175464)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 30878

    Call Limpa_Tela_Transfer

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 30878

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175465)

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

    If GridMovimentos.Row = 0 Then gError 43782

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
        
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 22924
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProduto
    
    End If
    sSelecao = "ControleEstoque<>?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    
    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_BotaoProdutos_Click:

     Select Case gErr
     
        Case 22924
     
        Case 43782
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr, Error)
     
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175466)
     
     End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Limpa_colItensNumIntDoc(colItensNumIntDoc As Collection)

Dim lErro As Long
Dim iIndice As Integer
Dim iCount As Integer

On Error GoTo Erro_Limpa_colItensNumIntDoc

    iCount = colItensNumIntDoc.Count
    Set colItensNumIntDoc = New Collection

    For iIndice = 1 To iCount / 2
        
        GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "0"
        colItensNumIntDoc.Add 0
        colItensNumIntDoc.Add 0
        
    Next
    
    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 30782
    
    Exit Sub

Erro_Limpa_colItensNumIntDoc:

    Select Case gErr
    
        Case 30782
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175467)
    
    End Select

    Exit Sub
    
End Sub

Private Function Verifica_TipoTelaAtual(lCodigo As Long, iTipoMov As Integer) As Long
'verifica se o tipo do movimento é transferencia

Dim iResultado As Integer
Dim objTipoMovEst  As New ClassTipoMovEst
Dim lErro As Long

On Error GoTo Erro_Verifica_TipoTelaAtual

    objTipoMovEst.iCodigo = iTipoMov

    lErro = CF("TipoMovEstoque_Le", objTipoMovEst)
    If lErro <> SUCESSO And lErro <> 30372 Then gError 30795

    If lErro = 30372 Then gError 30796

    If objTipoMovEst.iTransferencia = 0 Then gError 30797

    Verifica_TipoTelaAtual = SUCESSO

    Exit Function
    
Erro_Verifica_TipoTelaAtual:

    Verifica_TipoTelaAtual = gErr
    
    Select Case gErr
    
        Case 30795

        Case 30796
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMOVEST_NAO_CADASTRADO", gErr, objTipoMovEst.iCodigo)

        Case 30797
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_TRANSFERENCIA", gErr, lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175468)

    End Select
    
    iAlterado = 0

    Exit Function
    
    

End Function

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    'se o codigo foi trocado
    If lCodigoAntigo <> StrParaLong(Trim(Codigo.Text)) Then

        If Len(Trim(Codigo.ClipText)) > 0 Then
        
            Call Limpa_colItensNumIntDoc(colItensNumIntDoc)
            
            objMovEstoque.lCodigo = Codigo.Text
            objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
            'Lê MovEstoque no Banco de Dados
            lErro = CF("MovEstoque_Le", objMovEstoque)
            If lErro <> SUCESSO And lErro <> 30128 Then gError 34914
        
            'Le o Movimento de Estoque e Verifica se ele já foi estornado
            lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque)
            If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 34914
            
            'Se todos os Itens do Movimento foram estornados
            If lErro = 78885 Then gError 78891
            
            If lErro = SUCESSO Then
                
                lErro = Verifica_TipoTelaAtual(objMovEstoque.lCodigo, objMovEstoque.iTipoMov)
                If lErro <> SUCESSO Then gError 34902
    
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")
                
                If vbMsg = vbNo Then gError 34915
                
                lErro = Preenche_Tela(objMovEstoque)
                If lErro <> SUCESSO Then gError 34916
                
            End If

        End If
        
        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
            
        Case 34902
            lCodigoAntigo = 0
        
        Case 34914, 34916
        
        Case 34915
        
        Case 78891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175469)
    
    End Select
    
    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim colSelecao As New Collection
Dim objMovEstoque As New ClassMovEstoque

    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    Call Chama_Tela("MovEstoqueTransferenciaLista", colSelecao, objMovEstoque, objEventoCodigo)

End Sub
Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub CargaCombo_TipoOrigem(objTipoOrigem As Object)
    
    objTipoOrigem.AddItem TRANSF_DISPONIVEL_STRING
    objTipoOrigem.ItemData(objTipoOrigem.NewIndex) = MOV_EST_SAIDA_TRANSF_DISP
    
    If giTipoVersao = VERSAO_FULL Then
        objTipoOrigem.AddItem TRANSF_DEFEITUOSO_STRING
        objTipoOrigem.ItemData(objTipoOrigem.NewIndex) = MOV_EST_SAIDA_TRANSF_DEFEIT
        objTipoOrigem.AddItem TRANSF_INDISPONIVEL_STRING
        objTipoOrigem.ItemData(objTipoOrigem.NewIndex) = MOV_EST_SAIDA_TRANSF_INDISP
    End If
    
    objTipoOrigem.AddItem STRING_TRANSF_CONSIG3
    objTipoOrigem.ItemData(objTipoOrigem.NewIndex) = MOV_EST_SAIDA_TRANSF_CONSIG_TERC
    objTipoOrigem.AddItem STRING_TRANSF_CONSIG
    objTipoOrigem.ItemData(objTipoOrigem.NewIndex) = MOV_EST_SAIDA_TRANSF_CONSIG_NOSSO
    
    objTipoOrigem.AddItem TRANSF_OUTRAS_TERC_STRING
    objTipoOrigem.ItemData(objTipoOrigem.NewIndex) = MOV_EST_SAIDA_TRANSF_OUTRAS_TERC
    
    'incluida linha em branco p/poder tratar caso em que nao tem qtde suficiente.
    objTipoOrigem.AddItem ""
    objTipoOrigem.ItemData(objTipoOrigem.NewIndex) = 0
    
End Sub

Private Sub CargaCombo_TipoDestino(objTipoDestino As Object)

    objTipoDestino.AddItem TRANSF_DISPONIVEL_STRING
    objTipoDestino.ItemData(objTipoDestino.NewIndex) = MOV_EST_ENTRADA_TRANSF_DISP
    
    If giTipoVersao = VERSAO_FULL Then
    
        objTipoDestino.AddItem TRANSF_DEFEITUOSO_STRING
        objTipoDestino.ItemData(objTipoDestino.NewIndex) = MOV_EST_ENTRADA_TRANSF_DEFEIT
        objTipoDestino.AddItem TRANSF_INDISPONIVEL_STRING
        objTipoDestino.ItemData(objTipoDestino.NewIndex) = MOV_EST_ENTRADA_TRANSF_INDISP
        
    End If
   
    objTipoDestino.AddItem TRANSF_OUTRAS_TERC_STRING
    objTipoDestino.ItemData(objTipoDestino.NewIndex) = MOV_EST_ENTRADA_TRANSF_OUTRAS_TERC
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner
    
    Set colItensNumIntDoc = New Collection
    
    Set objEventoCodigo = New AdmEvento
    Set objEventoAlmoxOrigem = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoAlmoxDestino = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoRastroLote = New AdmEvento 'Inserido por Wagner
    
    'Preenche a List de Combo Boxes Tipo Origem e Tipo Destino
    'Preenche o ItemData com o correspondente

    Call CargaCombo_TipoOrigem(TipoOrigem)
    Call CargaCombo_TipoDestino(TipoDestino)
    
    'Preenche a Data
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Inicializa mascara de contaContabilEstEntrada e ContaContabilEstSaida
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEstEntrada)
    If lErro <> SUCESSO Then gError 49666
    
    'Inicializa mascara de contaContabilEstEntrada e ContaContabilEstSaida
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEstSaida)
    If lErro <> SUCESSO Then gError 49667

    'Inicializa a Máscara do Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 30777

    'Formato para a Quantidade
    Quantidade.Format = FORMATO_ESTOQUE
    
    QuantOrigem.Caption = Formata_Estoque(0)
    
    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 78295
    
    'Inicializa o GridMovimentos
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridMovimentos(objGrid)
    If lErro <> SUCESSO Then gError 30778
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE)
    If lErro <> SUCESSO Then gError 39631
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 30777, 30778, 39631, 49666, 49667, 78295

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175470)

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
    If lErro <> SUCESSO Then gError 78292

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

        Case 78292
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175471)

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
    objGridInt.colColuna.Add ("Almox. Origem")
    objGridInt.colColuna.Add ("Tipo Origem")
    objGridInt.colColuna.Add ("Almox. Destino")
    objGridInt.colColuna.Add ("Tipo Destino")
    objGridInt.colColuna.Add ("Lote / O.P./Serie Ini.")
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Conta Contabil Origem")
    objGridInt.colColuna.Add ("Conta Contabil Destino")
    objGridInt.colColuna.Add ("Estorno")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (AlmoxOrigem.Name)
    objGridInt.colCampo.Add (TipoOrigem.Name)
    objGridInt.colCampo.Add (AlmoxDestino.Name)
    objGridInt.colCampo.Add (TipoDestino.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add ("ContaContabilEstSaida")
    objGridInt.colCampo.Add ("ContaContabilEstEntrada")
    objGridInt.colCampo.Add (Estorno.Name)
    
    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_AlmoxOrigem_Col = 5
    iGrid_TipoOrigem_Col = 6
    iGrid_AlmoxDestino_Col = 7
    iGrid_TipoDestino_Col = 8
    iGrid_Lote_Col = 9
    iGrid_FilialOP_Col = 10
    iGrid_ContaContabilEstSaida_Col = 11
    iGrid_ContaContabilEstEntrada_Col = 12
    iGrid_Estorno_Col = 13
        
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

'Extrai os campos da tela que correspondem aos campos no Banco de Dados
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MovEstoqueTransferencia"

    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    If Len(Trim(Data.ClipText)) <> 0 Then
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
    'colSelecao.Add "NumIntDocEst", OP_IGUAL, 0

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175472)

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
    If lErro <> SUCESSO Then gError 30779

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 30779

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175473)

    End Select

    Exit Sub

End Sub

Function Preenche_Tela(objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim iIndice As Integer

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
    
    'Limpa o Label QuantOrigem
    QuantOrigem.Caption = ""
    
    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection
    Set objMovEstoque.colItens = New ColItensMovEstoque
    
    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovEstoqueItens_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30116 Then gError 30780

    'Coloca os Dados na Tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objMovEstoque.lCodigo)
    Codigo.PromptInclude = True

    Call DateParaMasked(Data, objMovEstoque.dtData)

'hora
    Hora.PromptInclude = False
    'este teste está correto
    If objMovEstoque.dtData <> DATA_NULA Then Hora.Text = Format(objMovEstoque.dtHora, "hh:mm:ss")
    Hora.PromptInclude = True

    'Passa as Informações de NumIntDoc de colItens para colItensNumIntDoc
    For iIndice = 1 To objMovEstoque.colItens.Count
    
        colItensNumIntDoc.Add objMovEstoque.colItens(iIndice).lNumIntDoc
        
    Next

    lErro = Preenche_GridMovimentos(objMovEstoque.colItens)
    If lErro <> SUCESSO Then gError 30781
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovEstoque.colItens(1).lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39632

    iAlterado = 0
    lCodigoAntigo = objMovEstoque.lCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 30780, 30781, 39632

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175474)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMovimentos(colItens As ColItensMovEstoque) As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iIndice3 As Integer
Dim lErro As Long
Dim iLinha As Integer, iTipoMov As Integer
Dim iCodigoMovEst As Integer
Dim iCont As Integer
Dim sProdutoMascarado As String
Dim objTipoMovEst As ClassTipoMovEst
Dim sContaMascaradaEst As String
Dim colRatreamentoMovto As New Collection
Dim objRatreamentoMovto As New ClassRastreamentoMovto
Dim objFilialOP As New AdmFiliais
Dim colRastreamentoSerie As Collection 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_Preenche_GridMovimentos

    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 15/03/2006

    'Preenche GridMovimentos
    For iIndice = 1 To colItens.Count Step 2

        iLinha = iLinha + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_MascararProduto(colItens(iIndice).sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 30885

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado
        GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col) = colItens(iIndice).sProdutoDesc
        GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = colItens(iIndice).sSiglaUM
        GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(colItens(iIndice).dQuantidade)
        GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col) = "0"

        'Cada par de Ítens indica uma Transferência
        'O par iIndice e iIndice+1 são referentes a mesma Transferência -->
        '--> Isto é garantido na Leitura dos Ítens(MovEstoqueItens_Le)-->
        '--> que está ordenada pelo NumIntDoc que identifica o Item
        For iIndice3 = iIndice To iIndice + 1

            iTipoMov = colItens(iIndice3).iTipoMov
                        
            Select Case iTipoMov

                Case MOV_EST_ENTRADA_TRANSF_DISP_CONSIG3
                    iCodigoMovEst = MOV_EST_ENTRADA_TRANSF_DISP_CONSIG3

                Case MOV_EST_ENTRADA_TRANSF_DISP_CONSIG
                    iCodigoMovEst = MOV_EST_ENTRADA_TRANSF_DISP_CONSIG

                Case Else
                    iCodigoMovEst = MOV_EST_ENTRADA_TRANSF_DISP

            End Select

            'Guarda no ItemData de Disponível o Movimento estoque correspondente ao tipo da origem
            For iCont = 0 To TipoDestino.ListCount - 1
                If TipoDestino.List(iCont) = TRANSF_DISPONIVEL_STRING Then
                    TipoDestino.ItemData(iCont) = iCodigoMovEst
                End If
            Next
            
            If iTipoMov = MOV_EST_SAIDA_TRANSF_DEFEIT Or iTipoMov = MOV_EST_SAIDA_TRANSF_DISP Or iTipoMov = MOV_EST_SAIDA_TRANSF_INDISP Or iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_TERC Or iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_NOSSO Then
                For iIndice2 = 0 To TipoOrigem.ListCount - 1
                    If TipoOrigem.ItemData(iIndice2) = iTipoMov Then
                        
                        If Trim(colItens(iIndice3).sContaContabilEst) <> "" Then
                        
                            lErro = Mascara_MascararConta(colItens(iIndice3).sContaContabilEst, sContaMascaradaEst)
                            If lErro <> SUCESSO Then gError 49668
                             
                             GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstSaida_Col) = sContaMascaradaEst
                        
                        End If
                        
                        GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col) = colItens(iIndice3).sAlmoxarifadoNomeRed
                        GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TipoOrigem.List(iIndice2)
                        Exit For
                    End If
                Next
            End If

            If iTipoMov = MOV_EST_ENTRADA_TRANSF_DEFEIT Or iTipoMov = MOV_EST_ENTRADA_TRANSF_DISP Or iTipoMov = MOV_EST_ENTRADA_TRANSF_INDISP Or iTipoMov = MOV_EST_ENTRADA_TRANSF_CONSIG_TERC Or iTipoMov = MOV_EST_ENTRADA_TRANSF_CONSIG_NOSSO Or iTipoMov = MOV_EST_ENTRADA_TRANSF_DISP_CONSIG Or iTipoMov = MOV_EST_ENTRADA_TRANSF_DISP_CONSIG3 Then
                For iIndice2 = 0 To TipoDestino.ListCount - 1
                    If TipoDestino.ItemData(iIndice2) = iTipoMov Then
                    
                        If Trim(colItens(iIndice3).sContaContabilEst) <> "" Then
                                                
                            lErro = Mascara_MascararConta(colItens(iIndice3).sContaContabilEst, sContaMascaradaEst)
                            If lErro <> SUCESSO Then gError 49671
                            
                            GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstEntrada_Col) = sContaMascaradaEst
                        
                        End If
                        
                        GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxDestino_Col) = colItens(iIndice3).sAlmoxarifadoNomeRed
                        GridMovimentos.TextMatrix(iLinha, iGrid_TipoDestino_Col) = TipoDestino.List(iIndice2)
                        Exit For
                    End If
                Next
            End If
            
            Set colRatreamentoMovto = New Collection
            
            'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
            lErro = CF("RastreamentoMovto_Le_DocOrigem", colItens(iIndice3).lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRatreamentoMovto)
            If lErro <> SUCESSO And lErro <> 78414 Then gError 78612
            
            'Se existe rastreamento
            If colRatreamentoMovto.Count > 0 Then
                            
                'Seta o primeiro Lote
                Set objRatreamentoMovto = colRatreamentoMovto(1)
    
                gcolcolRastreamentoSerie.Add objRatreamentoMovto.colRastreamentoSerie 'Inserido por Wagner 15/03/2006
    
                If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col) = objRatreamentoMovto.sLote
                
                If objRatreamentoMovto.iFilialOP > 0 Then
                
                    objFilialOP.iCodFilial = objRatreamentoMovto.iFilialOP
    
                    'Le a Filial Empresa da OP para pegar a descrição
                    lErro = CF("FilialEmpresa_Le", objFilialOP)
                    If lErro <> SUCESSO Then gError 78830
    
                    GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
                
                End If
                
            '#####################################################
            'Inserido por Wagner 15/03/2006
            Else
                Set colRastreamentoSerie = New Collection
                gcolcolRastreamentoSerie.Add colRastreamentoSerie
            '#####################################################
                
            End If
        
        Next
    
        
    Next
    
    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 30782

    objGrid.iLinhasExistentes = colItens.Count / 2

    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 30782, 49668, 49671, 78612, 78830

        Case 30885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, colItens(iIndice).sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175475)

    End Select

    Exit Function

End Function


Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Estorno_Click()

Dim lErro As Long

On Error GoTo Erro_Estorno_Click

    iAlterado = REGISTRO_ALTERADO

    '#################################################################
    'Inserido por Wagner 13/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 177299
    '#################################################################

    Exit Sub
    
Erro_Estorno_Click:

    Select Case gErr
    
        Case 177299
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177300)
    
    End Select
    
    Exit Sub

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
        
        'Se o Lote não está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantOrigem_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
            If lErro <> SUCESSO Then gError 30151
        Else
            lErro = QuantOrigemLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78814
        End If
            
    End If

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case 30151, 78814

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175476)

    End Select

    Exit Sub

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

Private Sub objEventoAlmoxDestino_evSelecao(obj1 As Object)

Dim objAlmoxarifado As ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    AlmoxPadraoDestino.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

End Sub

Private Sub objEventoAlmoxOrigem_evSelecao(obj1 As Object)

Dim objAlmoxarifado As ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    AlmoxPadraoOrigem.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovEstoque As ClassMovEstoque

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objMovEstoque = obj1

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 30783

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 30783

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175477)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If GridMovimentos.Col <> iGrid_ContaContabilEstEntrada_Col And GridMovimentos.Col <> iGrid_ContaContabilEstSaida_Col Then
        Me.Show
        Exit Sub
    End If
        
    If objPlanoConta.sConta <> "" Then
   
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 43774
        
        If GridMovimentos.Col = iGrid_ContaContabilEstSaida_Col Then
            ContaContabilEstSaida.PromptInclude = False
            ContaContabilEstSaida.Text = sContaEnxuta
            ContaContabilEstSaida.PromptInclude = True
        Else
            ContaContabilEstEntrada.PromptInclude = False
            ContaContabilEstEntrada.Text = sContaEnxuta
            ContaContabilEstEntrada.PromptInclude = True
        End If
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = objGrid.objControle.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 43774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175478)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    If GridMovimentos.Row <> 0 Then

        'Verifica se a Coluna corrente é AlmoxOrigem ou AlmoxDestino
        If GridMovimentos.Col = iGrid_AlmoxOrigem_Col Or GridMovimentos.Col = iGrid_AlmoxDestino_Col Then

            'Verifica se o Produto da Linha corrente foi Preenchido
            lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 30790

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

                objAlmoxarifado.iCodigo = objEstoqueProduto.iAlmoxarifado

                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 30791

                If lErro = 25056 Then gError 30792

                GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = objAlmoxarifado.sNomeReduzido

                If GridMovimentos.Col = iGrid_AlmoxOrigem_Col Then
                    AlmoxOrigem.Text = objAlmoxarifado.sNomeReduzido
                    
                    'Se o Lote não está preenchido
                    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
                        lErro = QuantOrigem_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), objAlmoxarifado.sNomeReduzido, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
                        If lErro <> SUCESSO Then gError 30793
                    Else
                        lErro = QuantOrigemLote_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), objAlmoxarifado.sNomeReduzido, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
                        If lErro <> SUCESSO Then gError 78815
                    End If
                    
                Else
                    AlmoxDestino.Text = objAlmoxarifado.sNomeReduzido
                End If
            End If
            
            'preenche as contacontabeis com o padrao
            lErro = Preenche_ContaContabil_Tela()
            If lErro <> SUCESSO Then gError 52121
            
        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 30790, 30791, 30793, 78815

        Case 30792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case 52121
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175479)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30784

    'Verifica se a Coluna do Produto é a corrente e se o Produto não está preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 30785

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 30786

        If lErro = 28030 Then gError 30787

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        If Not (Me.ActiveControl Is Produto) Then
        
            'Preenche a célula de Produto
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = sProdutoMascarado
    
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 30788
            
            lErro = Preenche_ContaContabil_Tela()
            If lErro <> SUCESSO Then gError 52226

            'Se o Lote não está preenchido
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
                lErro = QuantOrigem_Calcula1(sProdutoMascarado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), objProduto)
                If lErro <> SUCESSO Then gError 34917
            Else
                lErro = QuantOrigemLote_Calcula1(sProdutoMascarado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)), objProduto)
                If lErro <> SUCESSO Then gError 78817
            End If
                
        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 30784, 30786, 30788, 34917, 52226, 78817

        Case 30785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case 30787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175480)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objTipoMovEst As New ClassTipoMovEst

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objMovEstoque Is Nothing Then

        objMovEstoque.iFilialEmpresa = giFilialEmpresa

        'Lê MovEstoque no Banco de Dados
        lErro = CF("MovEstoque_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30128 Then gError 30794

        If lErro <> 30128 Then 'Se ele existe

            lErro = Verifica_TipoTelaAtual(objMovEstoque.lCodigo, objMovEstoque.iTipoMov)
            If lErro <> SUCESSO Then gError 55424

            lErro = Preenche_Tela(objMovEstoque)
            If lErro <> SUCESSO Then gError 30798

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

        Case 30794, 30798, 55424

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175481)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 30801

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 30801

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175482)

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
    If lErro <> SUCESSO Then gError 89808

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 89808

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175483)

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
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then
        
            'trata os casos de estorno da versao light
            Call MovEstoque_Trata_Estorno_Versao_Light
        
            Call objContabil.Contabil_Carga_Modelo_Padrao
            
        End If
    
        Select Case iFrameAtual
        
            Case TAB_Movimentos
                Parent.HelpContextID = IDH_TRANSFERENCIA_MOVIMENTO
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_TRANSFERENCIA_CONTABILIZACAO
                        
        End Select
    
    End If
    
End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDestino_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoOrigem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoOrigem_Click()

Dim lErro As Long

On Error GoTo Erro_TipoOrigem_Click

    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_TipoOrigem_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175484)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub UnidadeMed_Click()

Dim lErro As Long

On Error GoTo Erro_UnidadeMed_Click
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UnidadeMed_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175485)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 30802

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 30802

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175486)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 30803

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 30803

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175487)

    End Select

    Exit Sub

End Sub

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
    If lErro <> SUCESSO Then gError 30804
    
    If colItensNumIntDoc.Count / 2 >= GridMovimentos.Row Then
        lNumIntDoc = colItensNumIntDoc.Item((GridMovimentos.Row * 2) - 1)
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
            If lErro <> SUCESSO And lErro <> 28030 Then gError 30806

            If lErro = 28030 Then gError 30807

            objClasseUM.iClasse = objProduto.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 30808

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


    ElseIf objControl.Name = "Quantidade" Or objControl.Name = "AlmoxOrigem" Or objControl.Name = "TipoOrigem" Or objControl.Name = "AlmoxDestino" Or objControl.Name = "TipoDestino" Or objControl.Name = "ContaContabilEstEntrada" Or objControl.Name = "ContaContabilEstSaida" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
           
    ElseIf objControl.Name = "Lote" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78217
    
            If lErro = 28030 Then gError 78218
        
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
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78319
    
            If lErro = 28030 Then gError 78320
        
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

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 30804, 30806, 30807, 30808, 78217, 78218, 78219, 78319, 78320

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175488)

    End Select

    Exit Sub

End Sub

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
        If lErro <> SUCESSO Then gError 78523
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78524
            
        If lErro = 28030 Then gError 78525
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78526
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78527
                
                'Calcula a Quantidade Origem
                lErro = QuantOrigemLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78822
            
            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                
                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))) > 0 Then
'??? lixo ?
''                    objOrdemProducao.iFilialEmpresa = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))
''                    objOrdemProducao.sCodigo = Lote.Text
''
''                    'Verifica se existe a OP
''                    lErro = CF("OrdemProducao_Le",objOrdemProducao)
''                    If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 78528
''
''                    If lErro = 30368 Then gError 78529
''
''                    If lErro = 55316 Then gError 78530
''
                    objRastroLote.sCodigo = Lote.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))
                    
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 78531
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 78532
                
                    'Calcula a Quantidade Origem
                    lErro = QuantOrigemLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78823
                
                End If
                
            '###############################################################
            'Inserido por Wagner 15/03/2006
            'Se for rastro por série
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                'Preenche a Quantidade do Lote
                lErro = QuantOrigem_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
                If lErro <> SUCESSO Then gError 78824
            '###############################################################
                
            End If
        
        End If
    
        'Se a quantidade está preenchida e não se trata de estorno
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))) > 0 Then

            dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
            
            'Testa a Quantidade requisitada
            lErro = Testa_QuantRequisitada(dQuantidade, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
            If lErro <> SUCESSO Then gError 78812
            
        End If
        
    Else
        'Calcula a Quantidade Origem
        lErro = QuantOrigem_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
        If lErro <> SUCESSO Then gError 78824
    End If
            
    '############################################
    'Inserido por Wagner 15/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141912
    '############################################
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78533

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 78523, 78524, 78526, 78528, 78531, 78533, 78812, 78822, 78823, 78824, 141912
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78527, 78532
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 78529
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175489)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39633

        If objGridInt.objGrid Is GridMovimentos Then

            Select Case objGridInt.objGrid.Col
                
                Case iGrid_AlmoxOrigem_Col
                    lErro = Saida_Celula_AlmoxOrigem(objGridInt)
                    If lErro <> SUCESSO Then gError 30819

                Case iGrid_AlmoxDestino_Col
                    lErro = Saida_Celula_AlmoxDestino(objGridInt)
                    If lErro <> SUCESSO Then gError 30820

                Case iGrid_TipoOrigem_Col
                    lErro = Saida_Celula_TipoOrigem(objGridInt)
                    If lErro <> SUCESSO Then gError 30821

                Case iGrid_TipoDestino_Col
                    lErro = Saida_Celula_TipoDestino(objGridInt)
                    If lErro <> SUCESSO Then gError 30822

                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 30823

                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 30824

                Case iGrid_Estorno_Col
                    lErro = Saida_Celula_Estorno(objGridInt)
                    If lErro <> SUCESSO Then gError 30825

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 30826
                    
                Case iGrid_ContaContabilEstEntrada_Col
                    lErro = Saida_Celula_ContaContabilEstEntrada(objGridInt)
                    If lErro <> SUCESSO Then gError 49669
                    
                Case iGrid_ContaContabilEstSaida_Col
                    lErro = Saida_Celula_ContaContabilEstSaida(objGridInt)
                    If lErro <> SUCESSO Then gError 49670
                        
                Case iGrid_Lote_Col
                    lErro = Saida_Celula_Lote(objGridInt)
                    If lErro <> SUCESSO Then gError 78222
            
                Case iGrid_FilialOP_Col
                    lErro = Saida_Celula_FilialOP(objGridInt)
                    If lErro <> SUCESSO Then gError 78274
            
            End Select

        End If

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30827

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 30819, 30820, 30821, 30822, 30823, 30824, 30825, 30826, 49669, 49670, 78222, 78274

        Case 30827
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39633
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175490)

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
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 78262
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 78263
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 78264
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
                            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 78265
                    
        End If
        
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) > 0 Then
'??? lixo ?
''            objOrdemProducao.iFilialEmpresa = Codigo_Extrai(FilialOP.Text)
''            objOrdemProducao.sCodigo = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)
''
''            'Verifica se existe a OP
''            lErro = CF("OrdemProducao_Le",objOrdemProducao)
''            If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 78560
''
''            If lErro = 30368 Then gError 78561
''
''            If lErro = 55316 Then gError 78562
''
            lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 78578
                                
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objRastroLote.sCodigo = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)
                objRastroLote.sProduto = sProdutoFormatado
                objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
            
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78563
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78564
                
                'Calcula a Quantidade Origem
                lErro = QuantOrigemLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(FilialOP.Text))
                If lErro <> SUCESSO Then gError 78825
                
            End If
            
        End If
        
        'Se a quantidade está preenchida e não se trata de estorno
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))) > 0 Then

            dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
            
            'Testa a Quantidade requisitada
            lErro = Testa_QuantRequisitada(dQuantidade, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
            If lErro <> SUCESSO Then gError 78811
            
        End If
        
    Else
    
        'Calcula a Quantidade Origem
        lErro = QuantOrigem_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
        If lErro <> SUCESSO Then gError 78826
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78266

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 78262, 78263, 78266, 78560, 78563, 78578, 78811, 78825, 78826
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78561
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)

        Case 78562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78564
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175491)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    iProdutoPreenchido = PRODUTO_VAZIO
    
    If Len(Produto.ClipText) <> 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199350

        Produto.Text = sProduto

        lErro = CF("Produto_Critica_Estoque", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 30828

        If lErro = 25077 Then gError 30829

    End If

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = ProdutoLinha_Preenche(objProduto)
        If lErro <> SUCESSO Then gError 30830
        
        'Se o Lote não está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
            'Calcula a Quantidade Disponível
            lErro = QuantOrigem_Calcula1(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), objProduto)
            If lErro <> SUCESSO Then gError 30831
        Else
            lErro = QuantOrigemLote_Calcula1(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78818
        End If
        
        If objProduto.iRastro = PRODUTO_RASTRO_OP Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
        End If
        
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30832
    
    Call Preenche_ContaContabil_Tela

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 30828, 30830, 30831, 61247, 78818, 199350
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30829
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30832
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175492)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sAlmoxarifadoPadrao As String
Dim colRastreamentoSerie As New Collection 'Inserido por Wagner 15/03/2006
Dim objTela As Object

On Error GoTo Erro_ProdutoLinha_Preenche

    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Descricao_Col) = objProduto.sDescricao

    'Almoxarifado
    
    '(Utiliza Almoxarifado Padrão Origem caso esteja preenchido)
        
    If Len(Trim(AlmoxPadraoOrigem.ClipText)) > 0 Then
        lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadraoOrigem)
        If lErro = SUCESSO Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col) = AlmoxPadraoOrigem.Text
        Else
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col) = ""
        End If
        
    Else
    
        'se não está preenchido
        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 52222
        
        'preenche o grid
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col) = sAlmoxarifadoPadrao
        
    End If
    
    '(Utiliza Almoxarifado Padrão Destino caso esteja preenchido)
    If Len(Trim(AlmoxPadraoDestino.ClipText)) > 0 Then
        lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadraoDestino)
        If lErro = SUCESSO Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxDestino_Col) = AlmoxPadraoDestino.Text
        Else
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxDestino_Col) = ""
        End If
        
    Else
        
        'se não está preenchido
        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 52223
        
        'preenche o grid
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxDestino_Col) = sAlmoxarifadoPadrao
        
    End If
    
    'Preenche Estorno com Valor 0 (Checked = False)
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) = "0"

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        gcolcolRastreamentoSerie.Add colRastreamentoSerie 'Inserido por Wagner 15/03/2006

        'Cada Transferência gera dois Movimentos
        colItensNumIntDoc.Add 0
        colItensNumIntDoc.Add 0

    End If
    
    Set objTela = Me
    
    lErro = CF("TRANSF_ProdutoLinha_Preenche", objTela, objProduto)
    If lErro <> SUCESSO Then gError 199431
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr
        
        Case 52222, 52223, 199431
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175493)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.Text)) <> 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 30836

        'Caso QuantOrigem estiver preenchida verificar se é maior
        If colItensNumIntDoc.Item(GridMovimentos.Row * 2 - 1) = 0 Then

            If Len(Trim(QuantOrigem.Caption)) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" Then
                
                lErro = Testa_QuantRequisitada(CDbl(Quantidade.Text), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
                If lErro <> SUCESSO Then gError 30837
                
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
    If lErro <> SUCESSO Then gError 30839

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 30836, 30837, 30839, 61238, 141911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175494)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AlmoxOrigem(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim objProduto As New ClassProduto 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_Saida_Celula_AlmoxOrigem

    Set objGridInt.objControle = AlmoxOrigem

    If Len(Trim(AlmoxOrigem.Text)) <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30844

        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, AlmoxOrigem, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 30845

        If lErro = 25157 Then gError 30846

        If lErro = 25162 Then gError 30847

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col) = objAlmoxarifado.sNomeReduzido

        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141948
        '###########################################################

        'Se o Lote não está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then   'Alterado por Wagner 15/03/2006
            lErro = QuantOrigem_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), AlmoxOrigem.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
            If lErro <> SUCESSO Then gError 30848
        Else
            lErro = QuantOrigemLote_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), AlmoxOrigem.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
            If lErro <> SUCESSO Then gError 78819
        End If
                
    Else
    
        'Limpa a Quantidade Disponível da Tela
        QuantOrigem.Caption = ""
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30849
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    Call Preenche_ContaContabil_Tela

    Saida_Celula_AlmoxOrigem = SUCESSO

    Exit Function

Erro_Saida_Celula_AlmoxOrigem:

    Saida_Celula_AlmoxOrigem = gErr

    'Limpa a Quantidade Disponível da Tela
    QuantOrigem.Caption = ""
    
    Select Case gErr

        Case 30844, 30845, 30848, 30849, 78819, 141948
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30846

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", AlmoxOrigem.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = AlmoxOrigem.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30847

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(AlmoxOrigem.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(AlmoxOrigem.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175495)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AlmoxDestino(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxDestino As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_AlmoxDestino

    Set objGridInt.objControle = AlmoxDestino

    If Len(Trim(AlmoxDestino.Text)) <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30854

        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, AlmoxDestino, objAlmoxDestino)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 30852

        If lErro = 25157 Then gError 30850

        If lErro = 25162 Then gError 30851
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30853
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    Call Preenche_ContaContabil_Tela

    Saida_Celula_AlmoxDestino = SUCESSO

    Exit Function

Erro_Saida_Celula_AlmoxDestino:

    Saida_Celula_AlmoxDestino = gErr

    Select Case gErr

        Case 30850

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", AlmoxDestino.Text)

            If vbMsg = vbYes Then

                objAlmoxDestino.sNomeReduzido = AlmoxDestino.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxDestino)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30851

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(AlmoxDestino.Text))

            If vbMsg = vbYes Then

                objAlmoxDestino.iCodigo = CInt(AlmoxDestino.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxDestino)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30852, 30853, 30854
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175496)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Estorno(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Estorno

    Set objGridInt.objControle = Estorno

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30855

    Saida_Celula_Estorno = SUCESSO

    Exit Function

Erro_Saida_Celula_Estorno:

    Saida_Celula_Estorno = gErr

    Select Case gErr

        Case 30855
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175497)

    End Select

    Exit Function

End Function

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
        
        'Se o Lote não está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantOrigem_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
            If lErro <> SUCESSO Then gError 55415
        Else
            lErro = QuantOrigemLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78820
        End If
        
        If colItensNumIntDoc.Item((GridMovimentos.Row * 2) - 1) = 0 Then
    
            'Se a quantidade está preenchida e não se trata de estorno
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))) > 0 Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
                
                'Testa a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
                If lErro <> SUCESSO Then gError 55416
                
            End If
            
        End If

    Else
    
        QuantOrigem.Caption = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30856

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 30856, 55415, 55416, 78820, 141949, 141950
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175498)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoOrigem(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_TipoOrigem

    Set objGridInt.objControle = TipoOrigem
   
    If Len(TipoOrigem.Text) > 0 And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col))) > 0 Then
        
        'Se o Lote não está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
            lErro = QuantOrigem_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), TipoOrigem.Text)
            If lErro <> SUCESSO Then gError 30848
        Else
            lErro = QuantOrigemLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78821
        End If
        
        If colItensNumIntDoc.Item((GridMovimentos.Row * 2) - 1) = 0 Then
    
            'Se a quantidade está preenchida e não se trata de estorno
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
                
                'Testa a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade, TipoOrigem.Text)
                If lErro <> SUCESSO Then gError 55420
            End If
    
        End If

    Else
    
        QuantOrigem.Caption = ""
        
    End If
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30857

    '############################################
    'Inserido por Wagner 15/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 177288
    '############################################

    Saida_Celula_TipoOrigem = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoOrigem:

    Saida_Celula_TipoOrigem = gErr

    Select Case gErr

        Case 30848, 30857, 55420, 61240, 61249, 78821, 177288
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175499)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_TipoDestino(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TipoDestino

    Set objGridInt.objControle = TipoDestino
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30858

    Saida_Celula_TipoDestino = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDestino:

    Saida_Celula_TipoDestino = gErr

    Select Case gErr

        Case 30858
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175500)

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

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer, lNumIntDoc As Long

    If colItensNumIntDoc.Count >= GridMovimentos.Row * 2 Then
        lNumIntDoc = colItensNumIntDoc.Item((GridMovimentos.Row * 2) - 1)
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
            colItensNumIntDoc.Remove 2 * iLinhaAnterior
            colItensNumIntDoc.Remove 2 * iLinhaAnterior - 1
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

Private Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub AlmoxOrigem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub AlmoxOrigem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub AlmoxOrigem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = AlmoxOrigem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub AlmoxDestino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub AlmoxDestino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub AlmoxDestino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = AlmoxDestino
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoDestino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub TipoDestino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub TipoDestino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = TipoDestino
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoOrigem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub TipoOrigem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub TipoOrigem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = TipoOrigem
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

Private Sub Limpa_Tela_Transfer()

Dim lErro As Long, lCodigo As Long
On Error GoTo Erro_Limpa_Tela_Transfer

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    If objGrid.iProibidoIncluir <> 0 And objGrid.iProibidoExcluir <> 0 Then
        'prepara o Grid para permitir inserir e excluir Linhas
        objGrid.iProibidoIncluir = 0
        objGrid.iProibidoExcluir = 0
        Call Grid_Inicializa(objGrid)
    End If
    
    'Limpa o Label's
    QuantOrigem.Caption = ""

    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection
    
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
    
    'Preenche a DataEntrada com a Data Atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0
    lCodigoAntigo = 0
    
    Set gobjAnotacao = Nothing

    Set gcolcolRastreamentoSerie = New Collection

    Exit Sub
    
Erro_Limpa_Tela_Transfer:

    Select Case gErr
             
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175501)
     
    End Select
    
    Exit Sub
    
End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iMovimento As Integer
Dim iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 30861

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 30864

    'Verifica se há algum movimento informado no Grid
    If Not objGrid.iLinhasExistentes > 0 Then gError 30865

    'Para cada Linha do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Unidade de Medida foi preenchida
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 55423

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 30866

        'Verifica se o AlmoxOrigem foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxOrigem_Col))) = 0 Then gError 30868

        'Verifica se o AlmoxDestino foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxDestino_Col))) = 0 Then gError 30870
        
        'Verifica se o TipoOrigem foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_TipoOrigem_Col))) = 0 Then gError 30867

        'Verifica se o Tipo Destino foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_TipoDestino_Col))) = 0 Then gError 30869
        
    Next

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30128 Then gError 30862
    
    If lErro = SUCESSO Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS2")
        If vbMsgRes = vbNo Then gError 78800
    
    End If
    
    'Tipo de Movimento só deve estar preenchido a nível de Itens
    objMovEstoque.iTipoMov = 0
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 92031
    
    lErro = Move_Tela_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError 30871

    'trata os casos de estorno da versao light
    Call MovEstoque_Trata_Estorno_Versao_Light

    'Grava os dados no BD (inclusive os dados Contábeis)
    lErro = CF("MovEstoque_Grava_Generico", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 30872
    
    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
            gobjAnotacao.sID = CStr(objMovEstoque.iFilialEmpresa) & "," & CStr(objMovEstoque.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 30872
            
        End If
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 30290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_ITEM_NAO_CADASTRADO", gErr, iIndice)

        Case 30321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_MOVTO_ESTOQUE_NAO_CADASTRADO", gErr, objMovEstoque.lCodigo)
            Codigo.SetFocus

        Case 30322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVTO_ESTOQUE_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 30861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 30862, 30871, 30872, 92031, 78800

        Case 30863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVTO_ESTOQUE_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 30864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus

        Case 30865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 30866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 30867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOORIGEM_NAO_INFORMADO", gErr, iIndice)

        Case 30868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXORIGEM_NAO_INFORMADO", gErr, iIndice)

        Case 30869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESTINO_NAO_INFORMADO", gErr, iIndice)

        Case 30870
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXDESTINO_NAO_INFORMADO", gErr, iIndice)

        Case 55423
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175502)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

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
    
    'ATENÇÃO: Cada linha do Grid gera dois Movimentos,
    '         um de Saída e outro , correspondente, de Entrada
    For iIndice = 1 To objGrid.iLinhasExistentes

        lErro = Move_Itens_Memoria(iIndice, objMovEstoque)
        If lErro <> SUCESSO Then gError 30319

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 30319, 30320, 30877

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175503)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCodigo = Nothing
    Set objEventoAlmoxOrigem = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoAlmoxDestino = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoRastroLote = Nothing 'Inserido por Wagner
    
    'Eventos associados a contabilidade
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

Function Move_Itens_Memoria(iIndice As Integer, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim iIndice2 As Integer
Dim iTipo_Col As Integer
Dim iAlmox_Col As Integer
Dim iTipoMov As Integer
Dim sProdutoFormatado As String, sTipoOrigem As String, sTipoDestino As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As ClassAlmoxarifado
Dim iContaContabil As Integer
Dim sContaFormatadaEst As String
Dim iContaPreenchida As Integer
Dim colRatreamentoMovto As New Collection

On Error GoTo Erro_Move_Itens_Memoria

    sProdutoFormatado = ""

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30873

    For iIndice2 = 0 To 1

        If iIndice2 = 1 Then
            iTipo_Col = iGrid_TipoOrigem_Col
            iAlmox_Col = iGrid_AlmoxOrigem_Col
            iContaContabil = iGrid_ContaContabilEstSaida_Col
        Else
            iTipo_Col = iGrid_TipoDestino_Col
            iAlmox_Col = iGrid_AlmoxDestino_Col
            iContaContabil = iGrid_ContaContabilEstEntrada_Col
        End If
        
        If GridMovimentos.TextMatrix(iIndice, iContaContabil) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", GridMovimentos.TextMatrix(iIndice, iContaContabil), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 49661
        
        Else
            sContaFormatadaEst = ""
        End If
        
        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iAlmox_Col)

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 30874

        If lErro = 25056 Then gError 30875
                
            
        sTipoOrigem = GridMovimentos.TextMatrix(iIndice, iGrid_TipoOrigem_Col)
        sTipoDestino = GridMovimentos.TextMatrix(iIndice, iGrid_TipoDestino_Col)
        
        'se estiver fazendo uma transferencia para o mesmo escaninho ==> erro
        If sTipoOrigem = sTipoDestino And GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxOrigem_Col) = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxDestino_Col) Then gError 55421
        
        If (sTipoOrigem = STRING_TRANSF_CONSIG3 Or sTipoOrigem = STRING_TRANSF_CONSIG) Then
             If sTipoDestino <> TRANSF_DISPONIVEL_STRING Then gError 15892
        End If
                
        If (sTipoDestino = STRING_TRANSF_CONSIG3) Then
             If sTipoOrigem <> TRANSF_DISPONIVEL_STRING Then gError 15893
        End If
        
        If (sTipoDestino = TRANSF_OUTRAS_TERC_STRING) Then
             If sTipoOrigem <> TRANSF_OUTRAS_TERC_STRING Then gError 15893
        End If

        iTipoMov = 0

        If GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "1" Then

            If GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_DEFEITUOSO_STRING Then
                
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_DEFEITUOSO
                Else
                    iTipoMov = MOV_EST_ESTORNO_ENTRADA_TRANSF_DEFEITUOSO
                End If
                
            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_DISPONIVEL_STRING Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_DISPONIVEL
                Else
                    iTipoMov = MOV_EST_ESTORNO_ENTRADA_TRANSF_DISPONIVEL
                End If

            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_INDISPONIVEL_STRING Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_INDISPONIVEL
                Else
                   iTipoMov = MOV_EST_ESTORNO_ENTRADA_TRANSF_INDISPONIVEL
                End If

            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = STRING_TRANSF_CONSIG3 Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_TERC
                Else
                   iTipoMov = MOV_EST_ESTORNO_ENTRADA_TRANSF_CONSIG_TERC
                End If
                
            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_OUTRAS_TERC_STRING Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_OUTRAS_TERC
                Else
                   iTipoMov = MOV_EST_ESTORNO_ENTRADA_TRANSF_OUTRAS_TERC
                End If
                
            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = STRING_TRANSF_CONSIG Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_NOSSO
                Else
                   iTipoMov = MOV_EST_ESTORNO_ENTRADA_TRANSF_CONSIG_NOSSO
                End If
                
            End If

        Else
            If GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_DEFEITUOSO_STRING Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_SAIDA_TRANSF_DEFEIT
                Else
                    iTipoMov = MOV_EST_ENTRADA_TRANSF_DEFEIT
                End If

            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_DISPONIVEL_STRING Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_SAIDA_TRANSF_DISP
                Else
                    If sTipoOrigem = STRING_TRANSF_CONSIG3 Then
                        iTipoMov = MOV_EST_ENTRADA_TRANSF_DISP_CONSIG3
                    ElseIf sTipoOrigem = STRING_TRANSF_CONSIG Then
                        iTipoMov = MOV_EST_ENTRADA_TRANSF_DISP_CONSIG
                    Else
                        iTipoMov = MOV_EST_ENTRADA_TRANSF_DISP
                    End If
                End If

            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_INDISPONIVEL_STRING Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_SAIDA_TRANSF_INDISP
                Else
                    iTipoMov = MOV_EST_ENTRADA_TRANSF_INDISP
                End If
                
            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = STRING_TRANSF_CONSIG3 Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_TERC
                Else
                    iTipoMov = MOV_EST_ENTRADA_TRANSF_CONSIG_TERC
                End If
            
            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = TRANSF_OUTRAS_TERC_STRING Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_SAIDA_TRANSF_OUTRAS_TERC
                Else
                    iTipoMov = MOV_EST_ENTRADA_TRANSF_OUTRAS_TERC
                End If
            
            ElseIf GridMovimentos.TextMatrix(iIndice, iTipo_Col) = STRING_TRANSF_CONSIG Then
                If iTipo_Col = iGrid_TipoOrigem_Col Then
                    iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_NOSSO
                Else
                    iTipoMov = MOV_EST_ENTRADA_TRANSF_CONSIG_NOSSO
                End If
            
            End If

        End If
        
        Set colRatreamentoMovto = New Collection
        
        'Move os dados do rastreamento para a Memória
        lErro = Move_RastroEstoque_Memoria(iIndice, colRatreamentoMovto)
        If lErro <> SUCESSO Then gError 78245

        objMovEstoque.colItens.Add colItensNumIntDoc(2 * iIndice - (1 - iIndice2)), iTipoMov, 0, 0, sProdutoFormatado, GridMovimentos.TextMatrix(iIndice, iGrid_Descricao_Col), GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col)), objAlmoxarifado.iCodigo, GridMovimentos.TextMatrix(iIndice, iAlmox_Col), 0, 0, CLng(GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col)), "", "", "", sContaFormatadaEst, 0, colRatreamentoMovto, Nothing, DATA_NULA
        
    Next

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case 15892, 15893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSF_OD", gErr, sTipoOrigem, sTipoDestino)
            
        Case 30873, 30874, 49661, 78245

        Case 30875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 55421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSFERENCIA_MESMO_ESCANINHO", gErr, iIndice)
            
        Case 61231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSFERENCIA_MESMO_ALMOXARIFADO", gErr, iIndice)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175504)

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
    If lErro <> SUCESSO Then gError 78234
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 78235

    If lErro = 28030 Then gError 78236
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            
            'Se não colocou o Número do Lote ---> Erro
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
            
                objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col)
            
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            'se preencheu o lote e não preencheu a filial ==> erro
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
                If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))) = 0 Then gError 78339
                objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col)
                objRastreamentoMovto.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))
            End If
            
            'se preencheu a filialop e não preencheu o lote ==> erro
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))) <> 0 Then
                If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) = 0 Then gError 78237
            End If
        
        '##################################################
        'Inserido por Wagner 15/03/2006
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
           
            For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(iLinha)
                objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
                colRastreamentoMovto.Add objRastreamentoMovto
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
        
        Case 78234, 78235
        
        Case 78236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 78237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 78339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
                
        Case 78827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREENCHIDO", gErr, iLinha)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175505)
    
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

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

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
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
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
                        If lErro <> SUCESSO Then gError 64211

                        objMnemonicoValor.colValor.Add dQuantidadeConvertida
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case QUANT_DISPONIVEL, QUANT_CONSIGNADA1, QUANT_CONSIGNADADETERC1
            
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col)) > 0 Then
                    
                    objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col)
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 75435
                    
                    'Se não encontrou o almoxarifado, erro
                    If lErro = 25060 Then gError 75436
                                        
                    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 75437
                    
                    'Lê as quantidades do Produto No almoxarifado de origem
                    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                    objEstoqueProduto.sProduto = sProdutoFormatado
                    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
                    If lErro <> SUCESSO And lErro <> 21306 Then gError 75438
                    
                    'Se não encontrou Estoque Produto, erro
                    If lErro = 21306 Then gError 75439
                                        
                    If Len(GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                                            
                        If objMnemonicoValor.sMnemonico = QUANT_DISPONIVEL Then
                            dQuantidade = objEstoqueProduto.dQuantDisponivel
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_CONSIGNADA1 Then
                            dQuantidade = objEstoqueProduto.dQuantConsig
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_CONSIGNADADETERC1 Then
                            dQuantidade = objEstoqueProduto.dQuantConsig3
                        End If
                        
                        lErro = CF("UMEstoque_Conversao", sProdutoFormatado, GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dQuantidade, dQuantidadeConvertida)
                        If lErro <> SUCESSO Then gError 75440

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

        Case ALMOX_ORIGEM
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col)
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
            
        Case ALMOX_DESTINO
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxDestino_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxDestino_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case ALMOX_ORIGEM
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case TIPO_ORIGEM
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col)) > 0 Then
                     objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case TIPO_DESTINO
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_TipoDestino_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_TipoDestino_Col)
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
            
        Case CONTACONTABILESTENTRADA1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstEntrada_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstEntrada_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case CONTACONTABILESTSAIDA1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstSaida_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstSaida_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case ESCANINHO_CUSTO
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                                                            
                    Select Case GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col)
                    
                        Case STRING_TRANSF_CONSIG
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_CONSIGNACAO
                            
                        Case STRING_TRANSF_CONSIG3
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO
                                                
                        Case TRANSF_OUTRAS_TERC_STRING
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_OUTROS
                        
                        Case Else
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
                            
                    End Select
                
                Else
                
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case Else
            gError 39653

    End Select

        Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39653
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64211, 75435, 75437, 75438, 75440, 79943
        
        Case 75436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)
                    
        Case 75439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175506)

    End Select

    Exit Function

End Function

Private Function Preenche_ContaContabil_Tela() As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim sAlmoxOrigem As String
Dim sAlmoxDestino As String
Dim sContaMascarada As String
Dim sProduto As String
Dim sConta As String

On Error GoTo Erro_Preenche_ContaContabil_Tela
        
    sAlmoxOrigem = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col)
    sAlmoxDestino = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxDestino_Col)
    sProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
    
    If Len(Trim(sProduto)) > 0 And Len(Trim(sAlmoxOrigem)) > 0 Then
    
        lErro = Preenche_ContaContabilEst(sAlmoxOrigem, sProduto, "Saida")
        If lErro <> SUCESSO Then gError 49702
        
    End If
    
    If Len(Trim(sProduto)) > 0 And Len(Trim(sAlmoxDestino)) > 0 Then
    
        lErro = Preenche_ContaContabilEst(sAlmoxDestino, sProduto, "Entrada")
        If lErro <> SUCESSO Then gError 49703
        
    End If
        
    Preenche_ContaContabil_Tela = SUCESSO
    
    Exit Function
    
Erro_Preenche_ContaContabil_Tela:

    Preenche_ContaContabil_Tela = gErr
    
        Select Case gErr
            
            Case 49702, 49703
            
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175507)
        
        End Select
        
        Exit Function
        
End Function

Function Preenche_ContaContabilEst(sAlmoxarifado As String, sProduto As String, sConta As String) As Long
'funcao de tela que devolve a conta contabil de estoque

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabilEst

    'preenche o obj de almoxarifado para ler o codigo
    objAlmoxarifado.sNomeReduzido = sAlmoxarifado

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25060 Then gError 49698
    
    If lErro = 25060 Then gError 52002

    'Formata o Produto para BD
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 49699

    objEstoqueProduto.sProduto = sProdutoFormatado
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    
    'Le a conta contabil de almoxarifado-produto a partir de codigos de almoxarifado e produto.
    lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 49991 Then gError 49700
    
    If lErro = SUCESSO Then
        
        lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 49701
        
        If sConta = "Entrada" Then
            ContaContabilEstEntrada.PromptInclude = False
            ContaContabilEstEntrada.Text = sContaEnxuta
            ContaContabilEstEntrada.PromptInclude = True
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEstEntrada_Col) = ContaContabilEstEntrada.Text
        
        ElseIf sConta = "Saida" Then
            ContaContabilEstSaida.PromptInclude = False
            ContaContabilEstSaida.Text = sContaEnxuta
            ContaContabilEstSaida.PromptInclude = True
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEstSaida_Col) = ContaContabilEstSaida.Text
    
        End If
    
    End If
   
    Preenche_ContaContabilEst = SUCESSO

    Exit Function

Erro_Preenche_ContaContabilEst:

    Preenche_ContaContabilEst = gErr

    Select Case gErr
        
        Case 49698, 49699, 49700
        
        Case 49701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
                
        Case 52002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175508)
    
    End Select
    
    Exit Function

End Function

Private Function QuantOrigem_Calcula(sProduto As String, sAlmoxarifado As String, sTipo As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantOrigem_Calcula

    If (objProduto Is Nothing) Then
        
        lErro = QuantOrigem_Calcula1(sProduto, sAlmoxarifado, sTipo)
        If lErro <> SUCESSO Then gError 55389
        
    Else
    
        lErro = QuantOrigem_Calcula1(sProduto, sAlmoxarifado, sTipo, objProduto)
        If lErro <> SUCESSO Then gError 55390

    End If

    lErro = Testa_Quantidade(sTipo)
    If lErro <> SUCESSO Then gError 55391

    QuantOrigem_Calcula = SUCESSO

    Exit Function

Erro_QuantOrigem_Calcula:

    QuantOrigem_Calcula = gErr

    Select Case gErr

        Case 55389, 55390, 55391, 61250, 61251

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175509)

    End Select

    Exit Function

End Function

Private Function QuantOrigem_Calcula1(sProduto As String, sAlmoxarifado As String, sTipo As String, Optional objProduto As ClassProduto) As Long
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

On Error GoTo Erro_QuantOrigem_Calcula1

    QuantOrigem.Caption = ""
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30810
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) > 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then
    
            If objProduto Is Nothing Then
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 30811
    
                If lErro = 28030 Then gError 30812
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 30813
    
            If lErro = 25056 Then gError 30814
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 30815
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantOrigem.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 30816
        
                If sTipo = TRANSF_DISPONIVEL_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel * dFator)
                ElseIf sTipo = TRANSF_DEFEITUOSO_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objEstoqueProduto.dQuantDefeituosa * dFator)
                ElseIf sTipo = TRANSF_INDISPONIVEL_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objEstoqueProduto.dQuantInd * dFator)
                ElseIf sTipo = STRING_TRANSF_CONSIG3 Then
                    QuantOrigem.Caption = Formata_Estoque(objEstoqueProduto.dQuantConsig3 * dFator)
                ElseIf sTipo = TRANSF_OUTRAS_TERC_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objEstoqueProduto.dQuantOutras3 * dFator)
                ElseIf sTipo = STRING_TRANSF_CONSIG Then
                    QuantOrigem.Caption = Formata_Estoque(objEstoqueProduto.dQuantConsig * dFator)
                End If
    
            End If
    
        End If

    End If
    
    QuantOrigem_Calcula1 = SUCESSO

    Exit Function

Erro_QuantOrigem_Calcula1:

    QuantOrigem_Calcula1 = gErr

    Select Case gErr

        Case 30810, 30811, 30813, 30815, 30816

        Case 30812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 30814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175510)

    End Select

    Exit Function

End Function

Private Function QuantOrigem_Calcula1_Light(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long
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

On Error GoTo Erro_QuantOrigem_Calcula1_Light

    QuantOrigem.Caption = ""
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 61242
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) > 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If objProduto Is Nothing Then
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 61243
    
                If lErro = 28030 Then gError 61244
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 61245
    
            If lErro = 25056 Then gError 61246
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 61247
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantOrigem.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 61248
        
                QuantOrigem.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel * dFator)
    
            End If
    
        End If

    End If
    
    QuantOrigem_Calcula1_Light = SUCESSO

    Exit Function

Erro_QuantOrigem_Calcula1_Light:

    QuantOrigem_Calcula1_Light = gErr

    Select Case gErr

        Case 61242, 61243, 61245, 61247, 61248

        Case 61244
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 61246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175511)

    End Select

    Exit Function

End Function

Private Function QuantOrigemLote_Calcula(sProduto As String, sAlmoxarifado As String, sTipo As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantOrigemLote_Calcula

    If (objProduto Is Nothing) Then
        
        lErro = QuantOrigemLote_Calcula1(sProduto, sAlmoxarifado, sTipo, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78801
        
    Else
    
        lErro = QuantOrigemLote_Calcula1(sProduto, sAlmoxarifado, sTipo, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)), objProduto)
        If lErro <> SUCESSO Then gError 78802

    End If

    lErro = Testa_Quantidade(sTipo)
    If lErro <> SUCESSO Then gError 78803

    QuantOrigemLote_Calcula = SUCESSO

    Exit Function

Erro_QuantOrigemLote_Calcula:

    QuantOrigemLote_Calcula = gErr

    Select Case gErr

        Case 78801, 78802, 78803

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175512)

    End Select

    Exit Function

End Function

Private Function QuantOrigemLote_Calcula1(sProduto As String, sAlmoxarifado As String, sTipo As String, sLote As String, iFilialOP As Integer, Optional objProduto As ClassProduto) As Long
'Descobre a quantidade disponivel e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoLoteSaldo As New ClassRastreamentoLoteSaldo

On Error GoTo Erro_QuantOrigemLote_Calcula1

    QuantOrigem.Caption = ""
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78804
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) > 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then
    
            If objProduto Is Nothing Then
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 78805
    
                If lErro = 28030 Then gError 78806
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 78807
    
            If lErro = 25056 Then gError 78808
    
            objRastreamentoLoteSaldo.iAlmoxarifado = objAlmoxarifado.iCodigo
            objRastreamentoLoteSaldo.sProduto = sProdutoFormatado
            objRastreamentoLoteSaldo.sLote = sLote
            objRastreamentoLoteSaldo.iFilialOP = iFilialOP
            
            'Lê a quantidade do Produto no Lote
            lErro = CF("RastreamentoLoteSaldo_Le", objRastreamentoLoteSaldo)
            If lErro <> SUCESSO And lErro <> 78633 Then gError 78809
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 78633 Then
            
                 QuantOrigem.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 78810
        
                If sTipo = TRANSF_DISPONIVEL_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantDispNossa * dFator)
                ElseIf sTipo = TRANSF_DEFEITUOSO_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantDefeituosa * dFator)
                ElseIf sTipo = TRANSF_INDISPONIVEL_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantIndOutras * dFator)
                ElseIf sTipo = STRING_TRANSF_CONSIG3 Then
                    QuantOrigem.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantConsig3 * dFator)
                ElseIf sTipo = TRANSF_OUTRAS_TERC_STRING Then
                    QuantOrigem.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantOutras3 * dFator)
                ElseIf sTipo = STRING_TRANSF_CONSIG Then
                    QuantOrigem.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantConsig * dFator)
                End If
    
            End If
    
        End If

    End If
    
    QuantOrigemLote_Calcula1 = SUCESSO

    Exit Function

Erro_QuantOrigemLote_Calcula1:

    QuantOrigemLote_Calcula1 = gErr

    Select Case gErr

        Case 78804, 78805, 78807, 78809, 78810

        Case 78806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 78808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175513)

    End Select

    Exit Function

End Function

Private Function Testa_Quantidade(sTipo As String) As Long

Dim dQuantidade As Double
Dim lErro As Long

On Error GoTo Erro_Testa_Quantidade

    If GridMovimentos.Row >= GridMovimentos.FixedRows Then

        If colItensNumIntDoc.Item((GridMovimentos.Row * 2) - 1) = 0 Then
    
            'Se a quantidade está preenchida e não se trata de linha estornada
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) > 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" And Len(Trim(sTipo)) > 0 Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
                
                'Calcula a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade, sTipo)
                If lErro <> SUCESSO Then gError 30817
            End If
    
        End If

    End If

    Testa_Quantidade = SUCESSO

    Exit Function

Erro_Testa_Quantidade:

    Testa_Quantidade = gErr

    Select Case gErr

        Case 30817, 61241 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175514)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_QuantRequisitada(ByVal dQuantAtual As Double, sTipo As String) As Long

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
Dim objProduto As New ClassProduto, sLoteAtual As String, sLote As String
Dim dQuantTotal As Double, iFilialOPAtual As Integer, iFilialOP As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Testa_QuantRequisitada

    sProdutoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
    sAlmoxarifadoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col)
    sUnidadeAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
    iFilialOPAtual = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))
    sLoteAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)

    If Len(sProdutoAtual) > 0 And Len(sAlmoxarifadoAtual) > 0 And Len(sUnidadeAtual) > 0 And Len(Trim(sTipo)) > 0 Then

        lErro = CF("Produto_Formata", sProdutoAtual, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30840

        objProduto.sCodigo = sProdutoFormatado

        'Lê o produto para saber qual é a sua ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 30841
    
        If lErro = 28030 Then gError 30842
    
        For iIndice = 1 To objGrid.iLinhasExistentes
    
            'Não pode somar a Linha atual
            If GridMovimentos.Row <> iIndice Then
    
                sCodProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)
                sAlmoxarifado = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxOrigem_Col)
                iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(iIndice, iGrid_FilialOP_Col))
                sLote = GridMovimentos.TextMatrix(iIndice, iGrid_Lote_Col)
    
                lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 55413
    
                'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) And GridMovimentos.TextMatrix(iIndice, iGrid_TipoOrigem_Col) = sTipo And iFilialOPAtual = iFilialOP And UCase(sLoteAtual) = UCase(sLote) Then
    
                    'Verifica se há alguma QuanTidade informada
                    If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) <> 0 Then
    
                        sUnidadeProd = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
                        dQuantidadeProd = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
    
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                        If lErro <> SUCESSO Then gError 30843
    
                        dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
    
                    End If
    
                End If
    
            End If
    
        Next
    
        dQuantTotal = dQuantTotal + dQuantAtual

        If dQuantTotal > StrParaDbl(QuantOrigem.Caption) Then vbMsg = Rotina_Aviso(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)

    End If

    Testa_QuantRequisitada = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada:

    Testa_QuantRequisitada = gErr

    Select Case gErr

        Case 30840, 30841, 30843, 55413

        Case 30842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case 55414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175515)

    End Select

    Exit Function

End Function

Private Function Testa_QuantRequisitada_Light(ByVal dQuantAtual As Double) As Long

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

On Error GoTo Erro_Testa_QuantRequisitada_Light

    sProdutoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
    sAlmoxarifadoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col)
    sUnidadeAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)

    If Len(sProdutoAtual) > 0 And Len(sAlmoxarifadoAtual) > 0 And Len(sUnidadeAtual) > 0 Then

        lErro = CF("Produto_Formata", sProdutoAtual, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 61232

        objProduto.sCodigo = sProdutoFormatado

        'Lê o produto para saber qual é a sua ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 61233
    
        If lErro = 28030 Then gError 61234
    
        For iIndice = 1 To objGrid.iLinhasExistentes
    
            'Não pode somar a Linha atual
            If GridMovimentos.Row <> iIndice Then
    
                sCodProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)
                sAlmoxarifado = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxOrigem_Col)
    
                lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 61235
    
                'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) Then
    
                    'Verifica se há alguma QuanTidade informada
                    If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) <> 0 Then
    
                        sUnidadeProd = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
                        dQuantidadeProd = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
    
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                        If lErro <> SUCESSO Then gError 61236
    
                        dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
    
                    End If
    
                End If
    
            End If
    
        Next
    
        dQuantTotal = dQuantTotal + dQuantAtual

        If dQuantTotal > StrParaDbl(QuantOrigem.Caption) Then gError 61237

    End If

    Testa_QuantRequisitada_Light = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada_Light:

    Testa_QuantRequisitada_Light = gErr

    Select Case gErr

        Case 61232, 61233, 61235, 61236

        Case 61234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case 61237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175516)

    End Select

    Exit Function

End Function

Private Sub MovEstoque_Trata_Estorno_Versao_Light()

Dim iEstorno As Integer
Dim iLinha As Integer

    If giTipoVersao = VERSAO_LIGHT Then

        iEstorno = 0

        For iLinha = 1 To objGrid.iLinhasExistentes

            If GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col) = MARCADO Then

                iEstorno = 1
                Exit For

            End If

        Next

        Call objContabil.Contabil_Trata_Estorno_Versao_Light(iEstorno)
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TRANSFERENCIA_MOVIMENTO
    Set Form_Load_Ocx = Me
    Caption = "Transferência"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Transfer"
    
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
        ElseIf Me.ActiveControl Is AlmoxPadraoOrigem Then
            Call AlmoxOrigemLabel_Click
        ElseIf Me.ActiveControl Is AlmoxPadraoDestino Then
            Call AlmoxDestinoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is ContaContabilEstEntrada Or Me.ActiveControl Is ContaContabilEstSaida Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is AlmoxOrigem Or Me.ActiveControl Is AlmoxDestino Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is Lote Then 'Inserido por Wagner
            Call BotaoLote_Click
        End If
        
'    ElseIf KeyCode = KEYCODE_CODBARRAS Then
'        Call Trata_CodigoBarras
    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        Call Trata_CodigoBarras1
        
    End If

End Sub


Private Sub AlmoxOrigemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxOrigemLabel, Source, X, Y)
End Sub

Private Sub AlmoxOrigemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxOrigemLabel, Button, Shift, X, Y)
End Sub

Private Sub AlmoxDestinoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxDestinoLabel, Source, X, Y)
End Sub

Private Sub AlmoxDestinoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxDestinoLabel, Button, Shift, X, Y)
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub QuantOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantOrigem, Source, X, Y)
End Sub

Private Sub QuantOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantOrigem, Button, Shift, X, Y)
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

'Já existe em Rastreamento Lote
Function RastreamentoLote_Le(objRastroLote As ClassRastreamentoLote) As Long
'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados

Dim lErro As Long
Dim lComando As Long
Dim tRastroLote As typeRastreamentoLote

On Error GoTo Erro_RastreamentoLote_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 75707

    tRastroLote.sObservacao = String(STRING_NOME, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_Executar(lComando, "SELECT DataValidade, DataEntrada, DataFabricacao, Observacao FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", tRastroLote.dtDataValidade, tRastroLote.dtDataEntrada, tRastroLote.dtDataFabricacao, tRastroLote.sObservacao, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 75708

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75709

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 75710

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RastreamentoLote_Le = SUCESSO

    Exit Function

Erro_RastreamentoLote_Le:

    RastreamentoLote_Le = gErr
    
    Select Case gErr

        Case 75707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 75708, 75709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)

        Case 75710 'RastreamentoLote não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175517)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function


'Já foi utilizada em outras telas do est
Function RastreamentoMovto_Le_DocOrigem(lNumIntDocOrigem As Long, iTipoDocOrigem As Integer, colRastreamentoMovto As Collection) As Long
'Lê a tabela de RastreamentoMovto através do Movimento de Estoque

Dim lErro As Long
Dim tRastreamentoMovto As typeRastreamentoMovto
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim lComando As Long

On Error GoTo Erro_RastreamentoMovto_Le_DocOrigem
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 78411
    
    tRastreamentoMovto.sProduto = String(STRING_PRODUTO, 0)
    tRastreamentoMovto.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    
    'Lê o Rastreamento Movto
    lErro = Comando_Executar(lComando, "SELECT RastreamentoMovto.NumIntDoc, RastreamentoMovto.TipoDocOrigem, RastreamentoMovto.NumIntDocOrigem, RastreamentoMovto.Produto, RastreamentoMovto.Quantidade, RastreamentoLote.Lote, RastreamentoLote.FilialOP FROM RastreamentoMovto, RastreamentoLote WHERE RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND TipoDocOrigem = ? AND NumIntDocOrigem = ?" _
    , tRastreamentoMovto.lNumIntDoc, tRastreamentoMovto.iTipoDocOrigem, tRastreamentoMovto.lNumIntDocOrigem, tRastreamentoMovto.sProduto, tRastreamentoMovto.dQuantidade, tRastreamentoMovto.sLote, tRastreamentoMovto.iFilialOP, iTipoDocOrigem, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 78412
           
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78413

    Do While lErro = AD_SQL_SUCESSO
        
        Set objRastreamentoMovto = New ClassRastreamentoMovto
        
        'passa para o objeto
        objRastreamentoMovto.lNumIntDoc = tRastreamentoMovto.lNumIntDoc
        objRastreamentoMovto.iTipoDocOrigem = tRastreamentoMovto.iTipoDocOrigem
        objRastreamentoMovto.lNumIntDocOrigem = tRastreamentoMovto.lNumIntDocOrigem
        objRastreamentoMovto.sProduto = tRastreamentoMovto.sProduto
        objRastreamentoMovto.dQuantidade = tRastreamentoMovto.dQuantidade
        objRastreamentoMovto.sLote = tRastreamentoMovto.sLote
        objRastreamentoMovto.iFilialOP = tRastreamentoMovto.iFilialOP
        
        colRastreamentoMovto.Add objRastreamentoMovto
                
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78421
    
    Loop
            
    Call Comando_Fechar(lComando)
    
    RastreamentoMovto_Le_DocOrigem = SUCESSO
        
    Exit Function
    
Erro_RastreamentoMovto_Le_DocOrigem:

    RastreamentoMovto_Le_DocOrigem = gErr
    
    Select Case gErr
            
        Case 78411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 78412, 78413, 78421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175518)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Copiada da class CTInventario de GlobaisTelasEst
Function RastreamentoLoteSaldo_Le(objRastroLoteSaldo As ClassRastreamentoLoteSaldo) As Long
'Lê a tabela de Rastreamento Lote Saldo

Dim lErro As Long
Dim lComando As Long
Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo

On Error GoTo Erro_RastreamentoLoteSaldo_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 78630

    tRastroLoteSaldo.sProduto = String(STRING_PRODUTO, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_Executar(lComando, "SELECT RastreamentoLoteSaldo.Produto, Almoxarifado, NumIntDocLote, QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3 FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLote.NumIntDoc = RastreamentoLoteSaldo.NumIntDocLote AND RastreamentoLoteSaldo.Produto = ? AND Almoxarifado = ? AND RastreamentoLote.Lote = ? AND RastreamentoLote.FilialOP = ?", _
    tRastroLoteSaldo.sProduto, tRastroLoteSaldo.iAlmoxarifado, tRastroLoteSaldo.lNumIntDocLote, tRastroLoteSaldo.dQuantDispNossa, tRastroLoteSaldo.dQuantReservada, tRastroLoteSaldo.dQuantReservadaConsig, tRastroLoteSaldo.dQuantEmpenhada, tRastroLoteSaldo.dQuantPedida, tRastroLoteSaldo.dQuantRecIndl, tRastroLoteSaldo.dQuantIndOutras, tRastroLoteSaldo.dQuantDefeituosa, tRastroLoteSaldo.dQuantConsig3, tRastroLoteSaldo.dQuantConsig, tRastroLoteSaldo.dQuantDemo3, tRastroLoteSaldo.dQuantDemo, tRastroLoteSaldo.dQuantConserto3, tRastroLoteSaldo.dQuantConserto, tRastroLoteSaldo.dQuantOutras3, tRastroLoteSaldo.dQuantOutras, tRastroLoteSaldo.dQuantOP, tRastroLoteSaldo.dQuantBenef, tRastroLoteSaldo.dQuantBenef3, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote, objRastroLoteSaldo.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 78631

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78632

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 78633

    objRastroLoteSaldo.sProduto = tRastroLoteSaldo.sProduto
    objRastroLoteSaldo.iAlmoxarifado = tRastroLoteSaldo.iAlmoxarifado
    objRastroLoteSaldo.lNumIntDocLote = tRastroLoteSaldo.lNumIntDocLote
    objRastroLoteSaldo.dQuantDispNossa = tRastroLoteSaldo.dQuantDispNossa
    objRastroLoteSaldo.dQuantReservada = tRastroLoteSaldo.dQuantReservada
    objRastroLoteSaldo.dQuantReservadaConsig = tRastroLoteSaldo.dQuantReservadaConsig
    objRastroLoteSaldo.dQuantEmpenhada = tRastroLoteSaldo.dQuantEmpenhada
    objRastroLoteSaldo.dQuantPedida = tRastroLoteSaldo.dQuantPedida
    objRastroLoteSaldo.dQuantRecIndl = tRastroLoteSaldo.dQuantRecIndl
    objRastroLoteSaldo.dQuantIndOutras = tRastroLoteSaldo.dQuantIndOutras
    objRastroLoteSaldo.dQuantDefeituosa = tRastroLoteSaldo.dQuantDefeituosa
    objRastroLoteSaldo.dQuantConsig3 = tRastroLoteSaldo.dQuantConsig3
    objRastroLoteSaldo.dQuantConsig = tRastroLoteSaldo.dQuantConsig
    objRastroLoteSaldo.dQuantDemo3 = tRastroLoteSaldo.dQuantDemo3
    objRastroLoteSaldo.dQuantDemo = tRastroLoteSaldo.dQuantDemo
    objRastroLoteSaldo.dQuantConserto3 = tRastroLoteSaldo.dQuantConserto3
    objRastroLoteSaldo.dQuantConserto = tRastroLoteSaldo.dQuantConserto
    objRastroLoteSaldo.dQuantOutras3 = tRastroLoteSaldo.dQuantOutras3
    objRastroLoteSaldo.dQuantOutras = tRastroLoteSaldo.dQuantOutras
    objRastroLoteSaldo.dQuantOP = tRastroLoteSaldo.dQuantOP
    objRastroLoteSaldo.dQuantBenef = tRastroLoteSaldo.dQuantBenef
    objRastroLoteSaldo.dQuantBenef3 = tRastroLoteSaldo.dQuantBenef3

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RastreamentoLoteSaldo_Le = SUCESSO

    Exit Function

Erro_RastreamentoLoteSaldo_Le:

    RastreamentoLoteSaldo_Le = gErr

    Select Case gErr

        Case 78630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78631, 78632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO", gErr, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote)

        Case 78633 'RastreamentoLoteSaldo não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175519)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

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
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col))) = 0 Then gError 177296
    
    Set objAlmoxarifado = New ClassAlmoxarifado
    
    objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col)

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
    
    
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        'Preenche a Quantidade
        lErro = QuantOrigem_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
        If lErro <> SUCESSO Then gError 140228
    Else
        'Preenche a Quantidade do Lote
        lErro = QuantOrigemLote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
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
'Inserido por Wagner 15/03/2006
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
    objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxOrigem_Col)
    objItemMovEstoque.sProduto = sProdutoFormatado
    objItemMovEstoque.sSiglaUM = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) = "1" Then
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_DEFEITUOSO_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_DEFEITUOSO
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_DISPONIVEL_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_DISPONIVEL
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_INDISPONIVEL_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_INDISPONIVEL
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG3 Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_TERC
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_OUTRAS_TERC_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_OUTRAS_TERC
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_NOSSO
        End If
    Else
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_DEFEITUOSO_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_SAIDA_TRANSF_DEFEIT
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_DISPONIVEL_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_SAIDA_TRANSF_DISP
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_INDISPONIVEL_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_SAIDA_TRANSF_INDISP
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG3 Then
            objItemMovEstoque.iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_TERC
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = TRANSF_OUTRAS_TERC_STRING Then
            objItemMovEstoque.iTipoMov = MOV_EST_SAIDA_TRANSF_OUTRAS_TERC
        ElseIf GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG Then
            objItemMovEstoque.iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_NOSSO
        End If
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
    
        If GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col) = "1" Then
            If GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_DEFEITUOSO_STRING Then
                iTipoMovto = MOV_EST_ESTORNO_SAIDA_TRANSF_DEFEITUOSO
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_DISPONIVEL_STRING Then
                iTipoMovto = MOV_EST_ESTORNO_SAIDA_TRANSF_DISPONIVEL
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_INDISPONIVEL_STRING Then
                iTipoMovto = MOV_EST_ESTORNO_SAIDA_TRANSF_INDISPONIVEL
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG3 Then
                iTipoMovto = MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_TERC
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_OUTRAS_TERC_STRING Then
                iTipoMovto = MOV_EST_ESTORNO_SAIDA_TRANSF_OUTRAS_TERC
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG Then
                iTipoMovto = MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_NOSSO
            End If
        Else
            If GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_DEFEITUOSO_STRING Then
                iTipoMovto = MOV_EST_SAIDA_TRANSF_DEFEIT
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_DISPONIVEL_STRING Then
                iTipoMovto = MOV_EST_SAIDA_TRANSF_DISP
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_INDISPONIVEL_STRING Then
                iTipoMovto = MOV_EST_SAIDA_TRANSF_INDISP
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG3 Then
                iTipoMovto = MOV_EST_SAIDA_TRANSF_CONSIG_TERC
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = TRANSF_OUTRAS_TERC_STRING Then
                iTipoMovto = MOV_EST_SAIDA_TRANSF_OUTRAS_TERC
            ElseIf GridMovimentos.TextMatrix(iLinha, iGrid_TipoOrigem_Col) = STRING_TRANSF_CONSIG Then
                iTipoMovto = MOV_EST_SAIDA_TRANSF_CONSIG_NOSSO
            End If
        End If
        
        If dQuantidadeAnterior <> 0 And Len(Trim(sLoteIniAnterior)) <> 0 And iTipoMovtoAnt <> iTipoMovto Then
            
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
                objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovimentos.TextMatrix(iLinha, iGrid_AlmoxOrigem_Col)
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
    
        Case 141921, 141922, 141925, 141927, 141920, 141929, 177240, 177237
        
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
                
    'Caso QuantOrigem estiver preenchida verificar se é maior
    If colItensNumIntDoc.Item(iLinha * 2 - 1) = 0 Then

        If Len(Trim(QuantOrigem.Caption)) <> 0 And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) <> "1" Then
            
            lErro = Testa_QuantRequisitada(dQuantidade, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoOrigem_Col))
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


'Private Sub Trata_CodigoBarras()
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim sProdutoEnxuto As String
'
'On Error GoTo Erro_Trata_CodigoBarras
'
'    If objGrid.iLinhasExistentes + 1 = GridMovimentos.Row Then
'
'        'Verifica se o Produto está preenchido
'        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then
'
'            Call Chama_Tela_Modal("CodigoBarras", objProduto)
'
'            If objProduto.sCodigoBarras <> "Cancel" Then
'
'                'Lê os demais atributos do Produto
'                lErro = CF("Produto_Le", objProduto)
'                If lErro <> SUCESSO And lErro <> 28030 Then gError 199300
'
'                'Se não encontrou o Produto --> Erro
'                If lErro = 28030 Then gError 199301
'
'                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
'                If lErro <> SUCESSO Then gError 199302
'
'                Produto.PromptInclude = False
'                Produto.Text = sProdutoEnxuto
'                Produto.PromptInclude = True
''                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
''
''                'Preenche a Linha do Grid
''                lErro = ProdutoLinha_Preenche(objProduto)
''                If lErro <> SUCESSO Then gError 199303
'
'            End If
'
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_Trata_CodigoBarras:
'
'    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = ""
'
'    Select Case gErr
'
'        Case 199300, 199303
'
'        Case 199301
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
'
'        Case 199302
'            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199304)
'
'    End Select
'
'    Exit Sub
'
'End Sub

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
                    If lErro <> SUCESSO Then gError 210845
                    
            End If
            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 210846
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 210847
    
                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 210848
    
                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 210849
        
                Me.Show
        
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
                
                gError 210862
                
            Else
            
                gError 210851
    
    
            End If
            
'            GridMovimentos.SetFocus
'            GridMovimentos.FocusRect = flexFocusHeavy
    
        End If
    
    End If

    Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_Trata_CodigoBarras1:

    Trata_CodigoBarras1 = gErr

'    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 210845 To 210847, 210850, 210851, 210862

        Case 210848
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 210849
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210852)

    End Select

    Exit Function

End Function


