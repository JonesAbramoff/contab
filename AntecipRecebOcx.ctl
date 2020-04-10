VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl AntecipRecebOcx 
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   9330
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   1
      Left            =   195
      TabIndex        =   0
      Top             =   795
      Width           =   8985
      Begin VB.Frame Frame3 
         Caption         =   "Dados Principais"
         Height          =   2175
         Left            =   195
         TabIndex        =   38
         Top             =   -15
         Width           =   8415
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5295
            TabIndex        =   2
            Top             =   285
            Width           =   1935
         End
         Begin VB.ComboBox CodConta 
            Height          =   315
            Left            =   1800
            TabIndex        =   3
            Top             =   780
            Width           =   1695
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   300
            Left            =   2910
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1267
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   300
            Left            =   5295
            TabIndex        =   4
            Top             =   1260
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1785
            TabIndex        =   5
            Top             =   1267
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1785
            TabIndex        =   1
            Top             =   292
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Natureza 
            Height          =   315
            Left            =   1770
            TabIndex        =   80
            Top             =   1695
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   825
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   1725
            Width           =   840
         End
         Begin VB.Label LabelNaturezaDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3000
            TabIndex        =   81
            Top             =   1695
            Width           =   4815
         End
         Begin VB.Label LabelSeqMovto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5280
            TabIndex        =   78
            Top             =   750
            Width           =   1860
         End
         Begin VB.Label LabelSequencial 
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
            Left            =   4140
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   77
            Top             =   810
            Width           =   1020
         End
         Begin VB.Label Label3 
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
            Left            =   4695
            TabIndex        =   46
            Top             =   345
            Width           =   465
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
            Left            =   1020
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   47
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Label14 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4665
            TabIndex        =   48
            Top             =   1320
            Width           =   510
         End
         Begin VB.Label Label15 
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
            Left            =   1200
            TabIndex        =   49
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label LabelCodConta 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1110
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   840
            Width           =   570
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Recebimento"
         Height          =   720
         Left            =   195
         TabIndex        =   45
         Top             =   2310
         Width           =   8415
         Begin VB.ComboBox TipoMeioPagto 
            Height          =   315
            Left            =   1770
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1095
            TabIndex        =   51
            Top             =   300
            Width           =   585
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Complemento"
         Height          =   1230
         Left            =   195
         TabIndex        =   39
         Top             =   3075
         Width           =   8415
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1755
            TabIndex        =   7
            Top             =   240
            Width           =   5085
         End
         Begin MSMask.MaskEdBox NumRefExterna 
            Height          =   300
            Left            =   1755
            TabIndex        =   8
            Top             =   682
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Externo:"
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
            TabIndex        =   52
            Top             =   735
            Width           =   1185
         End
         Begin VB.Label Label11 
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
            Height          =   195
            Left            =   825
            TabIndex        =   53
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Situação Atual"
         Height          =   765
         Left            =   195
         TabIndex        =   40
         Top             =   4380
         Width           =   8415
         Begin VB.CommandButton BotaoSaldo 
            Caption         =   "Utilização do Saldo do Adiantamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3285
            TabIndex        =   9
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton BotaoBaixas 
            Caption         =   "Baixas de Títulos a Receber com o Adiantamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   6585
            TabIndex        =   11
            Top             =   120
            Width           =   1770
         End
         Begin VB.CommandButton Botao_AntecipReceb 
            Caption         =   "Adiantamentos da Filial do Cliente..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   4905
            TabIndex        =   10
            Top             =   120
            Width           =   1650
         End
         Begin VB.Label Label5 
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
            Left            =   1020
            TabIndex        =   54
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Saldo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   285
            Left            =   1725
            TabIndex        =   55
            Top             =   308
            Width           =   1530
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   9030
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4800
         TabIndex        =   79
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
         Left            =   6240
         TabIndex        =   18
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
         Left            =   6270
         TabIndex        =   16
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   20
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
         Left            =   7695
         TabIndex        =   17
         Top             =   60
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   26
         Top             =   1920
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
         TabIndex        =   28
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   27
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   30
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   41
         Top             =   3450
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
            TabIndex        =   56
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
            TabIndex        =   57
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   58
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   59
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
         TabIndex        =   21
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   22
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   42
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
         Left            =   570
         TabIndex        =   15
         Top             =   525
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
         TabIndex        =   14
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
         Left            =   3810
         TabIndex        =   13
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
         TabIndex        =   29
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
         Left            =   6330
         TabIndex        =   31
         Top             =   1515
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
         Left            =   6330
         TabIndex        =   32
         Top             =   1515
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
         Left            =   6300
         TabIndex        =   19
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
         TabIndex        =   60
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   61
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
         TabIndex        =   62
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   63
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
         Top             =   945
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
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   70
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   71
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   72
         Top             =   3030
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
         TabIndex        =   73
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
         TabIndex        =   74
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
         TabIndex        =   75
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7095
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "AntecipRecebOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "AntecipRecebOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "AntecipRecebOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "AntecipRecebOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5715
      Left            =   75
      TabIndex        =   43
      Top             =   390
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   10081
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
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
   Begin VB.Label Label4 
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
      Left            =   3900
      TabIndex        =   76
      Top             =   2430
      Width           =   615
   End
End
Attribute VB_Name = "AntecipRecebOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'início da Contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const FILIAL_COD As String = "FilialCli_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialCli_Nome"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialCli_CGC_CPF"
Private Const CONTA_COD As String = "Conta_Codigo"
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta"
Private Const SEQUENCIAL1 As String = "Sequencial"
Private Const VALOR1 As String = "Valor"
Private Const HISTORICO1 As String = "Historico"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const DATA1 As String = "Data"
Private Const DOCEXTERNO As String = "Docto_Externo"

Dim iFrameAtual As Integer
Dim glSequencial As Long
Public iAlterado As Integer
Dim iClienteAlterado As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoAntecipRec As AdmEvento
Attribute objEventoAntecipRec.VB_VarHelpID = -1
Private WithEvents objEventoCodConta As AdmEvento
Attribute objEventoCodConta.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2

Private Sub Botao_AntecipReceb_Click()

Dim objAntecipReceb As New ClassAntecipReceb
Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim lErro As Long
Dim lPosicaoSeparador As Long
Dim iCodFilial As Integer

On Error GoTo Erro_BotaoAntecipRec_Click

    'Se Cliente não está preenchido
    If Len(Trim(Cliente.Text)) = 0 Then Error 15447

    'Se Filial não está preenchida
    If Len(Trim(Filial.Text)) = 0 Then Error 15448

    'Lê os dados do Cliente que está na tela
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then Error 15449
    
    'Filtro
    colSelecao.Add objCliente.lCodigo
    colSelecao.Add Codigo_Extrai(Filial.Text)

    'Abre o Browse de Antecipações de recebimento de uma Filial
    Call Chama_Tela("AntecipRecebLista", colSelecao, objAntecipReceb, objEventoAntecipRec)

    Exit Sub

Erro_BotaoAntecipRec_Click:

    Select Case Err

        Case 15447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 15448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 15449
            Cliente.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142882)

    End Select

    Exit Sub

End Sub

Private Sub BotaoBaixas_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_BotaoBaixas_Click

    'Verifica se a Conta corrente foi informada
    If Len(Trim(CodConta.Text)) = 0 Then Error 15455

    'Verifica se o Sequencial foi informado
    If glSequencial = 0 Then Error 15451
   
    'Passa o Código da Conta que está na tela para o Obj
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodConta.Text)
    objMovContaCorrente.lSequencial = glSequencial

    lErro = CF("MovContaCorrente_ObterNumMovto", objMovContaCorrente)
    If lErro <> SUCESSO Then Error 15474
    
    'Filtro
    colSelecao.Add objMovContaCorrente.lNumMovto

    'Abre o Browse de Antecipações de recebimento de uma Filial
    Call Chama_Tela("BaixasRecLista", colSelecao, Nothing, Nothing, "NumMovCta = ?")

    Exit Sub

Erro_BotaoBaixas_Click:

    Select Case Err

        Case 15474
            
        Case 15451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANTECIPRECEB_NAO_CARREGADO", Err)

        Case 15455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142882)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro  As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objAntecipReceb As New ClassAntecipReceb
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim lPosicaoSeparador As Long

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Cliente foi informado
    If Len(Trim(Cliente.Text)) = 0 Then Error 15450

    'Lê o Código do Cliente que está na tela
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then Error 15453

    'Passa o Código do Cliente que está na tela para o Obj
    objAntecipReceb.lCliente = objCliente.lCodigo

    'Verifica se a Filial foi informada
    If Len(Trim(Filial.Text)) = 0 Then Error 15454

    'Passa o Código da Filial que está na tela para o Obj
    objAntecipReceb.iFilial = Codigo_Extrai(Filial.Text)

    'Verifica se a Conta corrente foi informada
    If Len(Trim(CodConta.Text)) = 0 Then Error 15455

    'Passa o Código da Conta que está na tela para o Obj
    objAntecipReceb.iCodConta = Codigo_Extrai(CodConta.Text)

    'Verifica se o Sequencial foi informado
    If glSequencial = 0 Then Error 15451

    'Passa o Sequencial que está na tela para o Obj
    objAntecipReceb.lSequencial = glSequencial

    'Pede a confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ANTECIPREC", objAntecipReceb.iCodConta)

    'Se confirmar
    If vbMsgRes = vbYes Then

        'Exclui o Recebimento antecipado
        lErro = CF("AntecipRec_Exclui", objAntecipReceb, objContabil)
        If lErro <> SUCESSO Then Error 15452

        'Limpa os campos da tela
        Call Limpa_Tela_AntecipRec

        'Preenche o campo Data com a data corrente do sistema
        Data.Text = Format(gdtDataHoje, "dd/mm/yy")
        
        'Zera iAlterado
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 15450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 15451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANTECIPRECEB_NAO_CARREGADO", Err)

        Case 15452, 15453 'Tratado na rotina chamada

        Case 15454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 15455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142883)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 15456

    'Limpa a tela
    Call Limpa_Tela_AntecipRec

    'Preenche o campo Data com a data corrente do sistema
    Data.Text = Format(gdtDataHoje, "dd/mm/yy")
    
    'Zera iAlterado
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 15456 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142884)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alterações e confirma se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 15457

    'Limpa os campos da tela
    Call Limpa_Tela_AntecipRec

    'Preenche o campo Data com a data corrente do sistema
    Data.Text = Format(gdtDataHoje, "dd/mm/yy")
    
    'Zera iAlterado
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 15457 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142885)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    iClienteAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate
    
    If iClienteAlterado = 0 Then Exit Sub
    
    'Limpa a Combo de Filiais
    Filial.Clear

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Lê os dados do Cliente
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then Error 15458

        'Lê os dados da Filial do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then Error 15459

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
 
    End If
    
    iClienteAlterado = 0
    glSequencial = 0
    LabelSeqMovto.Caption = ""

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case Err

        Case 15458, 15459 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142886)

    End Select

    Exit Sub

End Sub

Private Sub CodConta_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0
    LabelSeqMovto.Caption = ""
    
End Sub

Private Sub CodConta_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_CodConta_Validate

    'Verifica se a Conta está preenchida
    If Len(Trim(CodConta.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o ítem selecionado na ComboBox CodConta
    If CodConta.Text = CodConta.List(CodConta.ListIndex) Then Exit Sub

    'Verifica se o a Conta existe na Combo, e , se existir, seleciona
    lErro = Combo_Seleciona(CodConta, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15461

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then
    
        objContaCorrenteInt.iCodigo = iCodigo
        
        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 15462
    
        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 15463
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43525
        
        End If
        
        'Passa o código da Conta para a tela
        CodConta.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    
    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 15464

    Exit Sub

Erro_CodConta_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 15461, 15462 'Tratados nas Rotinas Chamadas
        
        Case 15463
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            End If
            
       Case 15464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, CodConta.Text)
             
        Case 43525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodConta.Text, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142887)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0
    LabelSeqMovto.Caption = ""

End Sub

Private Sub Data_GotFocus()

Dim lSequencialAux As Long
    
    lSequencialAux = glSequencial
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    glSequencial = lSequencialAux

End Sub

Private Sub Data_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se a Data está vazia
    If Len(Data.ClipText) > 0 Then

        'Verifica se a Data é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 15465

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 15465

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142888)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = 0
    glSequencial = 0
    LabelSeqMovto.Caption = ""

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0
    LabelSeqMovto.Caption = ""

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'Confirmação ao fechar a tela

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoCliente = Nothing
    Set objEventoAntecipRec = Nothing
    Set objEventoCodConta = Nothing
    Set objEventoNatureza = Nothing
    
    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub
    
    'Verifica se é a filial selecionada na Combo
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub
    
    'Verifica se a Filial existe na Combo. Se existir, seleciona
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15466
    
    'Se a Filial(CODIGO) não existe na Combo
    If lErro = 6730 Then

        'Verifica se o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then Error 15467

        'Lê o Código do Cliente que está na tela
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then Error 15626

        'Passa o Código do Cliente que está na tela para o Obj
        objFilialCliente.lCodCliente = objCliente.lCodigo

        'Passa o Código da Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo
        
        'Pesquisa se existe Filial com o Código em questão
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then Error 15468

        'Se não existe Filial com o Código em questão
        If lErro = 12567 Then Error 15469
        
        'Coloca a Filial na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
    End If
    
    'Se a Filial(STRING) não existe na Combo
    If lErro = 6731 Then Error 15470
    
    Exit Sub
    
Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case Err
    
        Case 15466, 15468, 15626 'Tratado na Rotina chamada
                    
        Case 15467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
                    
        Case 15469
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, objCliente.sNomeReduzido)
                    
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If
                            
        Case 15470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE", Err, Filial.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142889)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Historico_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Historico_Validate

    'Se o text não estiver preenchido
    If Len(Trim(Historico.Text)) = 0 Then Exit Sub
    
    'Verifica se o text é maior que o tamanho máximo
    If Len(Trim(Historico.Text)) > 50 Then Error 15471

    'Se o que foi digitado no text é numérico
    If IsNumeric(Trim(Historico.Text)) Then

        'verifica se é inteiro
        lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
        If lErro <> SUCESSO Then Error 40751
        
        'preenche o objeto
        objHistMovCta.iCodigo = CInt(Trim(Historico.Text))
        
        'verifica na tabela de HisMovCta se existe hitorico relacionado com o codigo passado
        lErro = CF("HistMovCta_Le", objHistMovCta)
        If lErro <> SUCESSO And lErro <> 15011 Then Error 40752
    
        'se não existir ----> Error
        If lErro = 15011 Then Error 40753
                
        Historico.Text = objHistMovCta.sDescricao
        
    End If

    Exit Sub

Erro_Historico_Validate:

    Cancel = True


    Select Case Err

        Case 15471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", Err)
        
        Case 40751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, Historico.Text)
        
        Case 40752
        
        Case 40753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142890)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Visible = False
    
    End If
    
    iFrameAtual = 1

    'Inicializa iAlterado
    iAlterado = 0

    Set objEventoCliente = New AdmEvento
    Set objEventoAntecipRec = New AdmEvento
    Set objEventoCodConta = New AdmEvento
    Set objEventoNatureza = New AdmEvento

    'Carrega a Combo Box CodConta
    lErro = Carrega_CodConta()
    If lErro <> SUCESSO Then Error 15477

    'Carrega a Como Box TipoMeioRecto
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then Error 15478

    'Carrega a Combo Box Historico
    lErro = Carrega_Historico()
    If lErro <> SUCESSO Then Error 15479

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 39552
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then Error 39552
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 16965, 15477, 15478, 15479, 39552

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142891)

    End Select
    
    iAlterado = 0
    
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
        If lErro <> SUCESSO Then Error 39554

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39555

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
        
        Case 39554 'Tratado nas rotinas chamadas
        
        Case 39555
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub LabelCliente_Click()
'Chamada do Browse de Cliente

Dim colSelecao As Collection
Dim objCliente As New ClassCliente
Dim objAntecipReceb As New ClassAntecipReceb
Dim lErro As Long
Dim iCodFilial As Integer

On Error GoTo Erro_LabelCliente_Click

    'Se o Cliente não está preenchido
    If Len(Trim(Cliente.Text)) = 0 Then

        objAntecipReceb.lCliente = 0

    'Se o Cliente está preenchido
    Else

        'Lê o Código do Cliente que está na tela
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then Error 15472

        'Passa o Código do Cliente que está na tela para o Obj
        objAntecipReceb.lCliente = objCliente.lCodigo

    End If

    'Chama a tela com a lista de Clientees
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

    Exit Sub

Erro_LabelCliente_Click:

    Select Case Err

        Case 15472
            Cliente.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142892)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodConta_Click()
'Chamada do Browse de Contas

Dim colSelecao As Collection
Dim objConta As New ClassContasCorrentesInternas
Dim objAntecipReceb As New ClassAntecipReceb

    If Len(Trim(CodConta.Text)) = 0 Then

        objAntecipReceb.iCodConta = 0

    Else

        objAntecipReceb.iCodConta = Codigo_Extrai(CodConta.Text)
        
        'Passa o Código da Conta que está na tela para o Obj
        objConta.iCodigo = objAntecipReceb.iCodConta

    End If

    'Chama a tela com a lista de Contas
    Call Chama_Tela("CtaCorrenteLista", colSelecao, objConta, objEventoCodConta)

    Exit Sub

End Sub

Private Sub LabelSequencial_Click()
    Call Botao_AntecipReceb_Click
End Sub

Private Sub NumRefExterna_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodConta_evSelecao(obj1 As Object)

Dim objConta As ClassContasCorrentesInternas
Dim bCancel As Boolean

    Set objConta = obj1
    
    CodConta.Text = CStr(objConta.iCodigo)
    CodConta_Validate (bCancel)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then
        
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub
        
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_ADIANTAM_CLIENTE_ID
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ADIANTAM_CLIENTE_CONTABILIZACAO
                        
        End Select
    
    End If

End Sub

Function Trata_Parametros(Optional objAntecipReceb As ClassAntecipReceb) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objAntecipReceb estiver preenchido
    If Not (objAntecipReceb Is Nothing) Then

        'Carrega na tela os dados relativos à Antecipação de Recebimento
        lErro = Traz_AntecipRec_Tela(objAntecipReceb)
        If lErro <> SUCESSO Then Error 15473

    Else

        'Preenche o campo Data com a data corrente do sistema
        Data.Text = Format(gdtDataHoje, "dd/mm/yy")
        
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 15473 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142893)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Traz_AntecipRec_Tela(objAntecipReceb As ClassAntecipReceb) As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim iIndice As Integer, bCancel As Boolean
Dim sNaturezaEnxuta As String

On Error GoTo Erro_Traz_AntecipRec_Tela

    lErro = CF("AntecipRec_Movto_Le", objAntecipReceb)
    If lErro <> SUCESSO Then Error 15474

    'Coloca o Cliente na tela
    Cliente.Text = CStr(objAntecipReceb.lCliente)
        
    'Carrega as Filiais do Cliente
    Call Cliente_Validate(bCancel)
    
    objFilialCliente.lCodCliente = objAntecipReceb.lCliente
    objFilialCliente.iCodFilial = objAntecipReceb.iFilial
    
    'Pesquisa se existe Filial com o código em questão
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then Error 40754

    'Se não existe Filial com o Código em questão
    If lErro = 12567 Then Error 40755
        
    'Coloca a Filial na tela
    Filial.Text = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        
    CodConta.Text = CStr(objAntecipReceb.iCodConta)
    Call CodConta_Validate(bCancel)
  
    'Coloca a Data na tela
    Data.Text = Format(objAntecipReceb.dtData, "dd/MM/yy")
        
    'Coloca o Valor na tela
    Valor.Text = CStr(objAntecipReceb.dValor)
    
    'Verifica se o TiPoMeioPago existe
    objTipoMeioPagto.iTipo = objAntecipReceb.iTipoMeioPagto

    lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
    If lErro <> SUCESSO And lErro <> 11909 Then Error 40758

    If lErro = 11909 Then Error 40759

    TipoMeioPagto.Text = CStr(objAntecipReceb.iTipoMeioPagto) & SEPARADOR & objTipoMeioPagto.sDescricao
        
    'Coloca o NumRefExterna na tela
    NumRefExterna.PromptInclude = False
    NumRefExterna.Text = objAntecipReceb.sNumRefExterna
    NumRefExterna.PromptInclude = True
    
    'Coloca o Historico na tela
    Historico.Text = objAntecipReceb.sHistorico
    
    'Coloca o Saldo na tela
    Saldo.Caption = Format(objAntecipReceb.dSaldoNaoApropriado, "Standard")

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objAntecipReceb.lNumMovto)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 39553
    
    glSequencial = objAntecipReceb.lSequencial
    If glSequencial = 0 Then
        LabelSeqMovto.Caption = ""
    Else
        LabelSeqMovto.Caption = CStr(glSequencial)
    End If
    
    If Len(Trim(objAntecipReceb.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objAntecipReceb.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then Error 39553
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
    Else
    
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
        
    End If
    
    Call Natureza_Validate(bSGECancelDummy)
    
    Traz_AntecipRec_Tela = SUCESSO

    Exit Function

Erro_Traz_AntecipRec_Tela:

    Traz_AntecipRec_Tela = Err

    Select Case Err

        Case 15474, 39553, 40754, 40756, 40758 'Tratados nas rotinas chamadas
        
        Case 40757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objAntecipReceb.iCodConta)
        
        Case 40759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objAntecipReceb.iTipoMeioPagto)
        
        Case 40755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142894)

    End Select

    Exit Function

End Function

Private Sub Sequencial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoMeioPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    'Verifica se o TipoMeioPagto está preenchido
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchido com o ítem selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub

    'Tenta selecionar o TipoMeioPagto com o código digitado
    lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15600

    'Se o TipoMeioPagto já existe na Combo
    If lErro = 6730 Then

        objTipoMeioPagto.iTipo = iCodigo
    
        'Pesquisa no BD a existência do Tipo de pagamento passado por parâmetro
        lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
        If lErro <> SUCESSO And lErro <> 11909 Then Error 15601
    
        'Se não existir
        If lErro = 11909 Then Error 15602
        
        'Coloca o Tipo de Pagamento na tela
        TipoMeioPagto.Text = CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
    
    End If
    
    'Se o Tipo de pagamento não existe na Combo
    If lErro = 6731 Then Error 15603
    
    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True


    Select Case Err

        Case 15600, 15601
    
        Case 15602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)
            
        Case 15603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE1", Err, TipoMeioPagto.Text)
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142895)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text
                
        'Diminui data
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 33541
        
        Data.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_DownClick:
    
    Select Case Err
    
        Case 33541 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142896)
        
    End Select
    
    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text
        
        'Aumenta a data
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 33542
        
        Data.Text = sData
    
    End If
    
    Exit Sub
    
Erro_UpDown1_UpClick:
    
    Select Case Err
    
        Case 33542 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142897)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, bCancel As Boolean

    Set objCliente = obj1
    
    Cliente.Text = CStr(objCliente.lCodigo)
    Call Cliente_Validate(bCancel)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Se Valor está preenchido
    If Len(Trim(Valor.Text)) > 0 Then

        'Verifica se Valor é válido
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 15488

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 15488

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142898)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objAntecipReceb As ClassAntecipReceb) As Long
'Passa os dados do Recebimento Antecipado que estão na tela para o Obj

Dim iCodFilial As Integer
Dim lErro As Long
Dim objCliente As New ClassCliente
Dim lPosicaoSeparador As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Lê o Código do Cliente que está na tela
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then Error 15489

    'Passa o Código do Cliente que está na tela para o Obj
    objAntecipReceb.lCliente = objCliente.lCodigo

    'Passa o Código da Filial que está na tela para o Obj
    objAntecipReceb.iFilial = Codigo_Extrai(Filial.Text)

    'Passa o Código da Conta Corrente que está na tela para o Obj
    objAntecipReceb.iCodConta = Codigo_Extrai(CodConta.Text)

    'Passa o Sequencial que está na tela para o Obj
    objAntecipReceb.lSequencial = glSequencial

    'Passa o Valor que está na tela para o Obj
    objAntecipReceb.dValor = CDbl(Valor.Text)

    'Passa o Tipo de Pagamento que está na tela para o Obj
    objAntecipReceb.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)

    'Passa o Número que está na tela para o Obj
    If Len(Trim(NumRefExterna.Text)) > 0 Then objAntecipReceb.sNumRefExterna = NumRefExterna.Text

    'Passa o Saldo que está na tela para o Obj
    objAntecipReceb.dSaldoNaoApropriado = CDbl(Trim(Saldo.Caption))

    'Passa a Data que está na tela para o Obj
    objAntecipReceb.dtData = CDate(Data.Text)

    'Passa o Histórico que está na tela para o Obj
    objAntecipReceb.sHistorico = Historico.Text

    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then Error 15489
    
    objAntecipReceb.sNatureza = sNaturezaFormatada
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 15489 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142899)

    End Select

    Exit Function

End Function

Private Sub objEventoAntecipRec_evSelecao(obj1 As Object)
'Evento referente ao Browse de Recebimento antecipado exibido no

Dim objAntecipReceb As ClassAntecipReceb
Dim lErro As Long

On Error GoTo Erro_objEventoAntecipRec_evSelecao

    Set objAntecipReceb = obj1

    'Coloca na tela os dados do Recebimento antecipado passado pelo Obj
    lErro = Traz_AntecipRec_Tela(objAntecipReceb)
    If lErro <> SUCESSO Then Error 15490

    glSequencial = objAntecipReceb.lSequencial
    If glSequencial = 0 Then
        LabelSeqMovto.Caption = ""
    Else
        LabelSeqMovto.Caption = CStr(glSequencial)
    End If
    
    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoAntecipRec_evSelecao:

    Select Case Err

        Case 15490 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142900)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValor As Double
Dim objAntecipReceb As New ClassAntecipReceb
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim lPosicaoSeparador As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) = 0 Then Error 15491

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) = 0 Then Error 15492

    'Verifica se a Conta está preenchida
    If Len(Trim(CodConta.Text)) = 0 Then Error 15493

    'Verifica se a Data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then Error 15495

    'Verifica se o Valor está preenchido
    If Len(Trim(Valor.Text)) = 0 Then Error 15496

    'Verifica se o Tipo de Pagamento está preenchido
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 15497

    objTipoMeioPagto.iTipo = Codigo_Extrai(TipoMeioPagto.Text)

    'Verifica se o Valor é positivo
    dValor = CDbl(Trim(Valor.Text))
        
    'Verifica se Valor é válido
    lErro = Valor_Positivo_Critica(Valor.Text)
    If lErro <> SUCESSO Then Error 15500

    'Move os dados da tela para o Obj
    lErro = Move_Tela_Memoria(objAntecipReceb)
    If lErro <> SUCESSO Then Error 15501

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then Error 20826

    'Grava os dados da Antecipação de recebimento
    lErro = CF("AntecipRec_Grava", objAntecipReceb, objContabil)
    If lErro <> SUCESSO Then Error 15502

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 15491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 15492
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 15493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 15494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", Err)

        Case 15495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 15496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", Err)

        Case 15497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", Err)
        
        Case 15500, 15501, 15502, 20826 'Tratados nas Rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142901)

    End Select

    Exit Function

End Function

Function Limpa_Tela_AntecipRec() As Long

    Call Limpa_Tela(Me)

    Filial.Clear
    CodConta.Text = ""
    TipoMeioPagto.Text = ""
    Historico.Text = ""
    Saldo.Caption = "0,00"
    
    Natureza.PromptInclude = False
    Natureza.Text = ""
    Natureza.PromptInclude = True
    LabelNaturezaDesc.Caption = ""
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    'Zera iAlterado
    iAlterado = 0

End Function

Private Function Carrega_CodConta() As Long
'preenche combo de contas correntes internas

Dim lErro As Long
Dim colCodigoNomeRed As AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_CodConta

    Set colCodigoNomeRed = New AdmColCodigoNome

    'Lê cada Código e Nome reduzido da tabela ContasCorrentesInternas
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 15576

    'Preenche a ComboBox CodConta com os objetos da coleção colCodigoDescricao
    For Each objCodigoNome In colCodigoNomeRed

        CodConta.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodConta.ItemData(CodConta.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_CodConta = SUCESSO

    Exit Function

Erro_Carrega_CodConta:

    Carrega_CodConta = Err

    Select Case Err

        Case 15576 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142902)

    End Select

    Exit Function

End Function

Private Function Carrega_Historico() As Long
'Carrega a combo de Históricos com os históricos da tabela "HistPadraoMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'Lê o Código e a descrição de todos os históricos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then Error 15577

    For Each objCodigoNome In colCodigoNome

        'Insere na Combo de historicos
        Historico.AddItem objCodigoNome.sNome
        Historico.ItemData(Historico.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_Historico = SUCESSO

    Exit Function

Erro_Carrega_Historico:

    Carrega_Historico = Err

    Select Case Err

        Case 15577 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142903)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoMeioPagto() As Long

Dim lErro As Long
Dim colTipoMeioPagto As Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto
Dim colCodigoDescricao As AdmColCodigoNome

On Error GoTo Erro_Carrega_TipoMeioPagto

    Set colTipoMeioPagto = New Collection

    'Lê cada Tipo e Descrição da tabela TipoMeioPagto
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 15578

    'Preenche a ComboBox TipoMeioPagto com os objetos da coleção colTipoMeioPagto
    For Each objTipoMeioPagto In colTipoMeioPagto

        If objTipoMeioPagto.iInativo = TIPOMEIOPAGTO_ATIVO Then

            TipoMeioPagto.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
            TipoMeioPagto.ItemData(TipoMeioPagto.NewIndex) = objTipoMeioPagto.iTipo

        End If

    Next

    Carrega_TipoMeioPagto = SUCESSO

    Exit Function

Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = Err

    Select Case Err

        Case 15578 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142904)

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

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

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
Dim objCliente As New ClassCliente
Dim objFilial As New ClassFilialCliente, objConta As New ClassContasCorrentesInternas
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case CLIENTE_COD
            
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56503
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case CLIENTE_NOME
        
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56504
            
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case FILIAL_COD
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                objMnemonicoValor.colValor.Add objFilial.iCodFilial
            
            Else
                
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Case FILIAL_NOME_RED
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 56505
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 56506
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56507
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 39587
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
        
        Case DATA1
            If Len(Data.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
        
        Case CONTA_COD
            If CodConta.ListIndex <> -1 Then
                objMnemonicoValor.colValor.Add CodConta.ItemData(CodConta.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case CONTA_CONTABIL_CONTA
        
            If CodConta.ListIndex <> -1 Then
            
                objConta.iCodigo = CodConta.ItemData(CodConta.ListIndex)
                lErro = CF("ContaCorrenteInt_Le", objConta.iCodigo, objConta)
                If lErro <> SUCESSO Then Error 56508
                
                If objConta.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objConta.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56509
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                                
        Case VALOR1
            If Len(Valor.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Valor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case FORMA1
            If Len(TipoMeioPagto.Text) > 0 Then
                objMnemonicoValor.colValor.Add TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case HISTORICO1
            If Len(Historico.Text) > 0 Then
                objMnemonicoValor.colValor.Add Historico.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case DOCEXTERNO
            If Len(NumRefExterna.Text) > 0 Then
                objMnemonicoValor.colValor.Add NumRefExterna.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case SEQUENCIAL1
            If Len(LabelSeqMovto.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CLng(LabelSeqMovto.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case Else
            Error 39556

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 56503 To 56509
        
        Case 39556
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142905)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ADIANTAM_CLIENTE_ID
    Set Form_Load_Ocx = Me
    Caption = "Adiantamento de Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AntecipReceb"
    
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is CodConta Then
            Call LabelCodConta_Click
        End If
    
    End If
    
End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub LabelCodConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodConta, Source, X, Y)
End Sub

Private Sub LabelCodConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodConta, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
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

Private Sub Saldo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Saldo, Source, X, Y)
End Sub

Private Sub Saldo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Saldo, Button, Shift, X, Y)
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134026

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134026

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142906)

    End Select
    
    Exit Sub

End Sub


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

Private Function Inicializa_Mascara_Natureza() As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza.Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaEnxuta As String

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
    
    Else

        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833

        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
    
    End If

    Call Natureza_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Natureza_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza.ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_PAGAMENTO)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub


Public Sub LabelNatureza_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Natureza.ClipText
    
    colSelecao.Add NATUREZA_TIPO_PAGAMENTO
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")

End Sub

Public Sub Natureza_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoSaldo_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objMovContaCorrente As New ClassMovContaCorrente
Dim objAntecipRec As New ClassAntecipReceb

On Error GoTo Erro_BotaoSaldo_Click

    'Verifica se a Conta corrente foi informada
    If Len(Trim(CodConta.Text)) = 0 Then Error 15455

    'Verifica se o Sequencial foi informado
    If glSequencial = 0 Then Error 15451
   
    'Passa o Código da Conta que está na tela para o Obj
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodConta.Text)
    objMovContaCorrente.lSequencial = glSequencial

    lErro = CF("MovContaCorrente_ObterNumMovto", objMovContaCorrente)
    If lErro <> SUCESSO Then Error 15474
    
    objAntecipRec.lNumMovto = objMovContaCorrente.lNumMovto
    
    lErro = CF("AntecipRec_Le_NumMovto", objAntecipRec)
    If lErro <> SUCESSO Then Error 15474
    
    'Filtro
    colSelecao.Add objAntecipRec.lNumIntRec

    'Abre o Browse de Antecipações de recebimento de uma Filial
    Call Chama_Tela("RecebAntecipadosMovSaldoLista", colSelecao, Nothing, Nothing, "NumIntRec = ?")

    Exit Sub

Erro_BotaoSaldo_Click:

    Select Case Err

        Case 15474
            
        Case 15451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANTECIPRECEB_NAO_CARREGADO", Err)

        Case 15455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142882)

    End Select

    Exit Sub

End Sub
