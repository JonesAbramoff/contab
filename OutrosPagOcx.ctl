VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OutrosPagOcx 
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   9390
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4440
      Index           =   3
      Left            =   195
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4800
         TabIndex        =   87
         Tag             =   "1"
         Top             =   1560
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
         Left            =   6270
         TabIndex        =   27
         Top             =   420
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
         Height          =   300
         Left            =   6270
         TabIndex        =   25
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6270
         Style           =   2  'Dropdown List
         TabIndex        =   29
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
         Height          =   300
         Left            =   7710
         TabIndex        =   26
         Top             =   60
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   35
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
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2850
         IntegralHeight  =   0   'False
         Left            =   6360
         TabIndex        =   39
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
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
         TabIndex        =   30
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   37
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   36
         Top             =   2190
         Width           =   1770
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   52
         Top             =   3330
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
            TabIndex        =   64
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
            TabIndex        =   65
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   66
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   67
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   31
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   585
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   24
         Top             =   570
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         Height          =   1770
         Left            =   45
         TabIndex        =   38
         Top             =   1200
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2850
         Left            =   6330
         TabIndex        =   41
         Top             =   1530
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5027
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2850
         Left            =   6330
         TabIndex        =   40
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5027
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
         TabIndex        =   28
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
         Left            =   30
         TabIndex        =   68
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   69
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
         TabIndex        =   70
         Top             =   645
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   71
         Top             =   615
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   72
         Top             =   600
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
         TabIndex        =   73
         Top             =   630
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
         TabIndex        =   74
         Top             =   990
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
         TabIndex        =   75
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
         TabIndex        =   76
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
         TabIndex        =   77
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
         Height          =   345
         Left            =   1980
         TabIndex        =   78
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3885
         TabIndex        =   79
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   80
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
         TabIndex        =   81
         Top             =   600
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
         TabIndex        =   82
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
         TabIndex        =   83
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   735
      Width           =   9090
      Begin VB.Frame Frame2 
         Caption         =   "Cabeçalho"
         Height          =   2685
         Left            =   165
         TabIndex        =   50
         Top             =   60
         Width           =   8730
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   6360
            Picture         =   "OutrosPagOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Numeração Automática"
            Top             =   890
            Width           =   300
         End
         Begin VB.ComboBox Tipo 
            Height          =   315
            Left            =   1590
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   885
            Width           =   2670
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5370
            TabIndex        =   2
            Top             =   300
            Width           =   1815
         End
         Begin VB.ComboBox ComboFilialPC 
            Height          =   315
            Left            =   5370
            TabIndex        =   8
            Top             =   2070
            Width           =   1815
         End
         Begin MSMask.MaskEdBox NumTitulo 
            Height          =   300
            Left            =   5355
            TabIndex        =   4
            Top             =   885
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "999999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1590
            TabIndex        =   1
            Top             =   300
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorTotal 
            Height          =   300
            Left            =   1590
            TabIndex        =   5
            Top             =   1485
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   6450
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1515
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   5370
            TabIndex        =   6
            Top             =   1515
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumPC 
            Height          =   300
            Left            =   1590
            TabIndex        =   7
            Top             =   2070
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelTipo 
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
            Height          =   210
            Left            =   1020
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   930
            Width           =   480
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   990
            TabIndex        =   55
            Top             =   1545
            Width           =   510
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
            Left            =   4710
            TabIndex        =   56
            Top             =   360
            Width           =   525
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
            Left            =   4515
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   57
            Top             =   945
            Width           =   720
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
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   58
            Top             =   360
            Width           =   1035
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
            Left            =   4485
            TabIndex        =   59
            Top             =   1545
            Width           =   765
         End
         Begin VB.Label LblFilialPC 
            AutoSize        =   -1  'True
            Caption         =   "Filial Pedido Compra:"
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
            Left            =   3450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   60
            Top             =   2130
            Width           =   1800
         End
         Begin VB.Label LblNumPC 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Compra:"
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
            Left            =   150
            TabIndex        =   61
            Top             =   2130
            Width           =   1350
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Impostos"
         Height          =   855
         Left            =   150
         TabIndex        =   51
         Top             =   2910
         Width           =   8730
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
            Height          =   285
            Left            =   6660
            TabIndex        =   11
            Top             =   375
            Width           =   930
         End
         Begin MSMask.MaskEdBox ValorINSS 
            Height          =   300
            Left            =   5340
            TabIndex        =   10
            Top             =   375
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorIRRF 
            Height          =   300
            Left            =   1575
            TabIndex        =   9
            Top             =   375
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
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
            Left            =   4755
            TabIndex        =   62
            Top             =   405
            Width           =   495
         End
         Begin VB.Label Label1 
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
            Left            =   1215
            TabIndex        =   63
            Top             =   405
            Width           =   270
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4350
      Index           =   2
      Left            =   210
      TabIndex        =   12
      Top             =   795
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Frame Frame6 
         Caption         =   "Parcelas"
         Height          =   4065
         Left            =   240
         TabIndex        =   53
         Top             =   150
         Width           =   8460
         Begin VB.CheckBox CobrancaAutomatica 
            Caption         =   "Calcula cobrança automaticamente"
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
            Left            =   4890
            TabIndex        =   14
            Top             =   285
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   2445
            TabIndex        =   13
            Top             =   255
            Width           =   2295
         End
         Begin VB.CheckBox Suspenso 
            Caption         =   "Check1"
            Height          =   225
            Left            =   6240
            TabIndex        =   19
            Top             =   615
            Width           =   900
         End
         Begin VB.ComboBox TipoCobranca 
            Height          =   315
            Left            =   4275
            TabIndex        =   18
            Top             =   600
            Width           =   1935
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   225
            Left            =   1575
            TabIndex        =   16
            Top             =   585
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   2835
            TabIndex        =   17
            Top             =   615
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   375
            TabIndex        =   15
            Top             =   585
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1755
            Left            =   195
            TabIndex        =   20
            Top             =   750
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   3096
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
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
            Left            =   195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   84
            Top             =   315
            Width           =   2175
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7140
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "OutrosPagOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "OutrosPagOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "OutrosPagOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OutrosPagOcx.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4875
      Left            =   75
      TabIndex        =   49
      Top             =   390
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   8599
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   2370
      TabIndex        =   85
      Top             =   3510
      Width           =   615
   End
End
Attribute VB_Name = "OutrosPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'VARIAVEIS GLOBAIS
Public iAlterado As Integer
Private iFrameAtual As Integer
Private iFornecedorAlterado As Integer
Private iEmissaoAlterada As Integer
Private iValorTituloAlterado As Integer
Private iValorIRRFAlterado As Integer
Private iValorINSSAlterado As Integer

'GRID
Dim objGridParcela As AdmGrid
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_VenctoReal_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Cobranca_Col As Integer
Dim iGrid_Suspenso_Col As Integer

'BROWSERS
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1
Private WithEvents objEventoFilialPC As AdmEvento
Attribute objEventoFilialPC.VB_VarHelpID = -1

'Início da contabilidade
Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
Private Const NUMERO1 As String = "Numero_Nota_Fiscal"
Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALOR1 As String = "Valor_Total"
Private Const TIPO1 As String = "Tipo_Documento"
Private Const VALOR_IR As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const CONTA_DESP_ESTOQUE As String = "Conta_Desp_Estoque"
'Fim da contabilidade

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Cobranca = 2
Private Const TAB_Contabilizacao = 3

Private Sub CobrancaAutomatica_Click()

    iAlterado = REGISTRO_ALTERADO
    
    'Recalcula as Parcelas no GRID
    Call Recalcula_Cobranca

End Sub

Private Sub ComboFilialPC_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim sFornecedor As String

On Error GoTo Erro_ComboFilialPC_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(ComboFilialPC.Text)) = 0 Then Exit Sub
    
    'Verifica se é uma filial selecionada
    If ComboFilialPC.Text = ComboFilialPC.List(ComboFilialPC.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Seleciona(ComboFilialPC)
    If lErro <> SUCESSO And lErro <> 12250 Then Error 48907
    
    'Se não encontra valor que era CÓDIGO
    If lErro = 12250 Then Error 48908
    
    Exit Sub
    
Erro_ComboFilialPC_Validate:

    Cancel = True


    Select Case Err
    
       Case 48907
       
       Case 48908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA2", Err, ComboFilialPC.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164022)
    
    End Select
    
    Exit Sub

End Sub

Private Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializacao das Variaveis Globais
    iFrameAtual = 1
    iFornecedorAlterado = 0
    iEmissaoAlterada = 0
    iValorTituloAlterado = 0
    iValorIRRFAlterado = 0
    iValorINSSAlterado = 0
    
    'Inicializacao dos Browsers
    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCondPagto = New AdmEvento
    Set objEventoTipo = New AdmEvento
    Set objEventoFilialPC = New AdmEvento
    
    'Carrega na combo as Siglas dos Documentos
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then Error 18792
    
    'Carrega na combo os Tipos de Cobrança
    lErro = CF("Pagamentos_Carrega_ComboTiposCobr", TipoCobranca)
    If lErro <> SUCESSO Then Error 18798
        
    'Carrega na como as Condições de Pagamento
    lErro = CF("Pagamentos_Carrega_ComboCondPag", CondicaoPagamento)
    If lErro <> SUCESSO Then Error 18801
    
    'preenche a combo filialPc
    lErro = Carrega_FilialPC()
    If lErro <> SUCESSO Then Error 48906
    
    Set objGridParcela = New AdmGrid
  
    'Inicializa Grid Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcela)
    If lErro <> SUCESSO Then Error 18802
    
    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then Error 39572
    
    iAlterado = 0
  
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 14229, 18792, 18798, 18801, 18802, 39572, 48906 'Tratados nas Rotinas Chamadas
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164023)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colSigla As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosPag"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then Error 18804

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objTituloPagar.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Fornecedor", objTituloPagar.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objTituloPagar.iFilial, 0, "Filial"
    colCampoValor.Add "NumTitulo", objTituloPagar.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objTituloPagar.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "NumParcelas", objTituloPagar.iNumParcelas, 0, "NumParcelas"
    colCampoValor.Add "ValorTotal", objTituloPagar.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorIRRF", objTituloPagar.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "ValorINSS", objTituloPagar.dValorINSS, 0, "ValorINSS"
    colCampoValor.Add "INSSRetido", objTituloPagar.iINSSRetido, 0, "INSSRetido"
    colCampoValor.Add "SiglaDocumento", objTituloPagar.sSiglaDocumento, STRING_SIGLA_DOCUMENTO, "SiglaDocumento"
    colCampoValor.Add "FilialPedCompra", objTituloPagar.iFilialPedCompra, 0, "FilialPedCompra"
    colCampoValor.Add "NumPedCompra", objTituloPagar.lNumPedCompra, 0, "NumPedCompra"
    colCampoValor.Add "CondicaoPagto", objTituloPagar.iCondicaoPagto, 0, "CondicaoPagto"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    lErro = CF("TiposDocumento_Le_NaoOutrosPag", colSigla)
    If lErro <> SUCESSO Then Error 18803
    
    If colSigla.Count = 0 Then Exit Sub
    
    For iIndice = 1 To colSigla.Count
        colSelecao.Add "SiglaDocumento", OP_DIFERENTE, colSigla(iIndice)
    Next
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 18803, 18804 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164024)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_Tela_Preenche

    objTituloPagar.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objTituloPagar.lNumIntDoc <> 0 Then

        'Carrega objTituloPagar com os dados passados em colCampoValor
        objTituloPagar.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objTituloPagar.iFilial = colCampoValor.Item("Filial").vValor
        objTituloPagar.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objTituloPagar.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objTituloPagar.iNumParcelas = colCampoValor.Item("NumParcelas").vValor
        objTituloPagar.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objTituloPagar.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        objTituloPagar.dValorINSS = colCampoValor.Item("ValorINSS").vValor
        objTituloPagar.iINSSRetido = colCampoValor.Item("INSSRetido").vValor
        objTituloPagar.sSiglaDocumento = colCampoValor.Item("SiglaDocumento").vValor
        objTituloPagar.iFilialPedCompra = colCampoValor.Item("FilialPedCompra").vValor
        objTituloPagar.lNumPedCompra = colCampoValor.Item("NumPedCompra").vValor
        objTituloPagar.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        
        'Coloca na Tela os Dados de objTituloPagar
        lErro = Traz_OutrosPag_Tela(objTituloPagar)
        If lErro <> SUCESSO Then Error 18809

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 18809 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164025)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CondicaoPagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub
    
    'Verifica se é uma Condicaopagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18845
    
    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then
                
        objCondicaoPagto.iCodigo = iCodigo
                
        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 18846
        If lErro = 19205 Then Error 18847
        
        'Testa se pode ser usada em Contas a Pagar
        If objCondicaoPagto.iEmPagamento = 0 Then Error 18848
             
        'Coloca na Tela
        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida
        
        'Preenche o GridParcelas
        Call Recalcula_Cobranca
        
    End If
    
    'Não encontrou o valor que era STRING
    If lErro = 6731 Then Error 18850
    
    Exit Sub
    
Erro_CondicaoPagamento_Validate:
    
    Cancel = True
    
    Select Case Err
    
       Case 18845, 18846, 18849 'Tratados nas Rotinas Chasmadas
        
       Case 18847
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If
        
        Case 18848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", Err, objCondicaoPagto.iCodigo)
        
        Case 18850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagamento.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164026)
    
    End Select
    
    Exit Sub

End Sub

Private Sub DataEmissao_GotFocus()
Dim iEmissaoAux As Integer
    
    iEmissaoAux = iEmissaoAlterada
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    iEmissaoAlterada = iEmissaoAux

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub
    
    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18831
    
    'Se não encontra valor que era CÓDIGO
    If lErro = 6730 Then

        'Verifica se o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 18833

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo
        
        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 18832
        
        'Se não encontrou
        If lErro = 18272 Then
            
            objFornecedor.sNomeReduzido = sFornecedor
            
            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 58596
            
            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            'Sugere que cadastrar nova Filial
            Error 18834
            
        End If
            
        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
        
    End If
    
    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 18835
    
    Exit Sub
    
Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case Err
    
       Case 18831, 18832, 58596 'Tratados nas Rotinas Chamadas
       
       Case 18833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
       Case 18834
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
        
        Case 18835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164027)
    
    End Select
    
    Exit Sub

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub INSSRetido_Click()

Dim lErro As Long

On Error GoTo Erro_INSSRetido_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Recalcula as Parcelas do Grid
    Call Recalcula_Cobranca
    
    Exit Sub
    
Erro_INSSRetido_Click:
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164028)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LblFilialPc_Click()

Dim colSelecao As Collection
Dim objFilialPC As New AdmFiliais

    'Se a ComboFilialPC estiver preenchida, extrai a  sigla
    If Len(Trim(ComboFilialPC.Text)) > 0 Then
        objFilialPC.iCodFilial = ComboFilialPC.Text
    End If

    'Chama a tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialPC, objEventoFilialPC)

End Sub

Private Sub LabelTipo_Click()

Dim colSelecao As Collection
Dim objTipoDocumento As New ClassTipoDocumento

    'Se Tipo estiver preenchida, extrai a  sigla
    If Len(Trim(Tipo.Text)) > 0 Then
        objTipoDocumento.sSigla = Tipo.Text
    End If

    'Chama a tela Tipo
    Call Chama_Tela("TipoDocOutrosPagLista", colSelecao, objTipoDocumento, objEventoTipo)

End Sub

Private Sub NumPC_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumPC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumPC, iAlterado)

End Sub

Private Sub NumTitulo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)
            
End Sub

Private Sub objEventoFilialPC_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFilialPC As AdmFiliais

On Error GoTo Erro_objEventoFilialPC_evSelecao

    Set objFilialPC = obj1
    
    'Preenche a Filial PC
    ComboFilialPC.Text = objFilialPC.iCodFilial
    
    Me.Show
     
    Exit Sub
    
Erro_objEventoFilialPC_evSelecao:

    Select Case Err
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164029)
     
     End Select
     
     Exit Sub
     
End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, bCancel As Boolean

    Set objFornecedor = obj1

    'Preenche campo Fornecedor
    Fornecedor.Text = objFornecedor.sNomeReduzido
        
    'Executa o Validate do Fornecedor
    Call Fornecedor_Validate(bCancel)

    Me.Show
    
    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click

    'Se Fornecedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 18810
    
    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then Error 18811
    
    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Error 18812
    
    'Move os dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then Error 18813
    
    'Adiciona filtros: lFornecedor e iFilial
    colSelecao.Add objTituloPagar.lFornecedor
    colSelecao.Add objTituloPagar.iFilial
    colSelecao.Add objTituloPagar.sSiglaDocumento
    
    'Chama Tela OutrosPagLista
    Call Chama_Tela("OutrosPagLista", colSelecao, objTituloPagar, objEventoNumero)
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 18810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 18811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
    
        Case 18812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)
        
        Case 18813 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164030)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloPagar = obj1
    
    'Traz os dados de objTituloPagar para Teal
    lErro = Traz_OutrosPag_Tela(objTituloPagar)
    If lErro <> SUCESSO Then Error 18814
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show
    
    Exit Sub
    
Erro_objEventoNumero_evSelecao:

    Select Case Err
    
        Case 18814 'Tratado na Rotina Chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164031)
            
    End Select

    Exit Sub

End Sub

Private Sub CondPagtoLabel_Click()

Dim colSelecao As Collection
Dim objCondicaoPagto As New ClassCondicaoPagto

    'Se Condição de Pagto estiver preenchida, extrai o código
    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
    End If

    'Chama a Tela CondicoesPagamentoCPLista
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Preenche campo CondicaoPagamento
    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
    
    'Preenche GridParcelas a partir da Condição de Pagto
    lErro = GridParcelas_Preenche(objCondicaoPagto)
    If lErro <> SUCESSO Then Error 18498
        
    Me.Show
     
    Exit Sub
    
Erro_objEventoCondPagto_evSelecao:

    Select Case Err
    
        Case 18498 'Tratado na Rotina Chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164032)
     
     End Select
     
     Exit Sub
     
End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim objTipoDocumento As New ClassTipoDocumento

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloPagar Is Nothing) Then
        
        'Lê o Título
        lErro = CF("TituloPagar_Le", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18372 Then Error 18818
        
        If lErro <> SUCESSO Then Error 18819
        
        objTipoDocumento.sSigla = objTituloPagar.sSiglaDocumento
        
        'Valida a Sigla passada
        lErro = CF("TipoDocumento_Le", objTipoDocumento)
        If lErro <> SUCESSO Then Error 18820
             
        If objTipoDocumento.iEmNFFatPag <> 1 Then Error 18817
        
        'Traz os dados para a Tela
        lErro = Traz_OutrosPag_Tela(objTituloPagar)
        If lErro <> SUCESSO Then Error 18821
    
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case 18817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_OUTROSPAG", Err, objTituloPagar.lNumTitulo, objTituloPagar.sSiglaDocumento)
    
        Case 18818, 18821, 18820 'Tratados nas Rotinas Chamadas
                           
        Case 18819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOPAGAR_NAO_CADASTRADO", Err, objTituloPagar.lNumIntDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164033)
            
    End Select
    
    iAlterado = 0
    
    Exit Function
   
End Function

Private Function Traz_OutrosPag_Tela(objTituloPagar As ClassTituloPagar) As Long
'Traz os dados do Título passado para a Tela

Dim lErro As Long
Dim colParcelasPag As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim iLinha As Integer
Dim iIndice As Integer, bCancel As Boolean

On Error GoTo Erro_Traz_OutrosPag_Tela
    
    'Limpa a Tela
    Call Limpa_Tela_OutrosPag

    'Coloca os dados do Título Pagar na Tela
    NumTitulo.Text = objTituloPagar.lNumTitulo
    
    'Preenche o Fornecedor e a Filial
    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(bCancel)
    
    Filial.Text = objTituloPagar.iFilial
    Call Filial_Validate(bCancel)
    
    If objTituloPagar.iFilialPedCompra <> 0 Then
        ComboFilialPC.Text = objTituloPagar.iFilialPedCompra
        Call ComboFilialPC_Validate(bSGECancelDummy)
    Else
        ComboFilialPC.Text = ""
    End If
                    
    If objTituloPagar.lNumPedCompra <> 0 Then
        NumPC.PromptInclude = False
        NumPC.Text = objTituloPagar.lNumPedCompra
        NumPC.PromptInclude = True
    Else
        NumPC.PromptInclude = False
        NumPC.Text = ""
        NumPC.PromptInclude = True
    End If
    
    'Preenche a data de Emissao
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
        DataEmissao.PromptInclude = False
        DataEmissao.Text = Format(objTituloPagar.dtDataEmissao, "dd/mm/yy")
        DataEmissao.PromptInclude = True
    End If
    
    'Preenche  os Valores
    ValorTotal.Text = Format(objTituloPagar.dValorTotal, "Standard")
    ValorIRRF.Text = Format(objTituloPagar.dValorIRRF, "Standard")
    ValorINSS.Text = Format(objTituloPagar.dValorINSS, "Standard")
    INSSRetido.Value = objTituloPagar.iINSSRetido
    CondicaoPagamento.Text = ""
    
    'Preenche o Tipo
    Tipo.Text = objTituloPagar.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)
    
    'Preenche a Condicao de Pagamento
    If objTituloPagar.iCondicaoPagto <> 0 Then
    
        CondicaoPagamento.Text = CStr(objTituloPagar.iCondicaoPagto)
        Call CondicaoPagamento_Validate(bCancel)
    
    Else
    
        CondicaoPagamento.Text = ""
    
    End If
    
    'Lê as Parcelas a Pagar vinculadas ao Título
    lErro = CF("ParcelasPagar_Le", objTituloPagar, colParcelasPag)
    If lErro <> SUCESSO Then Error 18878
    
    If colParcelasPag.Count > NUM_MAXIMO_PARCELAS Then Error 18879
    
    'Limpa o Grid
    Call Grid_Limpa(objGridParcela)
    
    iLinha = 0
    
    'Preenche as linhas do Grid Parcelas com os dados de cada Parcela
    For Each objParcelaPagar In colParcelasPag
    
        iLinha = iLinha + 1
        
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objParcelaPagar.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_VenctoReal_Col) = Format(objParcelaPagar.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objParcelaPagar.dValor, "Standard")
        
        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = objParcelaPagar.iTipoCobranca Then
                GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next
        
        If objParcelaPagar.iStatus = STATUS_SUSPENSO Then
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "1"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "0"
        End If
        
    Next

    'Faz o número de linhas existentes do Grid ser igual ao número de Parcelas
    objGridParcela.iLinhasExistentes = iLinha
    
    'Faz refresh nas checkboxes
    Call Grid_Refresh_Checkbox(objGridParcela)
    
    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objTituloPagar.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 39571

    iFornecedorAlterado = 0
    iEmissaoAlterada = 0
    iValorTituloAlterado = 0
    iValorIRRFAlterado = 0
    iValorINSSAlterado = 0
    
    iAlterado = 0
    
    Traz_OutrosPag_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_OutrosPag_Tela:

    Traz_OutrosPag_Tela = Err
    
    Select Case Err
    
        Case 18878, 39571 'Tratados nas Rotinas Chamadas
        
        Case 18879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", Err, colParcelasPag.Count, NUM_MAXIMO_PARCELAS)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164034)
        
    End Select
    
    Exit Function

End Function

Private Sub Fornecedor_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = 1

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then
    
            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 18822
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then Error 18823
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
    
            'CODIGO ESPECÍFICO
            
            lErro = CF("Fornecedor_ObtemPadrao_CondPagto", objFornecedor, objCondicaoPagto)
            If lErro <> SUCESSO Then Error 56792
            
            'Coloca na Tela
            If objCondicaoPagto.iEmPagamento = 1 Then
                CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
                Call CondicaoPagamento_Validate(bCancel)
            End If
    
        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then
            
            'Limpa Combo de Filial
            Filial.Clear
    
        End If

        iFornecedorAlterado = 0
        
    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 18822, 18823, 56792 'Tratados nas Rotinas Chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164035)

    End Select

    Exit Sub

End Sub


Private Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumTitulo.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumTitulo.ClipText)
    If lErro <> SUCESSO Then Error 18836
    
    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True


    Select Case Err

        Case 18836

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164036)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    Tipo.Text = objTipoDocumento.sSigla
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipo_evSelecao:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164037)
     
     End Select
     
     Exit Sub

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoCobranca_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) <> 0 Then

        'Critica se é valor positivo
        lErro = Valor_Positivo_Critica(ValorTotal.Text)
        If lErro <> SUCESSO Then Error 18837
        
        'Põe o valor formatado na tela
        ValorTotal.Text = Format(ValorTotal.Text, "Standard")

    End If
    
    If iValorTituloAlterado = 1 Then
    
        Call Recalcula_Cobranca
        iValorTituloAlterado = 0
        
    End If
    
    Exit Sub

Erro_ValorTotal_Validate:

    Cancel = True


    Select Case Err

        Case 18837

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164038)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 18838

    If iEmissaoAlterada = 1 Then
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case Err

        Case 18838 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164039)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 18839

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 18839

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164040)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 18840

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 18840

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164041)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Tipo_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub
    
    'Verifica se o Tipo foi selecionado
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Tenta Localizar o Tipo no Text da Combo
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then Error 61834

    'Se nao encontrar -> Erro
    If lErro = 60483 Then Error 18842
    
    Exit Sub
    
Erro_Tipo_Validate:

    Cancel = True


    Select Case Err
            
        Case 18842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, Tipo.Text)
        
        Case 61834
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164042)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub CondicaoPagamento_Click()

Dim lErro As Long

On Error GoTo Erro_CondicaoPagamento_Click
   
    iAlterado = REGISTRO_ALTERADO

    Call Recalcula_Cobranca
    
    Exit Sub
    
Erro_CondicaoPagamento_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164043)
           
      End Select
            
    Exit Sub

End Sub


Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim dtDataEmissao As Date
Dim dtDataVenctoReal As Date
Dim iIndice As Integer
Dim dValorIRRF As Double

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o Grid Parcelas
    Call Grid_Limpa(objGridParcela)
        
    lErro = Atualiza_ValorParcelas(objCondicaoPagto)
    If lErro <> SUCESSO Then Error 18851

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
    
        dtDataEmissao = CDate(DataEmissao.Text)
    
        objCondicaoPagto.dtDataEmissao = dtDataEmissao
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then Error 18852
        
        'Número de Parcelas
        objGridParcela.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas
    
        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
        
            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")
            
            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then Error 18853
        
            'Coloca Data de Vencimento Real no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")
        
        Next
                        
    End If
    
    GridParcelas_Preenche = SUCESSO
    
    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = Err
    
    Select Case Err
    
        Case 18851, 18852, 18853 'Tratados nas Rotinas Chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164044)
        
    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39579

        'Verifica se o Grid passado é o GridParcelas
        If objGridInt.objGrid Is GridParcelas Then
        
            'Verifica qual a coluna atual do Grid
            Select Case objGridInt.objGrid.Col
            
                'Critica a Data de Vencimento e gera a Data de Vencto Reral
                Case iGrid_Vencimento_Col
                    lErro = Saida_Celula_Vencimento(objGridInt)
                    If lErro <> SUCESSO Then Error 18854
    
                'Faz a critica do valor da Parcela
                Case iGrid_Valor_Col
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then Error 18855
                    
                'Faz a crítica do Tipo de Cobranca
                Case iGrid_Cobranca_Col
                    lErro = Saida_Celula_Cobranca(objGridInt)
                    If lErro <> SUCESSO Then Error 18856
                           
            End Select
            
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 18857
       
    End If
       
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
     
        Case 18854, 18855, 18856 'Tratados nas Rotinas Chamadas

        Case 18857
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39579
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then
        
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then Error 18858
        
        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Valor_Col) Then CobrancaAutomatica.Value = vbUnchecked
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            For iIndice = 0 To TipoCobranca.ListCount - 1
                If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                    Exit For
                End If
            Next
        
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 18859
                    
    Saida_Celula_Valor = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err
    
    Select Case Err
    
        Case 18858, 18859
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164045)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vencimento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim dtDataVenctoReal As Date
Dim sDataVencimento As String

On Error GoTo Erro_Saida_Celula_Vencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 18860
        
        dtDataVencimento = CDate(DataVencimento.Text)
         
        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            dtDataEmissao = CDate(DataEmissao.Text)
            If dtDataVencimento < DataEmissao Then Error 18861
        End If
        
        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
        
        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then Error 18862
                  
        'Coloca data de Vencimento Real no Grid
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
            For iIndice = 0 To TipoCobranca.ListCount - 1
                If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                    Exit For
                End If
            Next
            
        End If
        
    End If
    
    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 18863

    Saida_Celula_Vencimento = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Vencimento:

    Saida_Celula_Vencimento = Err
    
    Select Case Err
    
        Case 18860, 18862, 18863
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 18861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, DataVencimento.Text, GridParcelas.Row, DataEmissao.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164046)
            
    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_Cobranca(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoCobranca

On Error GoTo Erro_Saida_Celula_Cobranca

    Set objGridInt.objControle = TipoCobranca
    
    'Verifica se o Tipo de Cobrança foi preenchido
    If Len(Trim(TipoCobranca.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoCobranca.Text <> TipoCobranca.List(TipoCobranca.ListIndex) Then
        
            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(TipoCobranca, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18867
            
            If lErro = 6730 Then Error 18865
            
            If lErro = 6731 Then Error 18866
            
        End If
    
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 18864

    Saida_Celula_Cobranca = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Cobranca:

    Saida_Celula_Cobranca = Err
    
    Select Case Err

        Case 18864
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 18865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_CADASTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 18866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", Err, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 18867
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", Err, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164047)
            
    End Select
    
    Exit Function
    
End Function

Private Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIRRF_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIRRF.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
        If lErro <> SUCESSO Then Error 18868
        
        'Põe o valor formatado na tela
        ValorIRRF.Text = Format(ValorIRRF.Text, "Standard")

    End If
    
    If iValorIRRFAlterado <> 0 Then
    
        Call Recalcula_Cobranca
        iValorIRRFAlterado = 0
        
    End If

    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True


    Select Case Err

        Case 18868
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164048)

    End Select

    Exit Sub

End Sub

Private Sub ValorINSS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorINSS_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorINSS.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorINSS.Text)
        If lErro <> SUCESSO Then Error 18869
        
        'Põe o valor formatado na tela
        ValorINSS.Text = Format(ValorINSS.Text, "Standard")
    
    End If
    
    If iValorINSSAlterado <> 0 Then
    
        If INSSRetido.Value = 1 Then Call Recalcula_Cobranca
        iValorINSSAlterado = 0
        
    End If

    Exit Sub

Erro_ValorINSS_Validate:

    Cancel = True


    Select Case Err

        Case 18869

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164049)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 18880
    
    'Limpa a Tela
    Call Limpa_Tela_OutrosPag
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 18880
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 164050)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValorIRRF As Double, dValorINSSRet As Double
Dim dSomaParcelas As Double
Dim iIndice As Integer
Dim dtDataVencimento As Date
Dim objTituloPagar As New ClassTituloPagar
Dim dValorPagar As Double
Dim colParcelas As New colParcelaPagar

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos obrigatórios estão preenchidos
    If Len(Trim(Fornecedor.ClipText)) = 0 Then Error 18881
    If Len(Trim(Filial.Text)) = 0 Then Error 18882
    If Len(Trim(Tipo.Text)) = 0 Then Error 18883
    If Len(Trim(NumTitulo.ClipText)) = 0 Then Error 18884
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Error 18885
    
    'Verifica se GridParcelas foi preenchido
    If objGridParcela.iLinhasExistentes = 0 Then Error 18886
    
    dSomaParcelas = 0
    
    'Loop no GridParcelas
    For iIndice = 1 To objGridParcela.iLinhasExistentes
    
        'Verifica se DataVencimento foi preenchida
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then Error 18887
        
        dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        
        'Verifica se DataVencimento é maior ou igual a DataEmissao
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            If dtDataVencimento < CDate(DataEmissao) Then Error 18888
        End If
        
        'Verifica a ordenação das Datas de Vencimento das Parcelas
        If iIndice > 1 Then
            If dtDataVencimento < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then Error 18889
        End If
            
        'Verifica se Valor da Parcela foi preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))) = 0 Then Error 18890
             
        'Verifica se Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        If lErro <> SUCESSO Then Error 18891
        
        'Acumula Valor Parcela em dSomaParcelas
        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        
    Next
   
    'Valor a Pagar
    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF.Text)
    If Len(Trim(ValorINSS.Text)) > 0 And INSSRetido.Value = vbChecked Then dValorINSSRet = CDbl(ValorINSS.Text)
    
    dValorPagar = CDbl(ValorTotal.Text) - dValorIRRF - dValorINSSRet
    
    'Compara Soma das Parcelas com Valor a Pagar
    If Format(dSomaParcelas, "0.00") <> Format(dValorPagar, "0.00") Then Error 18892
    
    'Move dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then Error 18893
    
    Set objTituloPagar.colParcelaPagar = colParcelas
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
        lErro = objContabil.Contabil_Testa_Data(objTituloPagar.dtDataEmissao)
        If lErro <> SUCESSO Then Error 20833
    End If
    
    'Grava o Titulo no BD (incluindo dados contábeis)
    lErro = CF("OutroPagamento_Grava", objTituloPagar, objContabil)
    If lErro <> SUCESSO Then Error 18894
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
        
Erro_Gravar_Registro:

    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
        
        Case 18881
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 18882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 18883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)
        
        Case 18884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)
            
        Case 18885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", Err)
        
        Case 18886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", Err)
            
        Case 18887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", Err, iIndice)
        
        Case 18888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, dtDataVencimento, iIndice, DataEmissao.Text)
        
        Case 18889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_NAO_ORDENADA", Err)
            
        Case 18890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_NAO_INFORMADO", Err, iIndice)

        Case 18891, 18893, 18894, 20833 'Tratados nas Rotinas Chamadas
         
        Case 18892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_INVALIDA", Err, dSomaParcelas, dValorPagar)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164051)

    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim vbMsgRes As VbMsgBoxResult
Dim colParcelas As New colParcelaPagar

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos identificadores estão preenchidos
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 18933
    If Len(Trim(Filial.Text)) = 0 Then Error 18934
    If Len(Trim(Tipo.Text)) = 0 Then Error 18935
    If Len(Trim(NumTitulo.Text)) = 0 Then Error 18936
    
    'Move dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then Error 18937
    
    'Verifica se campos identificadores correspondem a Título Baixado
    lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then Error 18940
    
    'Se corresponder, erro
    If lErro = SUCESSO Then Error 18941
    
    'Lê Título Pagar a partir dos campos identificadores
    lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18551 Then Error 18942
    
    'Se não encontrou ---> ERRO
    If lErro <> SUCESSO Then Error 18943
    
    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TITULO", objTituloPagar.lNumTitulo)
    
    'Se não confirmar, sai
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui o Título a Pagar
    lErro = CF("TituloPagar_Exclui", objTituloPagar, objContabil)
    If lErro <> SUCESSO Then Error 18944
    
    'Limpa a Tela
    Call Limpa_Tela_OutrosPag
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 18933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
            
        Case 18934
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 18935
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)
        
        Case 18936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)
            
        Case 18937, 18942, 18940, 18944 'Tratados nas Rotinas Chamadas
        
        Case 18941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_BAIXADO_EXCLUSAO", Err, objTituloPagar.lNumTitulo)
            
        Case 18943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOPAGAR_NAO_CADASTRADO1", Err, objTituloPagar.lNumTitulo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164052)
            
    End Select
    
    Exit Sub
            
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 18955
    
    'Limpa a Tela
    Call Limpa_Tela_OutrosPag
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 18955
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164053)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub DataEmissao_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iEmissaoAlterada = 1

End Sub

Private Sub DataVencimento_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimento_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridParcela)

End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcela)

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcela.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcela)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DataVencimentoReal_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboFilialPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboFilialPC_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub


Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
    
    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoTipo = Nothing
    Set objEventoFilialPC = Nothing
    
    Set objGrid1 = Nothing
    Set objGridParcela = Nothing
    Set objContabil = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Function Carrega_TipoDocumento()

Dim lErro As Long
Dim colSigla As New Collection
Dim objTipoDocumento As ClassTipoDocumento
Dim iIndice As Integer, iDefault As Integer

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Outros Pagamentos
    lErro = CF("TiposDocumento_Le_OutrosPag", colSigla)
    If lErro <> SUCESSO Then Error 18793

    iDefault = 0
    
    'Carrega a combo com as Siglas lidas
    For Each objTipoDocumento In colSigla
        iIndice = iIndice + 1
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
        If SIGLA_OUTROS_PAGAMENTOS = objTipoDocumento.sSigla Then iDefault = iIndice - 1
    Next

    Tipo.ListIndex = iDefault
    
    Carrega_TipoDocumento = SUCESSO

    Exit Function
    
Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = Err
    
    Select Case Err
    
        Case 18793
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164054)
            
    End Select
    
    Exit Function

End Function

Private Function Carrega_FilialPC() As Long
'Carrega ComboFilialPc com as Filiais Empresas

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim colCodigo As New Collection

On Error GoTo Erro_Carrega_FilialPC
        
    'Le Código
    lErro = CF("Codigos_Le", "FiliaisEmpresa", "FilialEmpresa", TIPO_INT, colCodigo, 0)
    If lErro <> SUCESSO Then Error 48905

    For iIndice1 = 1 To colCodigo.Count
            
            iCodigo = colCodigo.Item(iIndice1)
            
            'coloca na combo
            ComboFilialPC.AddItem CStr(iCodigo)
            ComboFilialPC.ItemData(ComboFilialPC.NewIndex) = iCodigo
    Next

    'Seleciona a Filial na qual o usuário entrou no Sistema
    If giFilialEmpresa <> EMPRESA_TODA Then

       For iIndice = 0 To ComboFilialPC.ListCount - 1

            If ComboFilialPC.ItemData(iIndice) = giFilialEmpresa Then

                ComboFilialPC.ListIndex = iIndice
                Exit For

            End If

        Next

    Else

        ComboFilialPC.ListIndex = 0

    End If

    Carrega_FilialPC = SUCESSO

    Exit Function

Erro_Carrega_FilialPC:

    Carrega_FilialPC = Err

    Select Case Err

        'Erro já tratado
        Case 48905

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164055)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Suspenso")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoCobranca.Name)
    objGridInt.colCampo.Add (Suspenso.Name)
    
    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_Col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Cobranca_Col = 4
    iGrid_Suspenso_Col = 5
    
    'tela em questão
    Set objGridInt.objForm = Me

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8
    
    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 900
    
    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_Parcelas = SUCESSO
    
    Exit Function
    
End Function

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcela, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcela, iAlterado)
    End If

End Sub

Private Sub GridParcelas_EnterCell()

    Call Grid_Entrada_Celula(objGridParcela, iAlterado)

End Sub

Private Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcela)

End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcela)

End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcela, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcela, iAlterado)
    End If

End Sub
Private Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcela)

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcela)

End Sub

Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcela)

End Sub

Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcela)

End Sub

Private Sub NumTitulo_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Suspenso_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Suspenso_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridParcela)

End Sub

Private Sub Suspenso_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcela)

End Sub

Private Sub Suspenso_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcela.objControle = Suspenso
    lErro = Grid_Campo_Libera_Foco(objGridParcela)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Tipo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoCobranca_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoCobranca_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridParcela)

End Sub

Private Sub TipoCobranca_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcela)

End Sub

Private Sub TipoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcela.objControle = TipoCobranca
    lErro = Grid_Campo_Libera_Foco(objGridParcela)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorINSS_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iValorINSSAlterado = 1

End Sub

Private Sub ValorIRRF_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iValorIRRFAlterado = 1

End Sub

Private Sub ValorParcela_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorParcela_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridParcela)

End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcela)

End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcela.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcela)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Limpa_Tela_OutrosPag()

Dim lErro As Long

    'Chama função que limpa TextBoxes e MaskedEdits da Tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridParcela)
    
    'Limpa os campos não são limpos pela função acima
    Filial.Clear
    ComboFilialPC.ListIndex = -1
    CondicaoPagamento.Text = ""
    Tipo.ListIndex = 1
    INSSRetido.Value = vbUnchecked
    
    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

End Sub

Private Sub ValorTotal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iValorTituloAlterado = 1
    
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
                Parent.HelpContextID = IDH_OUTROS_PAGAMENTOS_ID
                
            Case TAB_Cobranca
                Parent.HelpContextID = IDH_OUTROS_PAGAMENTOS_COBRANCA
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_OUTROS_PAGAMENTOS_CONTABILIZACAO
                        
        End Select
        
    End If

End Sub

Private Function Move_Tela_Memoria(objTituloPagar As ClassTituloPagar, colParcelas As colParcelaPagar) As Long
'Move os dados da Tela para objTituloPagar e colParcelas

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria
    
    'Move Fornecedor
    If Len(Trim(Fornecedor.Text)) > 0 Then
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 18875
        
        If lErro <> SUCESSO Then Error 18876
        
        objTituloPagar.lFornecedor = objFornecedor.lCodigo
        
    End If
    
    'Move Filial do Fornecedor
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    End If
    
    'Move FilialPC
    If Len(Trim(ComboFilialPC.Text)) > 0 Then
        objTituloPagar.iFilialPedCompra = Codigo_Extrai(ComboFilialPC.Text)
    End If
    
    'Move Número PC
    If Len(Trim(NumPC.ClipText)) > 0 Then objTituloPagar.lNumPedCompra = CLng(NumPC.ClipText)
    
    'Move Número do Título
    If Len(Trim(NumTitulo.ClipText)) > 0 Then objTituloPagar.lNumTitulo = CLng(NumTitulo.ClipText)
    
    'Move Data de Emissão
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objTituloPagar.dtDataEmissao = DATA_NULA
    Else
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    End If
      
    'Move outros Valores
    If Len(Trim(ValorTotal.ClipText)) > 0 Then objTituloPagar.dValorTotal = CDbl(ValorTotal.Text)
    objTituloPagar.iNumParcelas = objGridParcela.iLinhasExistentes
    objTituloPagar.dValorICMS = 0
    objTituloPagar.dValorICMSSubst = 0
    objTituloPagar.iCreditoICMS = 0
    objTituloPagar.dValorProdutos = 0
    objTituloPagar.dOutrasDespesas = 0
    objTituloPagar.dValorSeguro = 0
    objTituloPagar.dValorFrete = 0
    objTituloPagar.dValorIPI = 0
    objTituloPagar.dValorINSS = StrParaDbl(ValorINSS.Text)
    objTituloPagar.dValorIRRF = StrParaDbl(ValorIRRF.Text)
    objTituloPagar.iCreditoIPI = 0
    objTituloPagar.iINSSRetido = INSSRetido.Value
    objTituloPagar.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    
    'Move Tipo
    If Len(Trim(Tipo.Text)) > 0 Then objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    
    'Move para colParcelas os dados do Grid Parcelas
    lErro = Move_GridParcelas_Memoria(colParcelas)
    If lErro <> SUCESSO Then Error 18877
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err
    
        Case 18875, 18877 'Tratados nas Rotinas Chamadas
        
        Case 18876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164056)
    
    End Select
    
    Exit Function

End Function

Private Function Move_GridParcelas_Memoria(colParcelas As colParcelaPagar) As Long
'Move para a memória os dados existentes no Grid

Dim lErro As Long
Dim iIndice As Integer
Dim objParcelaPag As ClassParcelaPagar
     
    For iIndice = 1 To objGridParcela.iLinhasExistentes
    
        Set objParcelaPag = New ClassParcelaPagar
        
        'Preenche objParcelaPag com a linha do GridParcelas
        objParcelaPag.iNumParcela = iIndice
        If (GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) <> "" Then objParcelaPag.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        If (GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col)) <> "" Then objParcelaPag.dtDataVencimentoReal = CDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))
        If (GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col)) <> "" Then objParcelaPag.dValor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))) = 0 Then
            'Se estiver vazio usamos o Tipo Cobrança DEFAULT
            objParcelaPag.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
        Else
            objParcelaPag.iTipoCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
        End If
                
        If GridParcelas.TextMatrix(iIndice, iGrid_Suspenso_Col) = "1" Then
            objParcelaPag.iStatus = STATUS_SUSPENSO
        Else
            objParcelaPag.iStatus = STATUS_ABERTO
        End If
        
        'Adiciona objParcelaPag à coleção colParcelas
        With objParcelaPag
            colParcelas.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
        End With
    Next
        
    Move_GridParcelas_Memoria = SUCESSO
    
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

'MNEUMONICOS
Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor, objTipoFornecedor As New ClassTipoFornecedor
Dim objFilial As New ClassFilialFornecedor, sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case FORNECEDOR_COD
            
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 39573
                
                objMnemonicoValor.colValor.Add objFornecedor.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case FORNECEDOR_NOME
        
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 39574
            
                objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
        
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
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39575
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39576
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41972
                
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
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39577
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case NUMERO1
            If Len(Trim(NumTitulo.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NumTitulo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case EMISSAO1
            If Len(Trim(DataEmissao.ClipText)) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case VALOR1
            If Len(Trim(ValorTotal.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case TIPO1
            If Len(Tipo.Text) > 0 Then
                objMnemonicoValor.colValor.Add Tipo.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                    
        Case VALOR_IR
            If Len(Trim(ValorIRRF.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIRRF.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_INSS
            If Len(Trim(ValorINSS.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorINSS.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case INSS_RETIDO
            objMnemonicoValor.colValor.Add INSSRetido.Value
            
        Case CONTA_DESP_ESTOQUE
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 41969
                
                objTipoFornecedor.iCodigo = objFornecedor.iTipo
                lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
                If lErro <> SUCESSO Then Error 41970
                
                If objTipoFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTipoFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then Error 41971
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case Else
            Error 39578
            
    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err
            
        Case 39573, 39574, 39575, 39576, 39577, 41969, 41970, 41971, 41972
        
        Case 39578
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164057)
            
    End Select
        
    Exit Function
    
End Function

Private Function Atualiza_ValorParcelas(objCondicaoPagto As ClassCondicaoPagto) As Long
'Atualiza os valores das Parcelas em funcao de alguma alteracao dos dados da tela

Dim lErro As Long
Dim dValorINSSRetido As Double
Dim dValorIRRF As Double
Dim dValorPagar As Double
Dim colValorParcelas As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_ValorParcelas

    If Len(Trim(ValorTotal.Text)) > 0 Then
    
        dValorINSSRetido = StrParaDbl(IIf(INSSRetido.Value, ValorINSS.Text, 0))
        dValorIRRF = StrParaDbl(ValorIRRF)
        
        dValorPagar = CDbl(ValorTotal) - dValorIRRF - dValorINSSRetido

    End If
    
    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        objCondicaoPagto.dValorTotal = dValorPagar
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then Error 26154

        'Número de Parcelas
        objGridParcela.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcela.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    Atualiza_ValorParcelas = SUCESSO

    Exit Function
    
Erro_Atualiza_ValorParcelas:

    Atualiza_ValorParcelas = Err
    
    Select Case Err
    
        Case 26154 'Tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164058)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Recalcula_Cobranca()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Recalcula_Cobranca

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
    
        'Passa o código da Condição para objCondicaoPagto
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
    
        'Lê Condição a partir do código
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 18842
        If lErro = 19205 Then Error 18843
    
        'Preenche o GridParcelas
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then Error 18844

    End If
    
    Exit Sub
     
Erro_Recalcula_Cobranca:

    Select Case Err
          
        Case 18842, 18844

        Case 18843
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164059)
     
    End Select
     
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_OUTROS_PAGAMENTOS_ID
    Set Form_Load_Ocx = Me
    Caption = "Outros Pagamentos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OutrosPag"
    
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
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call LabelTipo_Click
        ElseIf Me.ActiveControl Is ComboFilialPC Then
            Call LblFilialPc_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is CondicaoPagamento Then
            Call CondPagtoLabel_Click
        End If
    
    End If
    
End Sub


Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub LblFilialPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilialPC, Source, X, Y)
End Sub

Private Sub LblFilialPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilialPC, Button, Shift, X, Y)
End Sub

Private Sub LblNumPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNumPC, Source, X, Y)
End Sub

Private Sub LblNumPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNumPC, Button, Shift, X, Y)
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

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

'##################################
'Inserido por Wagner
Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim sFiltro As String
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoProxNum_Click

    If Len(Trim(Fornecedor.Text)) = 0 Then gError 132165
    If Len(Trim(Filial.Text)) = 0 Then gError 132166
        
    objFornecedor.sNomeReduzido = Fornecedor.Text

    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 132167
    If lErro = 6681 Then gError 132168
    
    sFiltro = "Fornecedor = " & objFornecedor.lCodigo & " AND Filial = " & CStr(Codigo_Extrai(Filial.Text))
    
    lErro = CF("TitulosPag_Automatico", lCodigo, sFiltro)
    If lErro <> SUCESSO Then gError 132151
    
    NumTitulo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 132151, 132167

        Case 132165
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 132166
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 132168
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164060)
    
    End Select

    Exit Sub
    
End Sub
'##################################
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


