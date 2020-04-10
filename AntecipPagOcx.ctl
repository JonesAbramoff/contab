VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl AntecipPagOcx 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   9390
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5280
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   750
      Width           =   9030
      Begin VB.Frame Frame6 
         Caption         =   "Pedido de Compra"
         Height          =   600
         Left            =   225
         TabIndex        =   79
         Top             =   60
         Width           =   8175
         Begin VB.ComboBox ComboFilialPC 
            Height          =   315
            Left            =   5280
            TabIndex        =   80
            Top             =   210
            Width           =   1815
         End
         Begin MSMask.MaskEdBox NumPC 
            Height          =   300
            Left            =   1725
            TabIndex        =   81
            Top             =   210
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   83
            Top             =   270
            Width           =   1350
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
            Left            =   3420
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   270
            Width           =   1800
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Situação Atual"
         Height          =   795
         Left            =   180
         TabIndex        =   42
         Top             =   4335
         Width           =   8190
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
            Left            =   2715
            TabIndex        =   9
            Top             =   135
            Width           =   1575
         End
         Begin VB.CommandButton BotaoBaixas 
            Caption         =   "Baixas de Títulos a Pagar com o Adiantamento"
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
            Left            =   6435
            TabIndex        =   11
            Top             =   135
            Width           =   1695
         End
         Begin VB.CommandButton Botao_AntecipPag 
            Caption         =   "Adiantamentos à Filial do Fornecedor..."
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
            Left            =   4335
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   2070
         End
         Begin VB.Label Saldo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   285
            Left            =   1140
            TabIndex        =   66
            Top             =   330
            Width           =   1530
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
            Left            =   465
            TabIndex        =   67
            Top             =   375
            Width           =   555
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Principais"
         Height          =   1980
         Left            =   225
         TabIndex        =   39
         Top             =   720
         Width           =   8190
         Begin VB.ComboBox CodConta 
            Height          =   315
            Left            =   1770
            TabIndex        =   3
            Top             =   690
            Width           =   1995
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5295
            TabIndex        =   2
            Top             =   270
            Width           =   1935
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   2910
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1140
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
            Top             =   690
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "Standard"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1770
            TabIndex        =   5
            Top             =   1140
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1785
            TabIndex        =   1
            Top             =   270
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
            Left            =   1755
            TabIndex        =   84
            Top             =   1560
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
            Left            =   810
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   86
            Top             =   1590
            Width           =   840
         End
         Begin VB.Label LabelNaturezaDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3000
            TabIndex        =   85
            Top             =   1560
            Width           =   4605
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
            Left            =   1125
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   68
            Top             =   750
            Width           =   570
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
            TabIndex        =   69
            Top             =   1200
            Width           =   480
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
            Left            =   4695
            TabIndex        =   70
            Top             =   750
            Width           =   510
         End
         Begin VB.Label LabelFornecedor 
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
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   71
            Top             =   330
            Width           =   1035
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
            Left            =   4740
            TabIndex        =   72
            Top             =   330
            Width           =   465
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pagamento"
         Height          =   690
         Left            =   195
         TabIndex        =   43
         Top             =   2790
         Width           =   8190
         Begin VB.CommandButton BotaoImprimir 
            Caption         =   "Imprimir Cheque"
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
            Height          =   540
            Left            =   6075
            Picture         =   "AntecipPagOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   135
            Width           =   1590
         End
         Begin VB.ComboBox TipoMeioPagto 
            Height          =   315
            Left            =   1785
            TabIndex        =   6
            Top             =   255
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   300
            Left            =   4605
            TabIndex        =   7
            Top             =   255
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
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
            TabIndex        =   73
            Top             =   315
            Width           =   585
         End
         Begin VB.Label Label11 
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
            Height          =   195
            Left            =   3795
            TabIndex        =   74
            Top             =   315
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Complemento"
         Height          =   810
         Left            =   195
         TabIndex        =   44
         Top             =   3495
         Width           =   8190
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1770
            TabIndex        =   8
            Top             =   285
            Width           =   5370
         End
         Begin VB.Label Label10 
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
            Left            =   840
            TabIndex        =   75
            Top             =   345
            Width           =   825
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   2
      Left            =   150
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4800
         TabIndex        =   78
         Tag             =   "1"
         Top             =   1560
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
         Left            =   7725
         TabIndex        =   17
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   20
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
         Left            =   6270
         TabIndex        =   16
         Top             =   60
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
         Left            =   6270
         TabIndex        =   18
         Top             =   375
         Width           =   2700
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
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2850
         IntegralHeight  =   0   'False
         Left            =   6330
         TabIndex        =   30
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
         TabIndex        =   21
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
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
         Top             =   2190
         Width           =   1770
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   40
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
            TabIndex        =   46
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
            TabIndex        =   47
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   48
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   49
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   41
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         Left            =   15
         TabIndex        =   29
         Top             =   1230
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         Left            =   6330
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
         Left            =   30
         TabIndex        =   50
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   51
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
         TabIndex        =   52
         Top             =   645
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   53
         Top             =   615
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
         TabIndex        =   59
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
         TabIndex        =   60
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3885
         TabIndex        =   61
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7125
      ScaleHeight     =   495
      ScaleWidth      =   2100
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   75
      Width           =   2160
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "AntecipPagOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "AntecipPagOcx.ctx":025C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "AntecipPagOcx.ctx":03E6
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "AntecipPagOcx.ctx":0918
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5745
      Left            =   75
      TabIndex        =   38
      Top             =   420
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   10134
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
      Left            =   4185
      TabIndex        =   76
      Top             =   2445
      Width           =   615
   End
End
Attribute VB_Name = "AntecipPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()


Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoPedCompra As AdmEvento
Attribute objEventoPedCompra.VB_VarHelpID = -1

'mnemonicos
Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
Private Const CONTA_COD As String = "Conta_Codigo"
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta"
Private Const DATA1 As String = "Data"
Private Const VALOR1 As String = "Valor"
Private Const HISTORICO1 As String = "Historico"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const NUMERO1 As String = "Numero"

Dim iFrameAtual As Integer
Public iAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim glSequencial As Long 'sequencial do movto de cta do último adiantamento trazido para a tela

Private WithEvents objEventoAntecipPag As AdmEvento
Attribute objEventoAntecipPag.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCodConta As AdmEvento
Attribute objEventoCodConta.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2
    

Function Limpa_Tela_AntecipPag() As Long

    Call Limpa_Tela(Me)

    Filial.Clear
    CodConta.Text = ""
    TipoMeioPagto.Text = ""
    Historico.Text = ""
    Saldo.Caption = "0,00"
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    'Preenche o campo Data com a data corrente do sistema
    Data.Text = Format(gdtDataHoje, "dd/mm/yy")
    
    Natureza.PromptInclude = False
    Natureza.Text = ""
    Natureza.PromptInclude = True
    LabelNaturezaDesc.Caption = ""
    
    'Zera iAlterado
    iAlterado = 0

End Function

Private Sub Botao_AntecipPag_Click()

Dim objAntecipPag As New ClassAntecipPag
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection
Dim lErro As Long
Dim iCodFilial As Integer

On Error GoTo Erro_BotaoAntecipPag_Click

    'Se Fornecedor não está preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 15338
    
    'Se a Filial não está Preenchida
    If Len(Trim(Filial.Text)) = 0 Then Error 15339

    'Lê os dados do Fornecedor que está na tela
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then Error 15446
    
    'Passa Fornecedor e Filial para o Obj
    colSelecao.Add objFornecedor.lCodigo
    colSelecao.Add Codigo_Extrai(Filial.Text)

    'Abre o Browse de Antecipações de pagamento de uma Filial
    Call Chama_Tela("AntecipPagLista", colSelecao, objAntecipPag, objEventoAntecipPag)

    Exit Sub

Erro_BotaoAntecipPag_Click:

    Select Case Err

        Case 15338
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 15339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 15446 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142831)

    End Select

    Exit Sub

End Sub

Private Sub BotaoBaixar_Click()

Dim lErro  As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objAntecipPag As New ClassAntecipPag
Dim iCodFilial As Integer
Dim objFornecedor As New ClassFornecedor
Dim lPosicaoSeparador As Long

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Fornecedor foi informado
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 15238

    'Lê o Código do Fornecedor que está na tela
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then Error 15431

    'Passa o Código do Fornecedor que está na tela para o Obj
    objAntecipPag.lFornecedor = objFornecedor.lCodigo

    'Verifica se a Filial foi informada
    If Len(Trim(Filial.Text)) = 0 Then Error 15387

    'Passa o Código da Filial que está na tela para o Obj
    lPosicaoSeparador = InStr(Filial.Text, SEPARADOR)
    objAntecipPag.iFilial = CInt(Trim(left(Filial.Text, lPosicaoSeparador - 1)))

    'Verifica se a Conta corrente foi informada
    If Len(Trim(CodConta.Text)) = 0 Then Error 15388

    'Passa o Código da Conta que está na tela para o Obj
    lPosicaoSeparador = InStr(CodConta.Text, SEPARADOR)
    
    If lPosicaoSeparador <> 0 Then
        objAntecipPag.iCodConta = CInt(Trim(left(CodConta.Text, lPosicaoSeparador - 1)))
    Else
        objAntecipPag.iCodConta = CInt(CodConta.Text)
    End If
    
    If glSequencial = 0 Then Error 15961
    
    'Passa o Sequencial que está na tela para o Obj
    objAntecipPag.lSequencial = glSequencial

    'Pede a confirmação da baixa
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_BAIXA_ANTECIPPAG", objAntecipPag.iCodConta)

    'Se confirmar
    If vbMsgRes = vbYes Then

        'Baixa o Pagamento antecipado
        lErro = AntecipPag_Baixa(objAntecipPag)
        If lErro <> SUCESSO Then Error 15240

        'Limpa os campos da tela
        Call Limpa_Tela_AntecipPag

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 15238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 15240, 15431 'Tratados nas rotinas chamadas

        Case 15387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 15388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 15961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANTECIPPAG_NAO_CARREGADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142832)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro  As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objAntecipPag As New ClassAntecipPag
Dim iCodFilial As Integer
Dim objFornecedor As New ClassFornecedor
Dim lPosicaoSeparador As Long

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Fornecedor foi informado
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 15238

    'Lê o Código do Fornecedor que está na tela
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then Error 15431

    'Passa o Código do Fornecedor que está na tela para o Obj
    objAntecipPag.lFornecedor = objFornecedor.lCodigo

    'Verifica se a Filial foi informada
    If Len(Trim(Filial.Text)) = 0 Then Error 15387

    'Passa o Código da Filial que está na tela para o Obj
    lPosicaoSeparador = InStr(Filial.Text, SEPARADOR)
    objAntecipPag.iFilial = CInt(Trim(left(Filial.Text, lPosicaoSeparador - 1)))

    'Verifica se a Conta corrente foi informada
    If Len(Trim(CodConta.Text)) = 0 Then Error 15388

    'Passa o Código da Conta que está na tela para o Obj
    lPosicaoSeparador = InStr(CodConta.Text, SEPARADOR)
    
    If lPosicaoSeparador <> 0 Then
        objAntecipPag.iCodConta = CInt(Trim(left(CodConta.Text, lPosicaoSeparador - 1)))
    Else
        objAntecipPag.iCodConta = CInt(CodConta.Text)
    End If
    
    If glSequencial = 0 Then Error 15961
    
    'Passa o Sequencial que está na tela para o Obj
    objAntecipPag.lSequencial = glSequencial

    'Pede a confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ANTECIPPAG", objAntecipPag.iCodConta)

    'Se confirmar
    If vbMsgRes = vbYes Then

        'Exclui o Pagamento antecipado
        lErro = CF("AntecipPag_Exclui", objAntecipPag, objContabil)
        If lErro <> SUCESSO Then Error 15240

        'Limpa os campos da tela
        Call Limpa_Tela_AntecipPag

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 15238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 15240, 15431 'Tratados nas rotinas chamadas

        Case 15387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 15388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 15961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANTECIPPAG_NAO_CARREGADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142833)

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
    If lErro <> SUCESSO Then Error 15241

    'Limpa a tela
    Call Limpa_Tela_AntecipPag

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 15241 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142834)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim lNumImpressao As Long
Dim sLayoutCheque As String
Dim dtDataEmissao As Date
Dim iCodigo As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objBanco As New ClassBanco
Dim objInfoChequePag As New ClassInfoChequePag

On Error GoTo Erro_BotaoImprimir_Click
        
    'Extrai o código da combo CodConta e guarda na variável iCodigo
    iCodigo = Codigo_Extrai(CodConta.Text)
       
    'Se valor estiver vazio então dispara erro
    If Len(Trim(Valor.Text)) = 0 Then gError 80418
       
    'Se iCodigo estiver vazio então dispara erro
    If iCodigo = 0 Then gError 80413
    
    'Se Fornecedor estiver vazio então dispara erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 87004
    
    'Le a Conta Corrente a partir de iCodigo passado como parâmetro
    lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 80411

    'Caso a Conta Corrente não tiver sido encontrada dispara erro
    If lErro = 11807 Then gError 80414

    'Caso a Conta Corrente não for bancária dispara erro
    If objContaCorrenteInt.iCodBanco = 0 Then gError 80417
    
    'Atribui o valor retornado de objContaCorrenteInt.iCodBanco a objBanco.iCodBanco
    objBanco.iCodBanco = objContaCorrenteInt.iCodBanco
    
    'Le o Banco a partir de objBanco.iCodBanco
    lErro = CF("Banco_Le", objBanco)
    If lErro <> SUCESSO And lErro <> 16091 Then gError 80412
        
    'Caso o banco não tiver sido encontrado dispara erro
    If lErro = 16091 Then gError 80415
        
    'Atribui retorno de objBanco.sLayoutCheque a variavel sLayoutCheque
    sLayoutCheque = objBanco.sLayoutCheque
                                
    'Recolhe os dados do cheque da tela para o tPreparaImpCheque
    Call Move_tela_Cheque(objInfoChequePag, dtDataEmissao)
                                
    'Chama a função que prepara a impressão do cheque
    lErro = CF("PreparaImpressao_Cheque", lNumImpressao, objInfoChequePag)
    If lErro <> SUCESSO Then gError 80405

    'Chama a função responsável pela impressão do cheque
    lErro = ImprimirCheques(lNumImpressao, sLayoutCheque, dtDataEmissao)
    If lErro <> SUCESSO Then gError 87005
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 80405, 80411, 80412, 87005
                                    
        Case 80413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)
        
        Case 80414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", gErr, CodConta.Text)
        
        Case 80415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, objBanco.iCodBanco)
            
        Case 80417
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", gErr)

        Case 80418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)

        Case 87004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142835)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alterações e confirma se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 15242

    'Limpa os campos da tela
    Call Limpa_Tela_AntecipPag

    BotaoImprimir.Enabled = False

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 15242 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142836)

    End Select

    Exit Sub

End Sub

Private Sub CodConta_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

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

    'Verifica se a Conta existe na Combo e se existir, seleciona
    lErro = Combo_Seleciona(CodConta, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15245

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then
    
        objContaCorrenteInt.iCodigo = iCodigo
        
        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 15246
    
        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 15583
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43539

        End If
        
        'Passa o código da Conta para a tela
        CodConta.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    
    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 15584

    Exit Sub

Erro_CodConta_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 15245, 15246 'Tratado na rotina chamada
        
        Case 15583
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            End If
            
        Case 15584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, CodConta.Text)
        
        Case 43539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodConta.Text, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142837)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

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
        If lErro <> SUCESSO Then Error 15248

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 15248

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142838)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'Confirmação ao fechar a tela

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoAntecipPag = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoCodConta = Nothing
    Set objEventoPedCompra = Nothing
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
Dim iCodFilial As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a Filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub
    
    'Verifica se é a Filial selecionada na Combo
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub
    
    'Verifica se a Filial existe na Combo. Se existir, seleciona
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15421
    
    'Se a Filial(CODIGO) não existe na Combo
    If lErro = 6730 Then

        'Verifica se o Fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 15585

        'Lê os dados do Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 15636
        
        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
        objFilialFornecedor.iCodFilial = iCodigo
        
        'Pesquisa se existe Filial com o código em questão
        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then Error 15423

        'Se não existe Filial com o Código em questão
        If lErro = 12929 Then Error 15586
        
        'Coloca a Filial na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
        
    End If
    
    'Se a Filial(STRING) não existe na Combo
    If lErro = 6731 Then Error 15587
    
    Exit Sub
    
Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case Err
    
        Case 15421, 15423 'Tratados nas rotinas chamadas
        
        Case 15586
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FILIALFORNECEDOR_INEXISTENTE", iCodigo)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 15585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 15587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_INEXISTENTE", Err, Filial.Text)
            
        Case 15636 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142839)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 0 Then Exit Sub

    'Limpa a Combo de Filiais
    Filial.Clear

    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Lê os dados do Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 15443

        'Lê os dados da Filial do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then Error 15435

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
               
    End If

    glSequencial = 0
    iFornecedorAlterado = 0
    
    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 15435, 15443 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142840)

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
    If Len(Trim(Historico.Text)) > 50 Then Error 15278

    'Verifica se o que foi digitado é numerico
    If IsNumeric(Trim(Historico.Text)) Then
        
        'verifica se é inteiro
        lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
        If lErro <> SUCESSO Then Error 40744
        
        'preenche o objeto
        objHistMovCta.iCodigo = CInt(Trim(Historico.Text))
        
        'verifica na tabela de HisMovCta se existe hitorico relacionado com o codigo passado
        lErro = CF("HistMovCta_Le", objHistMovCta)
        If lErro <> SUCESSO And lErro <> 15011 Then Error 40745
    
        'se não existir ----> Error
        If lErro = 15011 Then Error 40746
                
        Historico.Text = objHistMovCta.sDescricao
   
    End If

    Exit Sub

Erro_Historico_Validate:

    Cancel = True


    Select Case Err

        Case 15278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", Err)
        
        Case 40744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, Historico.Text)
        
        Case 40745
        
        Case 40746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142841)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim sEspacos As String

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Visible = False
    
    End If
    
    iFrameAtual = 1

    Set objEventoAntecipPag = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCodConta = New AdmEvento
    Set objEventoPedCompra = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    
    'Inicializa iAlterado
    iAlterado = 0

    'Carrega a Combo Box CodConta
    lErro = Carrega_CodConta()
    If lErro <> SUCESSO Then Error 15215

    'Carrega a Como Box TipoMeioPagto
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then Error 15216

    'Carrega a Combo Box Historico
    lErro = Carrega_Historico()
    If lErro <> SUCESSO Then Error 15217
    
    
    'preenche a combo filialPc
    lErro = Carrega_FilialPC()
    If lErro <> SUCESSO Then Error 49529
    
    ComboFilialPC.ListIndex = 0
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_LIGHT Then
        
        ComboFilialPC.left = POSICAO_FORA_TELA
        ComboFilialPC.TabStop = False
        LblFilialPC.left = POSICAO_FORA_TELA
        LblFilialPC.Visible = False
        
    End If
    
    'Verifica se o Módulo de Compras está ativo
    If gcolModulo.Ativo(MODULO_COMPRAS) = MODULO_ATIVO Then
    
        'Coloca MousePointer no label de Pedido de Compra
        LblNumPC.MousePointer = vbArrowQuestion
        
    End If
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then Error 39557
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then Error 39557
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 15215, 15216, 15217, 15444, 39557, 49529 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142842)

    End Select
    
     iAlterado = 0
        
    Exit Sub

End Sub


Private Function Carrega_FilialPC() As Long
'Carrega ComboFilialPc com as Filiais Empresas

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Carrega_FilialPC
        
    'Le Código
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 49528

    For Each objCodigoDescricao In colCodigoDescricao
    
            'coloca na combo
            ComboFilialPC.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
            
            ComboFilialPC.ItemData(ComboFilialPC.NewIndex) = objCodigoDescricao.iCodigo
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

        Case 49528 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142843)

    End Select

    Exit Function

End Function

Private Sub ComboFilialPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboFilialPC_Click()
    
    iAlterado = REGISTRO_ALTERADO

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
    If lErro <> SUCESSO And lErro <> 12250 Then Error 49530
    
    'Se não encontra valor que era CÓDIGO
    If lErro = 12250 Then Error 49531
    
    Exit Sub
    
Erro_ComboFilialPC_Validate:

    Cancel = True


    Select Case Err
    
       Case 49530
       
       Case 49531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA2", Err, Codigo_Extrai(ComboFilialPC.Text))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142844)
    
    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objAntecipPag As ClassAntecipPag) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objAntecipPag estiver preenchido
    If Not (objAntecipPag Is Nothing) Then

        'Carrega na tela os dados relativos à Antecipação de pagamento
        lErro = Traz_AntecipPag_Tela(objAntecipPag)
        If lErro <> SUCESSO Then Error 15218

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

        Case 15218 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142845)

    End Select
    
    iAlterado = 0

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
        If lErro <> SUCESSO Then Error 39559

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39560

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
        
        Case 39559 'Tratado na rotina chamada
        
        Case 39560
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        End Select

    Exit Function

End Function

Private Sub LabelCodConta_Click()
'Chamada do Browse de Contas

Dim colSelecao As Collection
Dim objConta As New ClassContasCorrentesInternas
Dim objAntecipPag As New ClassAntecipPag

    If Len(Trim(CodConta.Text)) = 0 Then

        objAntecipPag.iCodConta = 0

    Else

        objAntecipPag.iCodConta = Codigo_Extrai(CodConta.Text)
        
        'Passa o Código da Conta que está na tela para o Obj
        objConta.iCodigo = objAntecipPag.iCodConta

    End If

    'Chama a tela com a lista de Contas
    Call Chama_Tela("CtaCorrenteLista", colSelecao, objConta, objEventoCodConta)

    Exit Sub

End Sub

Private Sub LabelFornecedor_Click()
'Chamada do Browse de Fornecedores

Dim colSelecao As Collection
Dim objFornecedor As New ClassFornecedor
Dim objAntecipPag As New ClassAntecipPag
Dim lErro As Long
Dim iCodFilial As Integer

On Error GoTo Erro_LabelFornecedor_Click

    If Len(Trim(Fornecedor.Text)) = 0 Then

        objAntecipPag.lFornecedor = 0

    Else

        'Lê o Código do Fornecedor que está na tela
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 15442

        'Passa o Código do Fornecedor que está na tela para o Obj
        objAntecipPag.lFornecedor = objFornecedor.lCodigo

    End If

    'Chama a tela com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

Erro_LabelFornecedor_Click:

    Select Case Err

        Case 15442 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142846)

    End Select

    Exit Sub

End Sub

Private Sub LblNumPC_Click()

Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

    'Verifica se o Módulo de Compras está ativo
    If gcolModulo.Ativo(MODULO_COMPRAS) = MODULO_ATIVO Then

    
        If Len(Trim(NumPC.Text)) > 0 Then
    
            objPedidoCompra.lCodigo = StrParaLong(NumPC.Text)
        
        End If
    
        Call Chama_Tela("PedComprasEnvLista", colSelecao, objPedidoCompra, objEventoPedCompra, "NOT EXISTS (SELECT Excluido FROM PagtosAntecipados WHERE Excluido = 0 AND FilialPedCompra = FilialEmpresa AND NumPedCompra = Codigo)")
    
    End If
    
    Exit Sub
    
End Sub


Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub NumPC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumPC, iAlterado)

End Sub

Private Sub objEventoPedCompra_evSelecao(obj1 As Object)

Dim objPedidoCompra As New ClassPedidoCompras, iIndice As Integer, lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto, objParc As New ClassCondicaoPagtoParc

On Error GoTo Erro_PedCompra_evSelecao

    Set objPedidoCompra = obj1
    
    NumPC.PromptInclude = False
    NumPC.Text = objPedidoCompra.lCodigo
    NumPC.PromptInclude = True
    
    lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 184221
    
    If lErro = SUCESSO Then
    
        'Coloca os dados encontrados na tela
        Fornecedor.Text = CStr(objPedidoCompra.lFornecedor)
        Call Fornecedor_Validate(bSGECancelDummy)
        
        'Coloca a Filial na tela
        Filial.Text = CStr(objPedidoCompra.iFilial)
        Call Filial_Validate(bSGECancelDummy)
        
    End If
    
    For iIndice = 0 To ComboFilialPC.ListCount - 1

        If ComboFilialPC.ItemData(iIndice) = objPedidoCompra.iFilialEmpresa Then

            ComboFilialPC.ListIndex = iIndice
            Exit For

        End If

    Next
    
    Call ComboFilialPC_Validate(bSGECancelDummy)
    
    If objPedidoCompra.iCondicaoPagto = 0 Then
        Valor.Text = Format(objPedidoCompra.dValorTotal, "Standard")
    Else
        objCondicaoPagto.iCodigo = objPedidoCompra.iCondicaoPagto
        lErro = CF("CondicaoPagto_Le_Parcelas", objCondicaoPagto)
        If lErro <> SUCESSO Then gError 184221
        
        If objCondicaoPagto.colParcelas.Count <= 1 Then
            Valor.Text = Format(objPedidoCompra.dValorTotal, "Standard")
        Else
            Set objParc = objCondicaoPagto.colParcelas(1)
            Valor.Text = Format(Arredonda_Moeda(objPedidoCompra.dValorTotal * objParc.dPercReceb), "Standard")
        End If
        
    End If
    
    Me.Show
    
    Exit Sub

Erro_PedCompra_evSelecao:

    Select Case gErr

        Case 184221
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184220)

    End Select

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
                Parent.HelpContextID = IDH_ADIANTAM_FORNEC_IDENT
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ADIANTAM_FORNEC_CONTABILIZACAO
                        
        End Select
    
    End If

End Sub

Private Sub TipoMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoMeioPagto_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call ValidaBotao_Cheque

End Sub

Private Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    'Verifica se o TipoMeioPagto está preenchido
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub

    'Tenta selecionar o TipoMeioPagto com o código digitado
    lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15232

    'Se o TipoMeioPagto já existe na Combo
    If lErro = 6730 Then

        objTipoMeioPagto.iTipo = iCodigo
    
        'Pesquisa no BD a existência do Tipo de pagamento passado por parâmetro
        lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
        If lErro <> SUCESSO And lErro <> 11909 Then Error 15233
    
        'Se não existir
        If lErro = 11909 Then Error 15587
        
        'Coloca o Tipo de Pagamento na tela
        TipoMeioPagto.Text = CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
    
        Call ValidaBotao_Cheque
        
    End If
    
    'Se o Tipo de pagamento não existe na Combo
    If lErro = 6731 Then Error 15588
    
    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True


    Select Case Err

        Case 15232, 15233

        Case 15587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)
            
        Case 15588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE1", Err, TipoMeioPagto.Text)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142847)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    'Se houver dados no campo Data
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        'Diminui a data em 1 dia
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 15219

        'Coloca a data (diminuída) no campo Data
        Data.Text = sData
        
    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case Err

        Case 15219 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142848)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    'Se houver dados no campo Data
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        'Aumenta a data em 1 dia
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 15220

        'Coloca a data (aumentada) no campo Data
        Data.Text = sData
        
    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case Err

        Case 15220 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142849)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1
    
    Fornecedor.Text = CStr(objFornecedor.lCodigo)
    Call Fornecedor_Validate(Cancel)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoAntecipPag_evSelecao(obj1 As Object)
'Evento referente ao Browse de Pagamento antecipado exibido no

Dim objAntecipPag As ClassAntecipPag
Dim lErro As Long

On Error GoTo Erro_objEventoAntecipPag_evSelecao

    Set objAntecipPag = obj1

    'Coloca na tela os dados do Pagamento antecipado passado pelo Obj
    lErro = Traz_AntecipPag_Tela(objAntecipPag)
    If lErro <> SUCESSO Then Error 15390
    
    glSequencial = objAntecipPag.lSequencial

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoAntecipPag_evSelecao:

    Select Case Err

        Case 15390 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142850)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodConta_evSelecao(obj1 As Object)

Dim objConta As ClassContasCorrentesInternas
Dim bCancel As Boolean

    Set objConta = obj1
    
    CodConta.Text = CStr(objConta.iCodigo)
    Call CodConta_Validate(bCancel)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

End Sub

Private Function Traz_AntecipPag_Tela(objAntecipPag As ClassAntecipPag) As Long

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iIndice As Integer, bCancel As Boolean
Dim sNaturezaEnxuta As String

On Error GoTo Erro_Traz_AntecipPag_Tela

    lErro = CF("AntecipPag_Movto_Le", objAntecipPag)
    If lErro <> SUCESSO Then Error 15280
    
    'Coloca os dados encontrados na tela
    Fornecedor.Text = CStr(objAntecipPag.lFornecedor)
    Call Fornecedor_Validate(bCancel)
    
    'Coloca a Filial na tela
    Filial.Text = CStr(objAntecipPag.iFilial)
    Call Filial_Validate(bCancel)
        
    CodConta.Text = CStr(objAntecipPag.iCodConta)
    Call CodConta_Validate(bCancel)
    
    Data.Text = Format(objAntecipPag.dtData, "dd/MM/yy")
    Valor.Text = Format(objAntecipPag.dValor, "Standard")
    
    TipoMeioPagto.Text = CStr(objAntecipPag.iTipoMeioPagto)
    Call TipoMeioPagto_Validate(bSGECancelDummy)
       
    Numero.PromptInclude = False
    Numero.Text = IIf(objAntecipPag.lNumero <> 0, CStr(objAntecipPag.lNumero), "")
    Numero.PromptInclude = True
    
    Historico.Text = objAntecipPag.sHistorico
    
    Saldo.Caption = Format(objAntecipPag.dSaldoNaoApropriado, "Standard")
    
    'Preenche a filialPc e o numPc
    If objAntecipPag.lNumPedCompra <> 0 Then
        
        NumPC.PromptInclude = False
        NumPC.Text = objAntecipPag.lNumPedCompra
        NumPC.PromptInclude = True
        
    Else
    
        NumPC.PromptInclude = False
        NumPC.Text = ""
        NumPC.PromptInclude = True
    
    End If
            
    If Len(Trim(objAntecipPag.iFilialPedCompra)) <> 0 Then
    
        For iIndice = 0 To ComboFilialPC.ListCount - 1

            If ComboFilialPC.ItemData(iIndice) = objAntecipPag.iFilialPedCompra Then

                ComboFilialPC.ListIndex = iIndice
                Exit For

            End If

        Next
        
        Call ComboFilialPC_Validate(bSGECancelDummy)
    End If
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objAntecipPag.lNumMovto)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 39558
    
    glSequencial = objAntecipPag.lSequencial
    
    If Len(Trim(objAntecipPag.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objAntecipPag.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then Error 39558
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
    Else
    
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
        
    End If
    
    Call Natureza_Validate(bSGECancelDummy)
    
    Traz_AntecipPag_Tela = SUCESSO

    Exit Function

Erro_Traz_AntecipPag_Tela:

    Traz_AntecipPag_Tela = Err

    Select Case Err

        Case 15280, 39558 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142851)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objAntecipPag As ClassAntecipPag) As Long

Dim iCodFilial As Integer
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim lPosicaoSeparador As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Passa os dados do Pagamento Antecipado que estão na tela para o Obj

    'Lê o Código do Fornecedor que está na tela
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then Error 15442

    'Passa o Código do Fornecedor que está na tela para o Obj
    objAntecipPag.lFornecedor = objFornecedor.lCodigo

    'Passa o Código da Filial que está na tela para o Obj
    objAntecipPag.iFilial = Codigo_Extrai(Filial.Text)

    'Passa o Código da Conta Corrente que está na tela para o Obj
    objAntecipPag.iCodConta = Codigo_Extrai(CodConta.Text)
    
    'Preenche o obj com Pedido de compra
    If Len(Trim(ComboFilialPC.Text)) <> 0 Then objAntecipPag.iFilialPedCompra = Codigo_Extrai(ComboFilialPC.Text)
    
    If Len(Trim(NumPC.Text)) <> 0 Then objAntecipPag.lNumPedCompra = CLng(NumPC.Text)
    
    'Passa o Sequencial que está na tela para o Obj
    objAntecipPag.lSequencial = glSequencial

    'Passa o Valor que está na tela para o Obj
    objAntecipPag.dValor = CDbl(Valor.Text)

    'Passa o Tipo de Pagamento que está na tela para o Obj
    objAntecipPag.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)

    'Passa o Número que está na tela para o Obj
    If Len(Trim(Numero.Text)) > 0 Then objAntecipPag.lNumero = CLng(Numero.Text)

    'Passa o Saldo que está na tela para o Obj
    objAntecipPag.dSaldoNaoApropriado = StrParaDbl(Saldo.Caption)

    'Passa a Data que está na tela para o Obj
    objAntecipPag.dtData = CDate(Data.Text)

    'Passa o Histórico que está na tela para o Obj
    objAntecipPag.sHistorico = Historico.Text

    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then Error 15442
    
    objAntecipPag.sNatureza = sNaturezaFormatada
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 15442 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142852)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValor As Double
Dim objAntecipPag As New ClassAntecipPag
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim lPosicaoSeparador As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 15221

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 15222

    'Verifica se a Conta está preenchida
    If Len(Trim(CodConta.Text)) = 0 Then gError 15223

    'Verifica se o Valor está preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 15226

    'Verifica se a Data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 15225
    
    'Verifica se o Tipo de Pagamento está preenchido
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then gError 15579

    lPosicaoSeparador = InStr(TipoMeioPagto.Text, SEPARADOR)

    objTipoMeioPagto.iTipo = CInt(left(TipoMeioPagto.Text, lPosicaoSeparador - 1))

    'Verifica se está no BD
    lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
    If lErro <> SUCESSO And lErro <> 11909 Then gError 15580

    If lErro = 11909 Then gError 15608
    
    If objTipoMeioPagto.iExigeNumero = TIPOMEIOPAGTO_EXIGENUMERO Then

        If Len(Trim(Numero.Text)) = 0 Then gError 15581

    End If

    'Move os dados da tela para o Obj
    lErro = Move_Tela_Memoria(objAntecipPag)
    If lErro <> SUCESSO Then gError 15228
    
    'Verifica se é uma alteracao
    If objAntecipPag.lSequencial <> 0 Then
            
         vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_ALTERACAO_ANTECIPPAG", objAntecipPag.iCodConta)
         
         'Se nao => Erro
         If vbMsgRes = vbNo Then gError 95150
            
    End If
    
    If objTipoMeioPagto.iExigeNumero <> TIPOMEIOPAGTO_EXIGENUMERO Then objAntecipPag.lNumero = 0

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 20825

    'Grava os dados da Antecipação de pagamento
    lErro = CF("AntecipPag_Grava", objAntecipPag, objContabil)
    If lErro <> SUCESSO Then gError 15229

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 15221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 15222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 15223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)

        Case 15225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 15226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr, Valor.Text)

        Case 15227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NEGATIVO", gErr)

        Case 15228, 15229, 15580, 20825 'Tratados nas rotinas chamadas

        Case 15579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", gErr)

        Case 15581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INFORMADO", gErr, objTipoMeioPagto.iTipo)

        Case 15608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", gErr, objTipoMeioPagto.iTipo)
            
        Case 95150
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142853)

    End Select

    Exit Function

End Function

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Se Valor está preenchido
    If Len(Trim(Valor.Text)) > 0 Then

        'Verifica se Valor é válido
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 15230

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 15230

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142854)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoMeioPagto() As Long

Dim lErro As Long
Dim colTipoMeioPagto As Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto
Dim colCodigoDescricao As AdmColCodigoNome

On Error GoTo Erro_Carrega_TipoMeioPagto

    Set colTipoMeioPagto = New Collection

    'Lê cada Tipo e Descrição da tabela TipoMeioPagto
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 15235

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

        Case 15235 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142855)

    End Select

    Exit Function

End Function

Private Function Carrega_CodConta() As Long

Dim lErro As Long
Dim colCodigoNomeRed As AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_CodConta

    Set colCodigoNomeRed = New AdmColCodigoNome

   'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 15236

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

        Case 15236 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142856)

    End Select

    Exit Function

End Function

Private Function Carrega_Historico() As Long
'Carrega a combo de Históricos com os históricos da tabela "HistPadraMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'Lê o Código e a descrição de todos os históricos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then Error 15593

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

        Case 15593 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142857)

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
Dim objFornecedor As New ClassFornecedor, sContaTela As String
Dim objFilial As New ClassFilialFornecedor, objConta As New ClassContasCorrentesInternas

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
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39576
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41973
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39577
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
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
                If lErro <> SUCESSO Then Error 41593
                
                If objConta.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objConta.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41974
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case VALOR1
            If Len(Trim(Valor.Text)) > 0 Then
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
        
        Case DATA1
            If Len(Trim(Data.ClipText)) > 0 Then
                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
        
        Case NUMERO1
            If Len(Trim(Numero.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Numero.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case Else
            Error 39561

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 39573, 39574, 39575, 39576, 39577, 41593, 41973, 41974 'Tratados nas Rotinas chamadas
        
        Case 39561
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142858)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ADIANTAM_FORNEC_IDENT
    Set Form_Load_Ocx = Me
    Caption = "Adiantamento à Fornecedor"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AntecipPag"
    
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
            Call LabelFornecedor_Click
        ElseIf Me.ActiveControl Is CodConta Then
            Call LabelCodConta_Click
        End If
    
    End If
    
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

Private Sub Saldo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Saldo, Source, X, Y)
End Sub

Private Sub Saldo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Saldo, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub LabelCodConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodConta, Source, X, Y)
End Sub

Private Sub LabelCodConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodConta, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedor, Source, X, Y)
End Sub

Private Sub LabelFornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedor, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
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

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
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

Public Sub ValidaBotao_Cheque()
'Apenas verifica o conteúdo da combo TipoMeioPagto e
'se a condição for satisfeita habilita o botão Imprimir

Dim iCodigo As Integer

    'Atribui o valor retornado de Codigo_Extrai a variavel iCodigo
    iCodigo = Codigo_Extrai(TipoMeioPagto.Text)
    
    'Verifica se iCodigo é igual a Constante Cheque
    If iCodigo <> Cheque Then
        'Se for diferente desabilita o botão
        BotaoImprimir.Enabled = False
    Else
        'Se for igual habilita o botão
        BotaoImprimir.Enabled = True
    End If
    
    Exit Sub

End Sub


Function Move_tela_Cheque(objInfoChequePag As ClassInfoChequePag, dtDataEmissao As Date) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer

On Error GoTo Erro_Move_tela_Cheque

    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then gError 87928

    'Recolhe os dados do cheque
    objInfoChequePag.sFavorecido = objFornecedor.sRazaoSocial
    objInfoChequePag.dValor = StrParaDbl(Valor.Text)
    objInfoChequePag.lNumRealCheque = StrParaLong(Numero.Text)
    dtDataEmissao = Data.Text
    
    Move_tela_Cheque = SUCESSO
    
    Exit Function
    
Erro_Move_tela_Cheque:

    Move_tela_Cheque = gErr
    
    Select Case gErr
    
        Case 87928
            'Tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142859)
            
    End Select
        
    Exit Function
    
End Function

Function ImprimirCheques(lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long
'chama a impressao de cheques

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String
Dim lErro As Long, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ImprimirCheques

    'a cidade deve vir do endereco da filial que está emitindo, se entrar como EMPRESA_TODA pegar da matriz
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 19466
    
    lErro = objRelatorio.ExecutarDireto("Cheques", "", 0, sLayoutCheques, "NIMPRESSAO", CStr(lNumImpressao), "DEMISSAO", CStr(dtDataEmissao), "TCIDADE", objFilialEmpresa.objEndereco.sCidade)
    If lErro <> SUCESSO Then Error 7431

    ImprimirCheques = SUCESSO

    Exit Function

Erro_ImprimirCheques:

    ImprimirCheques = Err

    Select Case Err

        Case 7431, 19466

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142860)

    End Select

    Exit Function

End Function

'??? transferir p/cprgrava
Function AntecipPag_Baixa(objAntecipPag As ClassAntecipPag) As Long
'Baixa o Pagamento antecipado

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(6) As Long
Dim tMovContaCorrente As typeMovContaCorrente
Dim tAntecipPag As typeAntecipPag
Dim iIndice As Integer
Dim lNumMovtoAux As Integer

On Error GoTo Erro_AntecipPag_Baixa

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 15400

    Next

    'Entra em transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 15399

    'Lê a tabela de Movimentos de Conta corrente
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumMovto, Tipo, TipoMeioPagto, Excluido, DataMovimento, Valor, Conciliado FROM MovimentosContaCorrente WHERE CodConta = ? AND Sequencial = ?", 0, tMovContaCorrente.lNumMovto, tMovContaCorrente.iTipo, tMovContaCorrente.iTipoMeioPagto, tMovContaCorrente.iExcluido, tMovContaCorrente.dtDataMovimento, tMovContaCorrente.dValor, tMovContaCorrente.iConciliado, objAntecipPag.iCodConta, objAntecipPag.lSequencial)
    If lErro <> AD_SQL_SUCESSO Then Error 15404

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 15405

    'Se o Movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 15406

    'Loca o Movimento
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 15407

    'Verifica se o Movimento já foi excluído
    If tMovContaCorrente.iExcluido = MOVCONTACORRENTE_EXCLUIDO Then Error 15408

    'Verifica se o movimento se refere a um Pagamento antecipado
    If tMovContaCorrente.iTipo <> MOVCCI_PAGTO_ANTECIPADO Then Error 15409

    objAntecipPag.dtData = tMovContaCorrente.dtDataMovimento
    objAntecipPag.dValor = tMovContaCorrente.dValor
    objAntecipPag.lNumMovto = tMovContaCorrente.lNumMovto
        
    'Exclui o Pagamento antecipado
    lErro = AntecipPag_Baixa_BD(alComando(), objAntecipPag)
    If lErro <> SUCESSO Then Error 15532

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 15419

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    AntecipPag_Baixa = SUCESSO

    Exit Function

Erro_AntecipPag_Baixa:

    AntecipPag_Baixa = Err

    Select Case Err

        Case 15399
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 15400
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 15404, 15405
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOSCONTACORRENTE1", Err, objAntecipPag.iCodConta, objAntecipPag.lSequencial)

        Case 15406
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVCONTACORRENTE_INEXISTENTE", Err, objAntecipPag.iCodConta, objAntecipPag.lSequencial)

        Case 15407
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOVIMENTOSCONTACORRENTE1", Err, objAntecipPag.iCodConta, objAntecipPag.lSequencial)

        Case 15408
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVCONTACORRENTE_EXCLUIDO", Err, objAntecipPag.iCodConta, objAntecipPag.lSequencial)

        Case 15409
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_ANTECIPPAG", Err)

        Case 15419
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 15532, 20505

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142861)

    End Select

    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'??? transferir p/cprgrava
Private Function AntecipPag_Baixa_BD(alComando() As Long, objAntecipPag As ClassAntecipPag) As Long
'Exclui o Pagamento antecipado do BD

Dim lErro As Long
Dim tAntecipPag As typeAntecipPag
Dim iAno As Integer
Dim iMes As Integer

On Error GoTo Erro_AntecipPag_Baixa_BD

    'Lê a tabela de Pagamentos antecipados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT SaldoNaoApropriado, Fornecedor, Filial_Fornecedor FROM PagtosAntecipados WHERE NumMovto = ?", 0, tAntecipPag.dSaldoNaoApropriado, tAntecipPag.lFornecedor, tAntecipPag.iFilial, objAntecipPag.lNumMovto)
    If lErro <> AD_SQL_SUCESSO Then Error 15411

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 15412

    'Se o Pagamento antecipado não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 15595
    
    'Loca o Pagamento antecipado
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 15413

    'Se o Fornecedor cadastrado for diferente do Fornecedor informado na tela
    If tAntecipPag.lFornecedor <> objAntecipPag.lFornecedor Then Error 15432

    'Se a Filial cadastrada for diferente da Filial informada na tela
    If tAntecipPag.iFilial <> objAntecipPag.iFilial Then Error 15433

    'se já estava "baixado"
    If tAntecipPag.dSaldoNaoApropriado < DELTA_VALORMONETARIO Then Error 15414

    'Atualiza o Registro como baixado na tabela PagtosAntecipados
    lErro = Comando_ExecutarPos(alComando(3), "UPDATE PagtosAntecipados SET SaldoNaoApropriado = 0", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 15418

    AntecipPag_Baixa_BD = SUCESSO

    Exit Function

Erro_AntecipPag_Baixa_BD:

    AntecipPag_Baixa_BD = Err

    Select Case Err

        Case 15411, 15412
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANTECIPPAG1", Err, objAntecipPag.lNumMovto)

        Case 15413
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ANTECIPPAG", Err, objAntecipPag.lNumMovto)

        Case 15414
            Call Rotina_Erro(vbOKOnly, "ERRO_ANTECIPPAG_JA_BAIXADO", Err)

        Case 15418
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ANTECIPPAG", Err, objAntecipPag.lNumMovto)

        Case 15432
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_COINCIDE", Err, objAntecipPag.lFornecedor, tAntecipPag.lFornecedor)

        Case 15433
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_COINCIDE", Err, objAntecipPag.iFilial, tAntecipPag.iFilial)

        Case 15595
            Call Rotina_Erro(vbOKOnly, "ERRO_ANTECIPPAG_INEXISTENTE", Err, objAntecipPag.iCodConta, objAntecipPag.lSequencial)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142862)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Preenche()

Static sNomeReduzidoParte As String
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134049

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134049

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142863)

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
    Call Chama_Tela("BaixasPagLista", colSelecao, Nothing, Nothing, "NumMovCta = ?")

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

Private Sub BotaoSaldo_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objMovContaCorrente As New ClassMovContaCorrente
Dim objAntecipPag As New ClassAntecipPag

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
    
    objAntecipPag.lNumMovto = objMovContaCorrente.lNumMovto

    lErro = CF("AntecipPag_Le_NumMovto", objAntecipPag)
    If lErro <> AD_SQL_SUCESSO And lErro <> 42845 Then gError 95327
    If lErro = 42845 Then gError 95328
    
    'Filtro
    colSelecao.Add objAntecipPag.lNumIntPag

    'Abre o Browse de Antecipações de recebimento de uma Filial
    Call Chama_Tela("PagtoAntecipadosMovSaldoLista", colSelecao, Nothing, Nothing, "NumIntPag = ?")

    Exit Sub

Erro_BotaoSaldo_Click:

    Select Case Err

        Case 15474, 95327
            
        Case 15451
            Call Rotina_Erro(vbOKOnly, "ERRO_ANTECIPRECEB_NAO_CARREGADO", Err)

        Case 15455
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)
            
        Case 95328
            Call Rotina_Erro(vbOKOnly, "ERRO_PAGTO_ANTECIPADO_INEXISTENTE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142882)

    End Select

    Exit Sub

End Sub
