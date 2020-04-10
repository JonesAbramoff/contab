VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PedidoServicoOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4830
      Index           =   1
      Left            =   105
      TabIndex        =   40
      Top             =   1050
      Width           =   9285
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1665
         Index           =   0
         Left            =   210
         TabIndex        =   48
         Top             =   255
         Width           =   8895
         Begin VB.CommandButton BotaoDesfazer 
            Caption         =   "Desfazer Baixa"
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
            Height          =   735
            Left            =   7770
            TabIndex        =   243
            Top             =   705
            Width           =   900
         End
         Begin VB.CommandButton BotaoSolicitacaoSRVAP 
            Caption         =   "Trazer Dados Apontados"
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
            Left            =   4500
            TabIndex        =   11
            Top             =   1230
            Width           =   2400
         End
         Begin VB.CommandButton BotaoSolicitacaoSRV 
            Caption         =   "Trazer Dados"
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
            Left            =   3045
            TabIndex        =   10
            Top             =   1230
            Width           =   1350
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   5475
            TabIndex        =   14
            Top             =   1665
            Visible         =   0   'False
            Width           =   2550
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
            Left            =   4005
            TabIndex        =   13
            Top             =   1680
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox FilialFaturamento 
            Height          =   315
            ItemData        =   "PedidoServicoOcx.ctx":0000
            Left            =   5085
            List            =   "PedidoServicoOcx.ctx":0002
            TabIndex        =   8
            Top             =   735
            Width           =   2145
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2940
            Picture         =   "PedidoServicoOcx.ctx":0004
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Numeração Automática"
            Top             =   240
            Width           =   300
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   2115
            TabIndex        =   4
            Top             =   225
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   6135
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   5085
            TabIndex        =   6
            Top             =   225
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodSolicSRV 
            Height          =   300
            Left            =   2115
            TabIndex        =   9
            Top             =   1245
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   2115
            TabIndex        =   12
            Top             =   1695
            Visible         =   0   'False
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label NatOpInternaEspelho 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2115
            TabIndex        =   208
            Top             =   720
            Width           =   525
         End
         Begin VB.Label LblNatOpInternaEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Oper.:"
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
            TabIndex        =   207
            Top             =   780
            Width           =   1650
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   23
            Left            =   6705
            TabIndex        =   206
            Top             =   255
            Width           =   615
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7365
            TabIndex        =   205
            Top             =   225
            Width           =   1305
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
            Height          =   195
            Index           =   41
            Left            =   4860
            TabIndex        =   203
            Top             =   1725
            Visible         =   0   'False
            Width           =   570
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
            Height          =   195
            Left            =   1380
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   202
            Top             =   1740
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label LabelSolicitaSRV 
            AutoSize        =   -1  'True
            Caption         =   "Solicitação de Serviço:"
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
            Left            =   60
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   197
            Top             =   1275
            Width           =   1995
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial Faturamento:"
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
            Index           =   17
            Left            =   3465
            TabIndex        =   51
            Top             =   780
            Width           =   1575
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
            Left            =   1335
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   4275
            TabIndex        =   49
            Top             =   285
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Preços"
         Height          =   795
         Index           =   2
         Left            =   225
         TabIndex        =   44
         Top             =   3180
         Width           =   8865
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   1320
            TabIndex        =   17
            Top             =   285
            Width           =   1875
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   4485
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   285
            Width           =   1815
         End
         Begin MSMask.MaskEdBox PercAcrescFin 
            Height          =   315
            Left            =   7995
            TabIndex        =   19
            Top             =   285
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tabela Preço:"
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
            Left            =   105
            TabIndex        =   47
            Top             =   345
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Acrésc Financ:"
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
            Left            =   6495
            TabIndex        =   46
            Top             =   345
            Width           =   1485
         End
         Begin VB.Label CondPagtoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto:"
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
            Left            =   3390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   45
            Top             =   345
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   795
         Index           =   6
         Left            =   210
         TabIndex        =   41
         Top             =   2160
         Width           =   8865
         Begin VB.CommandButton BotaoContato 
            Caption         =   "Clientes Futuros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7380
            TabIndex        =   204
            Top             =   165
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5130
            TabIndex        =   16
            Top             =   300
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1635
            TabIndex        =   15
            Top             =   315
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   930
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   43
            Top             =   360
            Width           =   660
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
            Left            =   4605
            TabIndex        =   42
            Top             =   345
            Width           =   465
         End
      End
      Begin VB.CheckBox FaturaIntegral 
         Caption         =   "Só libera o pedido integralmente"
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
         Left            =   240
         TabIndex        =   20
         Top             =   4260
         Width           =   3300
      End
      Begin VB.CommandButton BotaoTodosPedidos 
         Caption         =   "Todos os Pedidos de Serviço"
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
         Left            =   6180
         TabIndex        =   21
         Top             =   4215
         Width           =   2865
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4890
      Index           =   2
      Left            =   45
      TabIndex        =   94
      Top             =   1005
      Visible         =   0   'False
      Width           =   9420
      Begin VB.Frame FrameItens 
         BorderStyle     =   0  'None
         Caption         =   "Peças"
         Height          =   2700
         Index           =   2
         Left            =   90
         TabIndex        =   183
         Top             =   390
         Visible         =   0   'False
         Width           =   9210
         Begin MSMask.MaskEdBox PrecoTotalB 
            Height          =   225
            Left            =   2820
            TabIndex        =   237
            Top             =   1305
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox QuantFaturada 
            Height          =   225
            Left            =   5085
            TabIndex        =   194
            Top             =   1890
            Width           =   1500
            _ExtentX        =   2646
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantCancelada 
            Height          =   225
            Left            =   5265
            TabIndex        =   195
            Top             =   1440
            Width           =   1485
            _ExtentX        =   2619
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
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoServicoOcx.ctx":00EE
            Left            =   5370
            List            =   "PedidoServicoOcx.ctx":00F0
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   420
            Width           =   720
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1590
            MaxLength       =   250
            TabIndex        =   185
            Top             =   255
            Width           =   2490
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   300
            TabIndex        =   187
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   6330
            TabIndex        =   189
            Top             =   450
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   2775
            TabIndex        =   190
            Top             =   735
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   4065
            TabIndex        =   191
            Top             =   750
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   1755
            TabIndex        =   192
            Top             =   1170
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox QuantReservadaPV 
            Height          =   225
            Left            =   6480
            TabIndex        =   193
            Top             =   1080
            Width           =   1500
            _ExtentX        =   2646
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   4050
            TabIndex        =   188
            Top             =   435
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2355
            Left            =   75
            TabIndex        =   184
            Top             =   15
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4154
            _Version        =   393216
            Rows            =   21
            Cols            =   1
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Totais"
         Height          =   1290
         Index           =   1
         Left            =   60
         TabIndex        =   210
         Top             =   3180
         Width           =   9285
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   90
            TabIndex        =   96
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   75
            TabIndex        =   211
            Top             =   405
            Visible         =   0   'False
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   2745
            TabIndex        =   98
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   1410
            TabIndex        =   97
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercDescontoItens 
            Height          =   285
            Left            =   4065
            TabIndex        =   99
            ToolTipText     =   "Percentual de desconto dos itens"
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDescontoItens 
            Height          =   285
            Left            =   5400
            TabIndex        =   100
            ToolTipText     =   "Soma dos descontos dos itens"
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ValorProdutos2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   234
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Total"
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
            Index           =   21
            Left            =   8085
            TabIndex        =   233
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "IPI"
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
            Index           =   8
            Left            =   6735
            TabIndex        =   232
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Despesas"
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
            Index           =   7
            Left            =   2790
            TabIndex        =   231
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Seguro"
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
            Index           =   6
            Left            =   1470
            TabIndex        =   230
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Frete"
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
            Index           =   10
            Left            =   105
            TabIndex        =   229
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Base ISS"
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
            Index           =   15
            Left            =   5430
            TabIndex        =   228
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label ISSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   227
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Desconto"
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
            Index           =   4
            Left            =   5430
            TabIndex        =   226
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "% Desconto"
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
            Index           =   3
            Left            =   4125
            TabIndex        =   225
            Top             =   705
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ISS"
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
            Index           =   22
            Left            =   6735
            TabIndex        =   224
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   24
            Left            =   8100
            TabIndex        =   223
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ICMS ST"
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
            Index           =   25
            Left            =   4080
            TabIndex        =   222
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "BC ICMS ST"
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
            Index           =   26
            Left            =   2745
            TabIndex        =   221
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ICMS"
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
            Index           =   27
            Left            =   1470
            TabIndex        =   220
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Base ICMS"
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
            Index           =   28
            Left            =   165
            TabIndex        =   219
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   218
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4065
            TabIndex        =   217
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2745
            TabIndex        =   216
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1410
            TabIndex        =   215
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   90
            TabIndex        =   214
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   213
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   212
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   235
            Top             =   405
            Width           =   1140
         End
      End
      Begin VB.Frame FrameItens 
         BorderStyle     =   0  'None
         Caption         =   "Serviços"
         Height          =   2655
         Index           =   1
         Left            =   120
         TabIndex        =   170
         Top             =   390
         Width           =   9135
         Begin MSMask.MaskEdBox PrecoTotalSRVB 
            Height          =   225
            Left            =   5175
            TabIndex        =   236
            Top             =   1455
            Width           =   1185
            _ExtentX        =   2090
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
         Begin VB.TextBox DescServicoSRV 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2340
            MaxLength       =   250
            TabIndex        =   173
            Top             =   180
            Width           =   2490
         End
         Begin VB.ComboBox UnidadeMedSRV 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoServicoOcx.ctx":00F2
            Left            =   5085
            List            =   "PedidoServicoOcx.ctx":00F4
            Style           =   2  'Dropdown List
            TabIndex        =   172
            Top             =   480
            Width           =   720
         End
         Begin MSMask.MaskEdBox ServicoSRV 
            Height          =   225
            Left            =   2610
            TabIndex        =   174
            Top             =   510
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEntregaSRV 
            Height          =   225
            Left            =   1125
            TabIndex        =   175
            Top             =   165
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
         Begin MSMask.MaskEdBox DescontoSRV 
            Height          =   225
            Left            =   3825
            TabIndex        =   176
            Top             =   495
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
         Begin MSMask.MaskEdBox PercentDescSRV 
            Height          =   225
            Left            =   6270
            TabIndex        =   177
            Top             =   210
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox PrecoUnitarioSRV 
            Height          =   225
            Left            =   1155
            TabIndex        =   178
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox QuantidadeSRV 
            Height          =   225
            Left            =   5925
            TabIndex        =   179
            Top             =   570
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MSMask.MaskEdBox PrecoTotalSRV 
            Height          =   225
            Left            =   7335
            TabIndex        =   180
            Top             =   225
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox QuantFaturadaSRV 
            Height          =   225
            Left            =   2940
            TabIndex        =   181
            Top             =   840
            Width           =   1500
            _ExtentX        =   2646
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantCanceladaSRV 
            Height          =   225
            Left            =   1185
            TabIndex        =   182
            Top             =   795
            Width           =   1485
            _ExtentX        =   2619
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
         Begin MSFlexGridLib.MSFlexGrid GridItensSRV 
            Height          =   2100
            Left            =   15
            TabIndex        =   171
            Top             =   60
            Width           =   9150
            _ExtentX        =   16140
            _ExtentY        =   3704
            _Version        =   393216
            Rows            =   21
            Cols            =   1
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoServicos 
         Caption         =   "Serviços"
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
         Left            =   180
         TabIndex        =   101
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton BotaoPecas 
         Caption         =   "Peças"
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
         Left            =   1650
         TabIndex        =   102
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton BotaoEstoquePeca 
         Caption         =   "Estoque - Peças"
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
         Left            =   7545
         TabIndex        =   103
         Top             =   4530
         Width           =   1785
      End
      Begin VB.CommandButton BotaoGrade 
         Caption         =   "Grade ..."
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
         Left            =   210
         TabIndex        =   95
         Top             =   4530
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComctlLib.TabStrip TabStripSRV 
         Height          =   3090
         Left            =   60
         TabIndex        =   196
         Top             =   60
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   5450
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Serviços"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Peças"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4890
      Index           =   4
      Left            =   30
      TabIndex        =   145
      Top             =   1005
      Visible         =   0   'False
      Width           =   9420
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   4470
         Left            =   75
         TabIndex        =   147
         Top             =   375
         Width           =   9285
         Begin VB.ComboBox TipoPagto 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoServicoOcx.ctx":00F6
            Left            =   3075
            List            =   "PedidoServicoOcx.ctx":0106
            TabIndex        =   199
            Top             =   2280
            Width           =   1965
         End
         Begin VB.CommandButton BotaoTipoPagto 
            Caption         =   "Detalhamento Tipo de Pagto (F5)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7350
            TabIndex        =   198
            Top             =   3960
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.ComboBox TipoDesconto3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   152
            Top             =   1845
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3120
            TabIndex        =   151
            Top             =   1515
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   150
            Top             =   1215
            Width           =   1965
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3075
            Picture         =   "PedidoServicoOcx.ctx":0143
            Style           =   1  'Graphical
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   630
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3075
            Picture         =   "PedidoServicoOcx.ctx":019D
            Style           =   1  'Graphical
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   780
            Width           =   240
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7470
            TabIndex        =   153
            Top             =   1260
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox Desconto3Valor 
            Height          =   225
            Left            =   6105
            TabIndex        =   154
            Top             =   1905
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   155
            Top             =   1905
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
         Begin MSMask.MaskEdBox Desconto2Valor 
            Height          =   225
            Left            =   6135
            TabIndex        =   156
            Top             =   1590
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   157
            Top             =   1590
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
         Begin MSMask.MaskEdBox Desconto1Valor 
            Height          =   225
            Left            =   6120
            TabIndex        =   158
            Top             =   1260
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Ate 
            Height          =   225
            Left            =   4980
            TabIndex        =   159
            Top             =   1260
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
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   570
            TabIndex        =   160
            Top             =   1230
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
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   240
            Left            =   1695
            TabIndex        =   161
            Top             =   1245
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Percentual 
            Height          =   225
            Left            =   7500
            TabIndex        =   162
            Top             =   1605
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox Desconto3Percentual 
            Height          =   225
            Left            =   7455
            TabIndex        =   163
            Top             =   1905
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox DataReferencia 
            Height          =   300
            Left            =   1980
            TabIndex        =   165
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2745
            Left            =   45
            TabIndex        =   166
            Top             =   1005
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   4842
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox ValorDescontoTit 
            Height          =   300
            Left            =   4440
            TabIndex        =   164
            Top             =   210
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   3510
            TabIndex        =   242
            Top             =   240
            Width           =   885
         End
         Begin VB.Label ValorOriginalTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1980
            TabIndex        =   241
            Top             =   210
            Width           =   1185
         End
         Begin VB.Label Label1 
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
            Index           =   12
            Left            =   675
            TabIndex        =   240
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label ValorTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7995
            TabIndex        =   239
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total a Receber:"
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
            Left            =   6465
            TabIndex        =   238
            Top             =   255
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Data de Referência:"
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
            Index           =   9
            Left            =   150
            TabIndex        =   167
            Top             =   675
            Width           =   1740
         End
      End
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
         Left            =   75
         TabIndex        =   146
         Top             =   135
         Value           =   1  'Checked
         Width           =   3360
      End
   End
   Begin VB.CommandButton BotaoInfoAdic 
      Caption         =   "Informações Adicionais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   22
      Top             =   90
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4500
      Index           =   999
      Left            =   105
      TabIndex        =   126
      Top             =   1275
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame3 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   735
         Left            =   105
         TabIndex        =   139
         Top             =   3540
         Width           =   5790
         Begin VB.Label ProdutoDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1140
            TabIndex        =   141
            Top             =   300
            Width           =   4395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            TabIndex        =   140
            Top             =   330
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Reserva dos Produtos"
         Height          =   3450
         Left            =   105
         TabIndex        =   129
         Top             =   45
         Width           =   8940
         Begin VB.TextBox Responsavel 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5790
            TabIndex        =   130
            Top             =   795
            Width           =   2115
         End
         Begin MSMask.MaskEdBox UnidadeMedEst 
            Height          =   225
            Left            =   7965
            TabIndex        =   131
            Top             =   360
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   225
            Left            =   6675
            TabIndex        =   132
            Top             =   345
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
         Begin MSMask.MaskEdBox ProdutoAlmox 
            Height          =   225
            Left            =   1335
            TabIndex        =   133
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almox 
            Height          =   225
            Left            =   2550
            TabIndex        =   134
            Top             =   390
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantReservar 
            Height          =   225
            Left            =   3825
            TabIndex        =   135
            Top             =   390
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox QuantReservada 
            Height          =   225
            Left            =   5250
            TabIndex        =   136
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ItemPedido 
            Height          =   225
            Left            =   765
            TabIndex        =   137
            Top             =   390
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
            Height          =   2805
            Left            =   180
            TabIndex        =   138
            Top             =   285
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4948
            _Version        =   393216
            Rows            =   11
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoReserva 
         Caption         =   "Reserva dos Produtos Pedidos"
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
         Left            =   6105
         TabIndex        =   128
         Top             =   3600
         Width           =   2940
      End
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Libera Reservas do Pedido"
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
         Left            =   6090
         TabIndex        =   127
         Top             =   4005
         Width           =   2940
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4680
      Index           =   7
      Left            =   105
      TabIndex        =   125
      Top             =   1200
      Visible         =   0   'False
      Width           =   9270
      Begin TelasSRV.TabTributacaoFat TabTrib 
         Height          =   4560
         Left            =   135
         TabIndex        =   209
         Top             =   60
         Width           =   9120
         _extentx        =   16087
         _extenty        =   8043
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Private Sub Observacao_Change()Private Sub Observacao_Change()"
      Height          =   4605
      Index           =   6
      Left            =   105
      TabIndex        =   28
      Top             =   1260
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   3630
         Left            =   75
         TabIndex        =   30
         Top             =   90
         Width           =   9120
         Begin VB.ComboBox BLGTipoBloqueio 
            Height          =   315
            ItemData        =   "PedidoServicoOcx.ctx":01F7
            Left            =   555
            List            =   "PedidoServicoOcx.ctx":01F9
            TabIndex        =   33
            Top             =   330
            Width           =   1605
         End
         Begin VB.TextBox BLGObservacao 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3075
            MaxLength       =   250
            TabIndex        =   32
            Top             =   840
            Width           =   4245
         End
         Begin VB.TextBox BLGSeqBloqueio 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1710
            MaxLength       =   250
            TabIndex        =   31
            Top             =   2430
            Width           =   675
         End
         Begin MSMask.MaskEdBox BLGResponsavelLib 
            Height          =   270
            Left            =   7365
            TabIndex        =   34
            Top             =   330
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BLGDataLiberacao 
            Height          =   270
            Left            =   6150
            TabIndex        =   35
            Top             =   330
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BLGUsuario 
            Height          =   270
            Left            =   3315
            TabIndex        =   36
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BLGResponsavel 
            Height          =   270
            Left            =   4785
            TabIndex        =   37
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BLGDataBloqueio 
            Height          =   270
            Left            =   2190
            TabIndex        =   38
            Top             =   345
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid BLGGridBloqueio 
            Height          =   2715
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   4789
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoLiberaBloqueio 
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
         Left            =   6825
         Picture         =   "PedidoServicoOcx.ctx":01FB
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3840
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4605
      Index           =   5
      Left            =   120
      TabIndex        =   104
      Top             =   1200
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CheckBox ComissaoAutomatica 
         Caption         =   "Calcula comissão automaticamente"
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   525
         TabIndex        =   124
         Top             =   135
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   4125
         Index           =   0
         Left            =   60
         TabIndex        =   105
         Top             =   390
         Width           =   9060
         Begin VB.Frame SSFrame4 
            Caption         =   "Totais - Comissões"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   120
            TabIndex        =   108
            Top             =   3195
            Width           =   6975
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Percentual:"
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
               Left            =   2880
               TabIndex        =   114
               Top             =   360
               Width           =   1095
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
               Height          =   255
               Index           =   20
               Left            =   5040
               TabIndex        =   113
               Top             =   360
               Width           =   615
            End
            Begin VB.Label TotalPercentualComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3960
               TabIndex        =   112
               Top             =   345
               Width           =   825
            End
            Begin VB.Label TotalValorComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5640
               TabIndex        =   111
               Top             =   345
               Width           =   1155
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total:"
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
               Left            =   360
               TabIndex        =   110
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label TotalValorBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1440
               TabIndex        =   109
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.CommandButton BotaoVendedores 
            Caption         =   "Vendedores"
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
            Height          =   645
            Left            =   7440
            Picture         =   "PedidoServicoOcx.ctx":27F5
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   3300
            Width           =   1500
         End
         Begin VB.ComboBox DiretoIndireto 
            Height          =   315
            ItemData        =   "PedidoServicoOcx.ctx":2D9F
            Left            =   5070
            List            =   "PedidoServicoOcx.ctx":2DA9
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   915
            Width           =   1335
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   180
            Left            =   435
            TabIndex        =   115
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualBaixa 
            Height          =   180
            Left            =   7020
            TabIndex        =   116
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   180
            Left            =   7875
            TabIndex        =   117
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   180
            Left            =   3825
            TabIndex        =   118
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
            Height          =   180
            Left            =   2700
            TabIndex        =   119
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
            Height          =   180
            Left            =   1815
            TabIndex        =   120
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   180
            Left            =   5880
            TabIndex        =   121
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
            Height          =   180
            Left            =   5025
            TabIndex        =   122
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1950
            Left            =   150
            TabIndex        =   123
            Top             =   360
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   3440
            _Version        =   393216
            Rows            =   11
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Base ICMS Subst"
      Height          =   4605
      Index           =   3
      Left            =   120
      TabIndex        =   52
      Top             =   1260
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame8 
         Caption         =   "Volumes"
         Height          =   735
         Left            =   75
         TabIndex        =   85
         Top             =   2040
         Width           =   9090
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7140
            MaxLength       =   20
            TabIndex        =   88
            Top             =   285
            Width           =   1440
         End
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5280
            TabIndex        =   87
            Top             =   278
            Width           =   1335
         End
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3120
            TabIndex        =   86
            Top             =   278
            Width           =   1335
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1395
            TabIndex        =   89
            Top             =   285
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
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
            Left            =   4680
            TabIndex        =   93
            Top             =   338
            Width           =   600
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
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
            Left            =   2295
            TabIndex        =   92
            Top             =   338
            Width           =   750
         End
         Begin VB.Label Label30 
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
            Index           =   0
            Left            =   300
            TabIndex        =   91
            Top             =   338
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº :"
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
            Left            =   6750
            TabIndex        =   90
            Top             =   360
            Width           =   345
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados de Entrega"
         Height          =   2055
         Left            =   75
         TabIndex        =   66
         Top             =   -30
         Width           =   9075
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   3405
            TabIndex        =   77
            Top             =   1125
            Width           =   735
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   3405
            MaxLength       =   10
            TabIndex        =   76
            Top             =   690
            Width           =   1290
         End
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   6510
            TabIndex        =   75
            Top             =   270
            Width           =   2235
         End
         Begin VB.ComboBox FilialEntrega 
            Height          =   315
            Left            =   1830
            TabIndex        =   74
            Top             =   225
            Width           =   1935
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frete por conta do"
            Height          =   870
            Index           =   1
            Left            =   180
            TabIndex        =   71
            Top             =   600
            Width           =   1680
            Begin VB.OptionButton Emitente 
               Caption         =   "Emitente"
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
               Left            =   135
               TabIndex        =   73
               Top             =   285
               Value           =   -1  'True
               Width           =   1200
            End
            Begin VB.OptionButton Destinatario 
               Caption         =   "Destinatário"
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
               Left            =   135
               TabIndex        =   72
               Top             =   555
               Width           =   1440
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Redespacho"
            Height          =   1065
            Left            =   4995
            TabIndex        =   67
            Top             =   765
            Width           =   3870
            Begin VB.CheckBox RedespachoCli 
               Caption         =   "por conta do cliente"
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
               Left            =   225
               TabIndex        =   69
               Top             =   705
               Width           =   2100
            End
            Begin VB.ComboBox TranspRedespacho 
               Height          =   315
               Left            =   1515
               TabIndex        =   68
               Top             =   285
               Width           =   2220
            End
            Begin VB.Label TranspRedLabel 
               AutoSize        =   -1  'True
               Caption         =   "Transportadora:"
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
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   70
               Top             =   330
               Width           =   1365
            End
         End
         Begin MSComCtl2.UpDown UpDownEntregaPV 
            Height          =   300
            Left            =   2940
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   1605
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntregaPV 
            Height          =   300
            Left            =   1815
            TabIndex        =   79
            Top             =   1605
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "U.F. da Placa:"
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
            Left            =   2100
            TabIndex        =   84
            Top             =   1185
            Width           =   1245
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Placa Veículo:"
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
            Left            =   2085
            TabIndex        =   83
            Top             =   750
            Width           =   1275
         End
         Begin VB.Label TransportadoraLabel 
            AutoSize        =   -1  'True
            Caption         =   "Transportadora:"
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
            Left            =   5085
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   315
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filial para Entrega:"
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
            Index           =   0
            Left            =   150
            TabIndex        =   81
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de Entrega:"
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
            Index           =   5
            Left            =   300
            TabIndex        =   80
            Top             =   1635
            Width           =   1470
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Complemento"
         Height          =   1635
         Index           =   7
         Left            =   75
         TabIndex        =   53
         Top             =   2850
         Width           =   9120
         Begin VB.TextBox Mensagem 
            Height          =   300
            Left            =   1485
            MaxLength       =   250
            TabIndex        =   55
            Top             =   270
            Width           =   7320
         End
         Begin VB.ComboBox CanalVenda 
            Height          =   315
            Left            =   1470
            TabIndex        =   54
            Top             =   1170
            Width           =   1605
         End
         Begin MSMask.MaskEdBox PedidoCliente 
            Height          =   300
            Left            =   4395
            TabIndex        =   56
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   300
            Left            =   4395
            TabIndex        =   57
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   300
            Left            =   1470
            TabIndex        =   58
            Top             =   720
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cubagem 
            Height          =   300
            Left            =   7740
            TabIndex        =   59
            Top             =   750
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PedidoRepr 
            Height          =   300
            Left            =   7710
            TabIndex        =   200
            Top             =   1185
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ped. Representante:"
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
            Left            =   5895
            TabIndex        =   201
            Top             =   1230
            Width           =   1770
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cubagem:"
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
            Left            =   6795
            TabIndex        =   65
            Top             =   795
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Cliente:"
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
            Left            =   3090
            TabIndex        =   64
            Top             =   1230
            Width           =   1290
         End
         Begin VB.Label MensagemLabel 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem:"
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
            Left            =   495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   63
            Top             =   375
            Width           =   975
         End
         Begin VB.Label CanalVendaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Canal de Venda:"
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
            Left            =   30
            TabIndex        =   62
            Top             =   1230
            Width           =   1425
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Peso Líquido:"
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
            Left            =   3195
            TabIndex        =   61
            Top             =   765
            Width           =   1200
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto:"
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
            Left            =   465
            TabIndex        =   60
            Top             =   765
            Width           =   1005
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6675
      ScaleHeight     =   450
      ScaleWidth      =   2610
      TabIndex        =   168
      TabStop         =   0   'False
      Top             =   90
      Width           =   2670
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2100
         Picture         =   "PedidoServicoOcx.ctx":2DBF
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1590
         Picture         =   "PedidoServicoOcx.ctx":2F3D
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1095
         Picture         =   "PedidoServicoOcx.ctx":346F
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   600
         Picture         =   "PedidoServicoOcx.ctx":35F9
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   90
         Picture         =   "PedidoServicoOcx.ctx":3753
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Orçamento de Serviço"
      Height          =   645
      Index           =   0
      Left            =   15
      TabIndex        =   142
      Top             =   0
      Width           =   4725
      Begin VB.CommandButton BotaoOrcamento 
         Height          =   360
         Left            =   3840
         Picture         =   "PedidoServicoOcx.ctx":3855
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Trazer Dados"
         Top             =   195
         Width           =   360
      End
      Begin VB.CommandButton BotaoVerOrcamento 
         Height          =   360
         Left            =   4230
         Picture         =   "PedidoServicoOcx.ctx":3C27
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Trazer a tela de cadastro"
         Top             =   195
         Width           =   360
      End
      Begin VB.ComboBox FilialOrcamento 
         Height          =   315
         Left            =   2295
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
      Begin MSMask.MaskEdBox Orcamento 
         Height          =   300
         Left            =   870
         TabIndex        =   0
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label OrcamentoLabel 
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
         Left            =   75
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   144
         Top             =   300
         Width           =   720
      End
      Begin VB.Label LabelFilialOrcamento 
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
         Left            =   1785
         TabIndex        =   143
         Top             =   300
         Width           =   465
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5265
      Left            =   0
      TabIndex        =   169
      Top             =   675
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9287
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bloqueio"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
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
Attribute VB_Name = "PedidoServicoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTPedidoServico
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoDataReferenciaDown_Click()
     Call objCT.BotaoDataReferenciaDown_Click
End Sub

Private Sub BotaoDataReferenciaUp_Click()
     Call objCT.BotaoDataReferenciaUp_Click
End Sub

Private Sub BotaoOrcamento_Click()
    Call objCT.BotaoOrcamento_Click 'Por Leo em 09/04/03
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoSolicitacaoSRV_Click()
     Call objCT.BotaoSolicitacaoSRV_Click
End Sub

Private Sub BotaoSolicitacaoSRVAP_Click()
     Call objCT.BotaoSolicitacaoSRVAP_Click
End Sub

Private Sub BotaoVerOrcamento_Click()
    Call objCT.BotaoVerOrcamento_Click 'Por Leo em 04/04/02
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEntregaPV_Change()
    Call objCT.DataEntregaPV_Change
End Sub

Private Sub DataEntregaPV_GotFocus()
    Call objCT.DataEntregaPV_GotFocus
End Sub

Private Sub DataEntregaPV_Validate(Cancel As Boolean)
    Call objCT.DataEntregaPV_Validate(Cancel)
End Sub


Private Sub DataReferencia_GotFocus()
     Call objCT.DataReferencia_GotFocus
End Sub

Private Sub Destinatario_Click()
     Call objCT.Destinatario_Click
End Sub
'Trecho Inserido por Leo em 02/04/02 *********************
Private Sub DiretoIndireto_Change()

    Call objCT.DiretoIndireto_Change

End Sub

Private Sub DiretoIndireto_GotFocus()

    Call objCT.DiretoIndireto_GotFocus

End Sub

Private Sub DiretoIndireto_KeyPress(KeyAscii As Integer)

    Call objCT.DiretoIndireto_KeyPress(KeyAscii)

End Sub

Private Sub DiretoIndireto_Validate(Cancel As Boolean)

    Call objCT.DiretoIndireto_Validate(Cancel)

End Sub

'Leo em 02/04/02 até aqui ************************************

'Janaina
'Private Sub Embalagem_Change()
'     Call objCT.Embalagem_Change
'End Sub
'
Private Sub Emitente_Click()
     Call objCT.Emitente_Click
End Sub

Private Sub FilialFaturamento_Change()
     Call objCT.FilialFaturamento_Change
End Sub

Private Sub FilialFaturamento_Click()
     Call objCT.FilialFaturamento_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub FilialOrcamento_Change()
    Call objCT.FilialOrcamento_Change 'Por Leo em 26/03/02
End Sub

Private Sub FilialOrcamento_Click()
    Call objCT.FilialOrcamento_Click  'Por Leo em 26/03/02
End Sub

Private Sub FilialOrcamento_Validate(Cancel As Boolean)
    Call objCT.FilialOrcamento_Validate(Cancel) 'Por Leo em 26/03/02
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItens_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub LabelCliente_Click()
     Call objCT.LabelCliente_Click
End Sub
'Incluido por leo em 26/03/02
Private Sub LabelFilialOrcamento_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelFilialOrcamento, Source, X, Y)
End Sub
'Incluido por leo em 26/03/02
Private Sub LabelFilialOrcamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelFilialOrcamento, Button, Shift, X, Y)
End Sub

Private Sub LabelSolicitaSRV_Click()
    Call objCT.LabelSolicitaSRV_Click
End Sub

Private Sub MensagemLabel_Click()
     Call objCT.MensagemLabel_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Private Sub CondPagtoLabel_Click()
     Call objCT.CondPagtoLabel_Click
End Sub

Private Sub BotaoServicos_Click()
     Call objCT.BotaoServicos_Click
End Sub

Private Sub BotaoPecas_Click()
     Call objCT.BotaoPecas_Click
End Sub

Private Sub BotaoEstoquePeca_Click()
     Call objCT.BotaoEstoquePeca_Click
End Sub

Function Trata_Parametros(Optional objPedidoVenda As ClassPedidoDeVenda) As Long
     Trata_Parametros = objCT.Trata_Parametros(objPedidoVenda)
End Function

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub Orcamento_Change()
    Call objCT.Orcamento_Change 'Por leo em 26/03/02
End Sub

Private Sub Orcamento_GotFocus()
    Call objCT.Orcamento_GotFocus 'Por leo em 26/03/02
End Sub

'Incluido por leo em 26/03/02
Private Sub OrcamentoLabel_Click()
    Call objCT.OrcamentoLabel_Click
End Sub

'Incluido por leo em 26/03/02
Private Sub OrcamentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(OrcamentoLabel, Source, X, Y)
End Sub
'Incluido por leo em 26/03/02
Private Sub OrcamentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(OrcamentoLabel, Button, Shift, X, Y)
End Sub

Private Sub PesoBruto_Change()
     Call objCT.PesoBruto_Change
End Sub

Private Sub PesoLiquido_Change()
     Call objCT.PesoLiquido_Change
End Sub

Private Sub Placa_Change()
     Call objCT.Placa_Change
End Sub

Private Sub PlacaUF_Change()
     Call objCT.PlacaUF_Change
End Sub

Private Sub PlacaUF_Click()
     Call objCT.PlacaUF_Click
End Sub

Private Sub PlacaUF_Validate(Cancel As Boolean)
     Call objCT.PlacaUF_Validate(Cancel)
End Sub
''Janaina
'Private Sub QuantEmb_Change()
'    Call objCT.QuantEmb_Change
'End Sub
''Janaina
'Private Sub QuantProduto_Change()
'    Call objCT.QuantProduto_Change
'End Sub


Private Sub RedespachoCli_Click()
    Call objCT.RedespachoCli_Click
End Sub

Private Sub TabStripSRV_Click()
     Call objCT.TabStripSRV_Click
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub CanalVenda_Change()
     Call objCT.CanalVenda_Change
End Sub

Private Sub CanalVenda_Click()
     Call objCT.CanalVenda_Click
End Sub

Private Sub CobrancaAutomatica_Click()
     Call objCT.CobrancaAutomatica_Click
End Sub

Private Sub CondicaoPagamento_Change()
     Call objCT.CondicaoPagamento_Change
End Sub

Private Sub CondicaoPagamento_Click()
     Call objCT.CondicaoPagamento_Click
End Sub

Private Sub DataReferencia_Change()
     Call objCT.DataReferencia_Change
End Sub

Private Sub FilialEntrega_Change()
     Call objCT.FilialEntrega_Change
End Sub

Private Sub FilialEntrega_Click()
     Call objCT.FilialEntrega_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub GridItens_Click()
     Call objCT.GridItens_Click
End Sub

Private Sub GridItens_EnterCell()
     Call objCT.GridItens_EnterCell
End Sub

Private Sub GridItens_GotFocus()
     Call objCT.GridItens_GotFocus
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)
     Call objCT.GridItens_KeyPress(KeyAscii)
End Sub

Private Sub GridItens_LeaveCell()
     Call objCT.GridItens_LeaveCell
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
     Call objCT.GridItens_Validate(Cancel)
End Sub

Private Sub GridItens_RowColChange()
     Call objCT.GridItens_RowColChange
End Sub

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub PedidoCliente_Change()
     Call objCT.PedidoCliente_Change
End Sub

Private Sub PercAcrescFin_Change()
     Call objCT.PercAcrescFin_Change
End Sub

Private Sub ComissaoAutomatica_Click()
     Call objCT.ComissaoAutomatica_Click
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_GotFocus()
     Call objCT.Produto_GotFocus
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
     Call objCT.Produto_KeyPress(KeyAscii)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub UnidadeMed_Change()
     Call objCT.UnidadeMed_Change
End Sub

Private Sub UnidadeMed_GotFocus()
     Call objCT.UnidadeMed_GotFocus
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMed_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)
     Call objCT.UnidadeMed_Validate(Cancel)
End Sub

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub Quantidade_GotFocus()
     Call objCT.Quantidade_GotFocus
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
     Call objCT.Quantidade_KeyPress(KeyAscii)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub PrecoUnitario_Change()
     Call objCT.PrecoUnitario_Change
End Sub

Private Sub PrecoUnitario_GotFocus()
     Call objCT.PrecoUnitario_GotFocus
End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoUnitario_KeyPress(KeyAscii)
End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)
     Call objCT.PrecoUnitario_Validate(Cancel)
End Sub

Private Sub PrecoTotal_Change()
     Call objCT.PrecoTotal_Change
End Sub

Private Sub PrecoTotal_GotFocus()
     Call objCT.PrecoTotal_GotFocus
End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoTotal_KeyPress(KeyAscii)
End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)
     Call objCT.PrecoTotal_Validate(Cancel)
End Sub

Private Sub QuantCancelada_Change()
     Call objCT.QuantCancelada_Change
End Sub

Private Sub QuantCancelada_GotFocus()
     Call objCT.QuantCancelada_GotFocus
End Sub

Private Sub QuantCancelada_KeyPress(KeyAscii As Integer)
     Call objCT.QuantCancelada_KeyPress(KeyAscii)
End Sub

Private Sub QuantCancelada_Validate(Cancel As Boolean)
     Call objCT.QuantCancelada_Validate(Cancel)
End Sub

Private Sub QuantReservadaPV_Change()
     Call objCT.QuantReservadaPV_Change
End Sub

Private Sub QuantReservadaPV_GotFocus()
     Call objCT.QuantReservadaPV_GotFocus
End Sub

Private Sub QuantReservadaPV_KeyPress(KeyAscii As Integer)
     Call objCT.QuantReservadaPV_KeyPress(KeyAscii)
End Sub

Private Sub QuantReservadaPV_Validate(Cancel As Boolean)
     Call objCT.QuantReservadaPV_Validate(Cancel)
End Sub

Private Sub QuantFaturada_Change()
     Call objCT.QuantFaturada_Change
End Sub

Private Sub QuantFaturada_GotFocus()
     Call objCT.QuantFaturada_GotFocus
End Sub

Private Sub QuantFaturada_KeyPress(KeyAscii As Integer)
     Call objCT.QuantFaturada_KeyPress(KeyAscii)
End Sub

Private Sub QuantFaturada_Validate(Cancel As Boolean)
     Call objCT.QuantFaturada_Validate(Cancel)
End Sub

Private Sub PercentDesc_Change()
     Call objCT.PercentDesc_Change
End Sub

Private Sub PercentDesc_GotFocus()
     Call objCT.PercentDesc_GotFocus
End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)
     Call objCT.PercentDesc_KeyPress(KeyAscii)
End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)
     Call objCT.PercentDesc_Validate(Cancel)
End Sub

Private Sub Desconto_Change()
     Call objCT.Desconto_Change
End Sub

Private Sub Desconto_GotFocus()
     Call objCT.Desconto_GotFocus
End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto_KeyPress(KeyAscii)
End Sub

Private Sub Desconto_Validate(Cancel As Boolean)
     Call objCT.Desconto_Validate(Cancel)
End Sub

Private Sub DescricaoProduto_Change()
     Call objCT.DescricaoProduto_Change
End Sub

Private Sub DescricaoProduto_GotFocus()
     Call objCT.DescricaoProduto_GotFocus
End Sub

Private Sub DescricaoProduto_KeyPress(KeyAscii As Integer)
     Call objCT.DescricaoProduto_KeyPress(KeyAscii)
End Sub

Private Sub DescricaoProduto_Validate(Cancel As Boolean)
     Call objCT.DescricaoProduto_Validate(Cancel)
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)
     Call objCT.CondicaoPagamento_Validate(Cancel)
End Sub

Private Sub UpDownEntregaPV_DownClick()
    Call objCT.UpDownEntregaPV_DownClick
End Sub

Private Sub UpDownEntregaPV_UpClick()
    Call objCT.UpDownEntregaPV_UpClick
End Sub

Private Sub UserControl_EnterFocus()
''??? foi susbstituido por chamada no activate. Jones 17/02/2000
''''Call objCT.CargaPosFormLoad
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTPedidoServico
    Set objCT.objUserControl = Me
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorDespesas_Change()
     Call objCT.ValorDespesas_Change
End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)
     Call objCT.ValorDespesas_Validate(Cancel)
End Sub

Private Sub ValorFrete_Change()
     Call objCT.ValorFrete_Change
End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)
     Call objCT.ValorFrete_Validate(Cancel)
End Sub

Private Sub ValorProdutos_Change()
     Call objCT.ValorProdutos_Change
End Sub

Private Sub ValorSeguro_Change()
     Call objCT.ValorSeguro_Change
End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)
     Call objCT.ValorSeguro_Validate(Cancel)
End Sub

Private Sub FilialEntrega_Validate(Cancel As Boolean)
     Call objCT.FilialEntrega_Validate(Cancel)
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub CanalVenda_Validate(Cancel As Boolean)
     Call objCT.CanalVenda_Validate(Cancel)
End Sub

Private Sub DataReferencia_Validate(Cancel As Boolean)
     Call objCT.DataReferencia_Validate(Cancel)
End Sub

Private Sub BotaoVendedores_Click()
     Call objCT.BotaoVendedores_Click
End Sub

Private Sub BotaoLiberaBloqueio_Click()
     Call objCT.BotaoLiberaBloqueio_Click
End Sub

Private Sub BotaoLibera_Click()
     Call objCT.BotaoLibera_Click
End Sub

Private Sub BotaoReserva_Click()
     Call objCT.BotaoReserva_Click
End Sub

Private Sub DataVencimento_Change()
     Call objCT.DataVencimento_Change
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub ValorParcela_Change()
     Call objCT.ValorParcela_Change
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub TipoDesconto1_Change()
     Call objCT.TipoDesconto1_Change
End Sub

Private Sub TipoDesconto1_GotFocus()
     Call objCT.TipoDesconto1_GotFocus
End Sub

Private Sub TipoDesconto1_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto1_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto1_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto1_Validate(Cancel)
End Sub

Private Sub TipoDesconto2_Change()
     Call objCT.TipoDesconto2_Change
End Sub

Private Sub TipoDesconto2_GotFocus()
     Call objCT.TipoDesconto2_GotFocus
End Sub

Private Sub TipoDesconto2_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto2_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto2_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto2_Validate(Cancel)
End Sub

Private Sub TipoDesconto3_Change()
     Call objCT.TipoDesconto3_Change
End Sub

Private Sub TipoDesconto3_GotFocus()
     Call objCT.TipoDesconto3_GotFocus
End Sub

Private Sub TipoDesconto3_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto3_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto3_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto3_Validate(Cancel)
End Sub

Private Sub Desconto1Ate_Change()
     Call objCT.Desconto1Ate_Change
End Sub

Private Sub Desconto1Ate_GotFocus()
     Call objCT.Desconto1Ate_GotFocus
End Sub

Private Sub Desconto1Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto1Ate_Validate(Cancel)
End Sub

Private Sub Desconto2Ate_Change()
     Call objCT.Desconto2Ate_Change
End Sub

Private Sub Desconto2Ate_GotFocus()
     Call objCT.Desconto2Ate_GotFocus
End Sub

Private Sub Desconto2Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto2Ate_Validate(Cancel)
End Sub

Private Sub Desconto3Ate_Change()
     Call objCT.Desconto3Ate_Change
End Sub

Private Sub Desconto3Ate_GotFocus()
     Call objCT.Desconto3Ate_GotFocus
End Sub

Private Sub Desconto3Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto3Ate_Validate(Cancel)
End Sub

Private Sub Desconto1Valor_Change()
     Call objCT.Desconto1Valor_Change
End Sub

Private Sub Desconto1Valor_GotFocus()
     Call objCT.Desconto1Valor_GotFocus
End Sub

Private Sub Desconto1Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto1Valor_Validate(Cancel)
End Sub

Private Sub Desconto2Valor_Change()
     Call objCT.Desconto2Valor_Change
End Sub

Private Sub Desconto2Valor_GotFocus()
     Call objCT.Desconto2Valor_GotFocus
End Sub

Private Sub Desconto2Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto2Valor_Validate(Cancel)
End Sub

Private Sub Desconto3Valor_Change()
     Call objCT.Desconto3Valor_Change
End Sub

Private Sub Desconto3Valor_GotFocus()
     Call objCT.Desconto3Valor_GotFocus
End Sub

Private Sub Desconto3Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto3Valor_Validate(Cancel)
End Sub

Private Sub Desconto1Percentual_Change()
     Call objCT.Desconto1Percentual_Change
End Sub

Private Sub Desconto1Percentual_GotFocus()
     Call objCT.Desconto1Percentual_GotFocus
End Sub

Private Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto1Percentual_Validate(Cancel)
End Sub

Private Sub Desconto2Percentual_Change()
     Call objCT.Desconto2Percentual_Change
End Sub

Private Sub Desconto2Percentual_GotFocus()
     Call objCT.Desconto2Percentual_GotFocus
End Sub

Private Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto2Percentual_Validate(Cancel)
End Sub

Private Sub Desconto3Percentual_Change()
     Call objCT.Desconto3Percentual_Change
End Sub

Private Sub Desconto3Percentual_GotFocus()
     Call objCT.Desconto3Percentual_GotFocus
End Sub

Private Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto3Percentual_Validate(Cancel)
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_GotFocus()
     Call objCT.Vendedor_GotFocus
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
     Call objCT.Vendedor_KeyPress(KeyAscii)
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Private Sub PercentualComissao_Change()
     Call objCT.PercentualComissao_Change
End Sub

Private Sub PercentualComissao_GotFocus()
     Call objCT.PercentualComissao_GotFocus
End Sub

Private Sub PercentualComissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualComissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualComissao_Validate(Cancel As Boolean)
     Call objCT.PercentualComissao_Validate(Cancel)
End Sub

Private Sub ValorBase_Change()
     Call objCT.ValorBase_Change
End Sub

Private Sub ValorBase_GotFocus()
     Call objCT.ValorBase_GotFocus
End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBase_KeyPress(KeyAscii)
End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)
     Call objCT.ValorBase_Validate(Cancel)
End Sub

Private Sub ValorComissao_Change()
     Call objCT.ValorComissao_Change
End Sub

Private Sub ValorComissao_GotFocus()
     Call objCT.ValorComissao_GotFocus
End Sub

Private Sub ValorComissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorComissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)
     Call objCT.ValorComissao_Validate(Cancel)
End Sub

Private Sub PercentualEmissao_Change()
     Call objCT.PercentualEmissao_Change
End Sub

Private Sub PercentualEmissao_GotFocus()
     Call objCT.PercentualEmissao_GotFocus
End Sub

Private Sub PercentualEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualEmissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualEmissao_Validate(Cancel As Boolean)
     Call objCT.PercentualEmissao_Validate(Cancel)
End Sub

Private Sub ValorEmissao_Change()
     Call objCT.ValorEmissao_Change
End Sub

Private Sub ValorEmissao_GotFocus()
     Call objCT.ValorEmissao_GotFocus
End Sub

Private Sub ValorEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorEmissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorEmissao_Validate(Cancel As Boolean)
     Call objCT.ValorEmissao_Validate(Cancel)
End Sub

Private Sub PercentualBaixa_Change()
     Call objCT.PercentualBaixa_Change
End Sub

Private Sub PercentualBaixa_GotFocus()
     Call objCT.PercentualBaixa_GotFocus
End Sub

Private Sub PercentualBaixa_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualBaixa_KeyPress(KeyAscii)
End Sub

Private Sub PercentualBaixa_Validate(Cancel As Boolean)
     Call objCT.PercentualBaixa_Validate(Cancel)
End Sub

Private Sub ValorBaixa_Change()
     Call objCT.ValorBaixa_Change
End Sub

Private Sub ValorBaixa_GotFocus()
     Call objCT.ValorBaixa_GotFocus
End Sub

Private Sub ValorBaixa_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBaixa_KeyPress(KeyAscii)
End Sub

Private Sub ValorBaixa_Validate(Cancel As Boolean)
     Call objCT.ValorBaixa_Validate(Cancel)
End Sub

Private Sub GridComissoes_Click()
     Call objCT.GridComissoes_Click
End Sub

Private Sub GridComissoes_GotFocus()
     Call objCT.GridComissoes_GotFocus
End Sub

Private Sub GridComissoes_EnterCell()
     Call objCT.GridComissoes_EnterCell
End Sub

Private Sub GridComissoes_LeaveCell()
     Call objCT.GridComissoes_LeaveCell
End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridComissoes_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)
     Call objCT.GridComissoes_KeyPress(KeyAscii)
End Sub

Private Sub GridComissoes_Validate(Cancel As Boolean)
     Call objCT.GridComissoes_Validate(Cancel)
End Sub

Private Sub GridComissoes_RowColChange()
     Call objCT.GridComissoes_RowColChange
End Sub

Private Sub GridComissoes_Scroll()
     Call objCT.GridComissoes_Scroll
End Sub

Private Sub BLGTipoBloqueio_Change()
     Call objCT.TipoBloqueio_Change
End Sub

Private Sub BLGTipoBloqueio_GotFocus()
     Call objCT.TipoBloqueio_GotFocus
End Sub

Private Sub BLGTipoBloqueio_KeyPress(KeyAscii As Integer)
     Call objCT.TipoBloqueio_KeyPress(KeyAscii)
End Sub

Private Sub BLGTipoBloqueio_Validate(Cancel As Boolean)
     Call objCT.TipoBloqueio_Validate(Cancel)
End Sub

Private Sub BLGDataBloqueio_Change()
     Call objCT.DataBloqueio_Change
End Sub

Private Sub BLGDataBloqueio_GotFocus()
     Call objCT.DataBloqueio_GotFocus
End Sub

Private Sub BLGDataBloqueio_KeyPress(KeyAscii As Integer)
     Call objCT.DataBloqueio_KeyPress(KeyAscii)
End Sub

Private Sub BLGDataBloqueio_Validate(Cancel As Boolean)
     Call objCT.DataBloqueio_Validate(Cancel)
End Sub

Private Sub BLGObservacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub BLGObservacao_GotFocus()
     Call objCT.Observacao_GotFocus
End Sub

Private Sub BLGObservacao_KeyPress(KeyAscii As Integer)
     Call objCT.Observacao_KeyPress(KeyAscii)
End Sub

Private Sub BLGObservacao_Validate(Cancel As Boolean)
     Call objCT.Observacao_Validate(Cancel)
End Sub

Private Sub BLGUsuario_Change()
     Call objCT.CodUsuario_Change
End Sub

Private Sub BLGUsuario_GotFocus()
     Call objCT.CodUsuario_GotFocus
End Sub

Private Sub BLGUsuario_KeyPress(KeyAscii As Integer)
     Call objCT.CodUsuario_KeyPress(KeyAscii)
End Sub

Private Sub BLGUsuario_Validate(Cancel As Boolean)
     Call objCT.CodUsuario_Validate(Cancel)
End Sub

Private Sub BLGResponsavel_Change()
     Call objCT.ResponsavelBL_Change
End Sub

Private Sub BLGResponsavel_GotFocus()
     Call objCT.ResponsavelBL_GotFocus
End Sub

Private Sub BLGResponsavel_KeyPress(KeyAscii As Integer)
     Call objCT.ResponsavelBL_KeyPress(KeyAscii)
End Sub

Private Sub BLGResponsavel_Validate(Cancel As Boolean)
     Call objCT.ResponsavelBL_Validate(Cancel)
End Sub

Private Sub BLGGridBloqueio_Click()
     Call objCT.GridBloqueio_Click
End Sub

Private Sub BLGGridBloqueio_GotFocus()
     Call objCT.GridBloqueio_GotFocus
End Sub

Private Sub BLGGridBloqueio_EnterCell()
     Call objCT.GridBloqueio_EnterCell
End Sub

Private Sub BLGGridBloqueio_LeaveCell()
     Call objCT.GridBloqueio_LeaveCell
End Sub

Private Sub BLGGridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridBloqueio_KeyDown(KeyCode, Shift)
End Sub

Private Sub BLGGridBloqueio_KeyPress(KeyAscii As Integer)
     Call objCT.GridBloqueio_KeyPress(KeyAscii)
End Sub

Private Sub BLGGridBloqueio_Validate(Cancel As Boolean)
     Call objCT.GridBloqueio_Validate(Cancel)
End Sub

Private Sub BLGGridBloqueio_RowColChange()
     Call objCT.GridBloqueio_RowColChange
End Sub

Private Sub BLGGridBloqueio_Scroll()
     Call objCT.GridBloqueio_Scroll
End Sub

Private Sub GridAlocacao_Click()
     Call objCT.GridAlocacao_Click
End Sub

Private Sub GridAlocacao_GotFocus()
     Call objCT.GridAlocacao_GotFocus
End Sub

Private Sub GridAlocacao_EnterCell()
     Call objCT.GridAlocacao_EnterCell
End Sub

Private Sub GridAlocacao_LeaveCell()
     Call objCT.GridAlocacao_LeaveCell
End Sub

Private Sub GridAlocacao_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridAlocacao_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridAlocacao_KeyPress(KeyAscii As Integer)
     Call objCT.GridAlocacao_KeyPress(KeyAscii)
End Sub

Private Sub GridAlocacao_Validate(Cancel As Boolean)
     Call objCT.GridAlocacao_Validate(Cancel)
End Sub

Private Sub GridAlocacao_RowColChange()
     Call objCT.GridAlocacao_RowColChange
End Sub

Private Sub GridAlocacao_Scroll()
     Call objCT.GridAlocacao_Scroll
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoImprimir_Click()
     Call objCT.BotaoImprimir_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub FilialFaturamento_Validate(Cancel As Boolean)
     Call objCT.FilialFaturamento_Validate(Cancel)
End Sub

Private Sub PercAcrescFin_Validate(Cancel As Boolean)
     Call objCT.PercAcrescFin_Validate(Cancel)
End Sub

Private Sub VolumeEspecie_Change()
     Call objCT.VolumeEspecie_Change
End Sub

Private Sub VolumeMarca_Change()
     Call objCT.VolumeMarca_Change
End Sub

Private Sub VolumeEspecie_Validate(Cancel As Boolean)
     Call objCT.VolumeEspecie_Validate(Cancel)
End Sub

Private Sub VolumeMarca_Validate(Cancel As Boolean)
     Call objCT.VolumeMarca_Validate(Cancel)
End Sub

Private Sub VolumeNumero_Change()
     Call objCT.VolumeNumero_Change
End Sub

Private Sub VolumeQuant_Change()
     Call objCT.VolumeQuant_Change
End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)
     Call objCT.PesoLiquido_Validate(Cancel)
End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)
     Call objCT.PesoBruto_Validate(Cancel)
End Sub

Private Sub VolumeQuant_GotFocus()
     Call objCT.VolumeQuant_GotFocus
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        Set objCT.objUserControl = Nothing
        Set objCT = Nothing
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

'Private Sub Label13_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label13(Index), Source, X, Y)
'End Sub
'
'Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label13(Index), Button, Shift, X, Y)
'End Sub

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

'Private Sub Label17_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label17(Index), Source, X, Y)
'End Sub
'
'Private Sub Label17_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label17(Index), Button, Shift, X, Y)
'End Sub

'Private Sub Label7_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label7(Index), Source, X, Y)
'End Sub
'
'Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label7(Index), Button, Shift, X, Y)
'End Sub

'Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label2(Index), Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label19_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label19(Index), Source, X, Y)
'End Sub
'
'Private Sub Label19_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label19(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label18_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label18(Index), Source, X, Y)
'End Sub
'
'Private Sub Label18_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label18(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label15(Index), Source, X, Y)
'End Sub
'
'Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label16_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label16(Index), Source, X, Y)
'End Sub
'
'Private Sub Label16_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label16(Index), Button, Shift, X, Y)
'End Sub

'Private Sub Label6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label6(Index), Source, X, Y)
'End Sub
'
'Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label6(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label44_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label44(Index), Source, X, Y)
'End Sub
'
'Private Sub Label44_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label44(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label30_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label30(Index), Source, X, Y)
'End Sub
'
'Private Sub Label30_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label30(Index), Button, Shift, X, Y)
'End Sub


Private Sub ICMSSubstValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor1, Source, X, Y)
End Sub

Private Sub ICMSSubstValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstBase1, Source, X, Y)
End Sub

Private Sub ICMSSubstBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstBase1, Button, Shift, X, Y)
End Sub

Private Sub ICMSValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSValor1, Source, X, Y)
End Sub

Private Sub ICMSValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSBase1, Source, X, Y)
End Sub

Private Sub ICMSBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSBase1, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub IPIValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor1, Source, X, Y)
End Sub

Private Sub IPIValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor1, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub MensagemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MensagemLabel, Source, X, Y)
End Sub

Private Sub MensagemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MensagemLabel, Button, Shift, X, Y)
End Sub

Private Sub CanalVendaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CanalVendaLabel, Source, X, Y)
End Sub

Private Sub CanalVendaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CanalVendaLabel, Button, Shift, X, Y)
End Sub

'Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label32, Source, X, Y)
'End Sub
'
'Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label5, Source, X, Y)
'End Sub
'
'Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
'End Sub

Private Sub TranspRedLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TranspRedLabel, Source, X, Y)
End Sub

Private Sub TranspRedLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TranspRedLabel, Button, Shift, X, Y)
End Sub

Private Sub ProdutoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoDescricao, Source, X, Y)
End Sub

Private Sub ProdutoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoDescricao, Button, Shift, X, Y)
End Sub

'Private Sub LabelTotaisComissoes_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelTotaisComissoes, Source, X, Y)
'End Sub

'Private Sub LabelTotaisComissoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelTotaisComissoes, Button, Shift, X, Y)
'End Sub

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

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub TotalValorBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorBase, Source, X, Y)
End Sub

Private Sub TotalValorBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorBase, Button, Shift, X, Y)
End Sub

Private Sub TranspRedespacho_Change()
     Call objCT.TranspRedespacho_Change
End Sub

Private Sub TranspRedespacho_Click()
     Call objCT.TranspRedespacho_Click
End Sub

Private Sub TranspRedespacho_Validate(Cancel As Boolean)
     Call objCT.TranspRedespacho_Validate(Cancel)
End Sub

Private Sub Cubagem_Change()
     Call objCT.Cubagem_Change
End Sub

Private Sub Cubagem_Validate(Cancel As Boolean)
    Call objCT.Cubagem_Validate(Cancel)
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub TranspRedLabel_Click()
     Call objCT.TranspRedLabel_Click
End Sub

Private Sub BotaoTodosPedidos_Click()
    Call objCT.BotaoTodosPedidos_Click
End Sub

Private Sub BotaoContato_Click()
     Call objCT.BotaoContato_Click
End Sub

'#####################################
'Inserido por Wagner 03/08/2006
Private Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub

Private Sub LabelProjeto_Click()
    Call objCT.LabelProjeto_Click
End Sub

Private Sub Projeto_Change()
     Call objCT.Projeto_Change
End Sub

Private Sub Projeto_GotFocus()
     Call objCT.Projeto_GotFocus
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub
'#####################################

Private Sub TipoPagto_Change()
     Call objCT.TipoPagto_Change
End Sub

Private Sub TipoPagto_GotFocus()
     Call objCT.TipoPagto_GotFocus
End Sub

Private Sub TipoPagto_KeyPress(KeyAscii As Integer)
     Call objCT.TipoPagto_KeyPress(KeyAscii)
End Sub

Private Sub TipoPagto_Validate(Cancel As Boolean)
     Call objCT.TipoPagto_Validate(Cancel)
End Sub

Private Sub BotaoTipoPagto_Click()
     Call objCT.BotaoTipoPagto_Click
End Sub

Private Sub DataEntregaSRV_Change()
    Call objCT.DataEntregaSRV_Change
End Sub

Private Sub DataEntregaSRV_GotFocus()
    Call objCT.DataEntregaSRV_GotFocus
End Sub

Private Sub DataEntregaSRV_KeyPress(KeyAscii As Integer)
    Call objCT.DataEntregaSRV_KeyPress(KeyAscii)
End Sub

Private Sub DataEntregaSRV_Validate(Cancel As Boolean)
    Call objCT.DataEntregaSRV_Validate(Cancel)
End Sub

Private Sub DescServicoSRV_Change()
     Call objCT.DescServicoSRV_Change
End Sub

Private Sub DescServicoSRV_GotFocus()
     Call objCT.DescServicoSRV_GotFocus
End Sub

Private Sub DescServicoSRV_KeyPress(KeyAscii As Integer)
     Call objCT.DescServicoSRV_KeyPress(KeyAscii)
End Sub

Private Sub DescServicoSRV_Validate(Cancel As Boolean)
     Call objCT.DescServicoSRV_Validate(Cancel)
End Sub

Private Sub GridItensSRV_Click()
     Call objCT.GridItensSRV_Click
End Sub

Private Sub GridItensSRV_EnterCell()
     Call objCT.GridItensSRV_EnterCell
End Sub

Private Sub GridItensSRV_GotFocus()
     Call objCT.GridItensSRV_GotFocus
End Sub

Private Sub GridItensSRV_KeyPress(KeyAscii As Integer)
     Call objCT.GridItensSRV_KeyPress(KeyAscii)
End Sub

Private Sub GridItensSRV_LeaveCell()
     Call objCT.GridItensSRV_LeaveCell
End Sub

Private Sub GridItensSRV_Validate(Cancel As Boolean)
     Call objCT.GridItensSRV_Validate(Cancel)
End Sub

Private Sub GridItensSRV_RowColChange()
     Call objCT.GridItensSRV_RowColChange
End Sub

Private Sub GridItensSRV_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItensSRV_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridItensSRV_Scroll()
     Call objCT.GridItensSRV_Scroll
End Sub

Private Sub PercentDescSRV_Change()
     Call objCT.PercentDescSRV_Change
End Sub

Private Sub PercentDescSRV_GotFocus()
     Call objCT.PercentDescSRV_GotFocus
End Sub

Private Sub PercentDescSRV_KeyPress(KeyAscii As Integer)
     Call objCT.PercentDescSRV_KeyPress(KeyAscii)
End Sub

Private Sub PercentDescSRV_Validate(Cancel As Boolean)
     Call objCT.PercentDescSRV_Validate(Cancel)
End Sub

Private Sub PrecoTotalSRV_Change()
     Call objCT.PrecoTotalSRV_Change
End Sub

Private Sub PrecoTotalSRV_GotFocus()
     Call objCT.PrecoTotalSRV_GotFocus
End Sub

Private Sub PrecoTotalSRV_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoTotalSRV_KeyPress(KeyAscii)
End Sub

Private Sub PrecoTotalSRV_Validate(Cancel As Boolean)
     Call objCT.PrecoTotalSRV_Validate(Cancel)
End Sub

Private Sub PrecoUnitarioSRV_Change()
     Call objCT.PrecoUnitarioSRV_Change
End Sub

Private Sub PrecoUnitarioSRV_GotFocus()
     Call objCT.PrecoUnitarioSRV_GotFocus
End Sub

Private Sub PrecoUnitarioSRV_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoUnitarioSRV_KeyPress(KeyAscii)
End Sub

Private Sub PrecoUnitarioSRV_Validate(Cancel As Boolean)
     Call objCT.PrecoUnitarioSRV_Validate(Cancel)
End Sub

Private Sub ServicoSRV_Change()
     Call objCT.ServicoSRV_Change
End Sub

Private Sub ServicoSRV_GotFocus()
     Call objCT.ServicoSRV_GotFocus
End Sub

Private Sub ServicoSRV_KeyPress(KeyAscii As Integer)
     Call objCT.ServicoSRV_KeyPress(KeyAscii)
End Sub

Private Sub ServicoSRV_Validate(Cancel As Boolean)
     Call objCT.ServicoSRV_Validate(Cancel)
End Sub

Private Sub DescontoSRV_Change()
     Call objCT.DescontoSRV_Change
End Sub

Private Sub DescontoSRV_GotFocus()
     Call objCT.DescontoSRV_GotFocus
End Sub

Private Sub DescontoSRV_KeyPress(KeyAscii As Integer)
     Call objCT.DescontoSRV_KeyPress(KeyAscii)
End Sub

Private Sub DescontoSRV_Validate(Cancel As Boolean)
     Call objCT.DescontoSRV_Validate(Cancel)
End Sub

Private Sub UnidadeMedSRV_Change()
     Call objCT.UnidadeMedSRV_Change
End Sub

Private Sub UnidadeMedSRV_GotFocus()
     Call objCT.UnidadeMedSRV_GotFocus
End Sub

Private Sub UnidadeMedSRV_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMedSRV_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMedSRV_Validate(Cancel As Boolean)
     Call objCT.UnidadeMedSRV_Validate(Cancel)
End Sub

Private Sub QuantidadeSRV_Change()
     Call objCT.QuantidadeSRV_Change
End Sub

Private Sub QuantidadeSRV_GotFocus()
     Call objCT.QuantidadeSRV_GotFocus
End Sub

Private Sub QuantidadeSRV_KeyPress(KeyAscii As Integer)
     Call objCT.QuantidadeSRV_KeyPress(KeyAscii)
End Sub

Private Sub QuantidadeSRV_Validate(Cancel As Boolean)
     Call objCT.QuantidadeSRV_Validate(Cancel)
End Sub

Private Sub QuantCanceladaSRV_Change()
     Call objCT.QuantCanceladaSRV_Change
End Sub

Private Sub QuantCanceladaSRV_GotFocus()
     Call objCT.QuantCanceladaSRV_GotFocus
End Sub

Private Sub QuantCanceladaSRV_KeyPress(KeyAscii As Integer)
     Call objCT.QuantCanceladaSRV_KeyPress(KeyAscii)
End Sub

Private Sub QuantCanceladaSRV_Validate(Cancel As Boolean)
     Call objCT.QuantCanceladaSRV_Validate(Cancel)
End Sub

Private Sub QuantFaturadaSRV_Change()
     Call objCT.QuantFaturadaSRV_Change
End Sub

Private Sub QuantFaturadaSRV_GotFocus()
     Call objCT.QuantFaturadaSRV_GotFocus
End Sub

Private Sub QuantFaturadaSRV_KeyPress(KeyAscii As Integer)
     Call objCT.QuantFaturadaSRV_KeyPress(KeyAscii)
End Sub

Private Sub QuantFaturadaSRV_Validate(Cancel As Boolean)
     Call objCT.QuantFaturadaSRV_Validate(Cancel)
End Sub

Private Sub PedidoRepr_Change()
     Call objCT.PedidoRepr_Change
End Sub

Private Sub PedidoRepr_Validate(Cancel As Boolean)
     Call objCT.PedidoRepr_Validate(Cancel)
End Sub

Private Sub ValorDescontoTit_Change()
     Call objCT.ValorDescontoTit_Change
End Sub

Private Sub ValorDescontoTit_Validate(Cancel As Boolean)
     Call objCT.ValorDescontoTit_Validate(Cancel)
End Sub

Private Sub BotaoInfoAdic_Click()
     Call objCT.BotaoInfoAdic_Click
End Sub

Private Sub ValorDescontoItens_Change()
     Call objCT.ValorDescontoItens_Change
End Sub

Private Sub ValorDescontoItens_Validate(Cancel As Boolean)
     Call objCT.ValorDescontoItens_Validate(Cancel)
End Sub

Private Sub PercDescontoItens_Change()
     Call objCT.PercDescontoItens_Change
End Sub

Private Sub PercDescontoItens_Validate(Cancel As Boolean)
     Call objCT.PercDescontoItens_Validate(Cancel)
End Sub

Private Sub BotaoDesfazer_Click()
    Call objCT.BotaoDesfazer_Click
End Sub

