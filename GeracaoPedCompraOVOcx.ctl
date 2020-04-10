VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl GeracaoPedCompraOVOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      Caption         =   "Orçamento de Venda"
      Height          =   645
      Index           =   302
      Left            =   105
      TabIndex        =   50
      Top             =   120
      Width           =   4500
      Begin VB.CommandButton BotaoOrcamento 
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
         Left            =   1845
         TabIndex        =   54
         Top             =   225
         Width           =   1350
      End
      Begin VB.CommandButton BotaoVerOrcamento 
         Caption         =   "Consultar"
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
         Left            =   3270
         TabIndex        =   53
         Top             =   225
         Width           =   1005
      End
      Begin MSMask.MaskEdBox Orcamento 
         Height          =   300
         Left            =   870
         TabIndex        =   51
         Top             =   240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   52
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   14745
      ScaleHeight     =   480
      ScaleWidth      =   2040
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   135
      Width           =   2100
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "GeracaoPedCompraOVOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1560
         Picture         =   "GeracaoPedCompraOVOcx.ctx":018A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "GeracaoPedCompraOVOcx.ctx":0308
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   90
         Picture         =   "GeracaoPedCompraOVOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8220
      Index           =   2
      Left            =   165
      TabIndex        =   0
      Top             =   840
      Width           =   16665
      Begin VB.CommandButton BotaoPedCotacao 
         Caption         =   "Pedido de Cotação ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6516
         TabIndex        =   34
         Top             =   15
         Width           =   2130
      End
      Begin VB.Frame Frame4 
         Caption         =   "Opção"
         Height          =   1464
         Left            =   12570
         TabIndex        =   31
         Top             =   6585
         Width           =   4020
         Begin VB.CommandButton BotaoGeraPedidos 
            Caption         =   "Gera Pedidos de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   420
            TabIndex        =   33
            Top             =   816
            Width           =   2670
         End
         Begin VB.CommandButton BotaoGravaConcorrencia 
            Caption         =   "Grava Concorrência"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   420
            TabIndex        =   32
            Top             =   276
            Width           =   2670
         End
      End
      Begin VB.Frame FrameCotacoes 
         Caption         =   "Cotações"
         Height          =   6135
         Index           =   2
         Left            =   45
         TabIndex        =   3
         Top             =   420
         Width           =   16515
         Begin VB.ComboBox MotivoEscolha 
            Height          =   315
            ItemData        =   "GeracaoPedCompraOVOcx.ctx":093C
            Left            =   6585
            List            =   "GeracaoPedCompraOVOcx.ctx":093E
            TabIndex        =   12
            Top             =   2355
            Width           =   1995
         End
         Begin VB.TextBox DescProdutoCot 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2550
            MaxLength       =   50
            TabIndex        =   11
            Top             =   270
            Width           =   4000
         End
         Begin VB.CheckBox EscolhidoCot 
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
            Left            =   900
            TabIndex        =   10
            Top             =   255
            Width           =   1020
         End
         Begin VB.ComboBox Moeda 
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   648
            Width           =   1440
         End
         Begin MSMask.MaskEdBox Cotacao 
            Height          =   228
            Left            =   396
            TabIndex        =   4
            Top             =   1584
            Width           =   1440
            _ExtentX        =   2540
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
         Begin MSMask.MaskEdBox TaxaForn 
            Height          =   228
            Left            =   396
            TabIndex        =   5
            Top             =   1296
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   423
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
         Begin MSMask.MaskEdBox PrecoUnitarioReal 
            Height          =   228
            Left            =   396
            TabIndex        =   6
            Top             =   1008
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataCotacao 
            Height          =   225
            Left            =   4005
            TabIndex        =   8
            Top             =   450
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorItem 
            Height          =   225
            Left            =   2655
            TabIndex        =   9
            Top             =   2250
            Width           =   1170
            _ExtentX        =   2064
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
         Begin MSMask.MaskEdBox PrecoUnitarioCot 
            Height          =   225
            Left            =   2715
            TabIndex        =   13
            Top             =   1980
            Width           =   1080
            _ExtentX        =   1905
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PedCotacao 
            Height          =   225
            Left            =   5100
            TabIndex        =   14
            Top             =   1980
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   225
            Left            =   6240
            TabIndex        =   15
            Top             =   1980
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeEntrega 
            Height          =   225
            Left            =   4335
            TabIndex        =   16
            Top             =   2400
            Width           =   1125
            _ExtentX        =   1984
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
         Begin MSMask.MaskEdBox ValorPresenteCot 
            Height          =   225
            Left            =   3930
            TabIndex        =   17
            Top             =   1995
            Width           =   1170
            _ExtentX        =   2064
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
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   1755
            TabIndex        =   18
            Top             =   2385
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrazoEntrega 
            Height          =   225
            Left            =   7455
            TabIndex        =   19
            Top             =   1980
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   3
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantComprarCot 
            Height          =   225
            Left            =   5475
            TabIndex        =   20
            Top             =   2355
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox CondPagtoCot 
            Height          =   225
            Left            =   1425
            TabIndex        =   21
            Top             =   1950
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   30
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UMCot 
            Height          =   225
            Left            =   4065
            TabIndex        =   22
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFornCot 
            Height          =   225
            Left            =   45
            TabIndex        =   23
            Top             =   1965
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorCot 
            Height          =   225
            Left            =   6330
            TabIndex        =   24
            Top             =   330
            Visible         =   0   'False
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantComprarMaxCot 
            Height          =   225
            Left            =   5145
            TabIndex        =   25
            Top             =   270
            Width           =   1125
            _ExtentX        =   1984
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
         Begin MSMask.MaskEdBox ProdutoCot 
            Height          =   225
            Left            =   1230
            TabIndex        =   26
            Top             =   285
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataNecess 
            Height          =   225
            Left            =   270
            TabIndex        =   27
            Top             =   2325
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCotacoes 
            Height          =   5820
            Left            =   105
            TabIndex        =   28
            Top             =   240
            Width           =   16335
            _ExtentX        =   28813
            _ExtentY        =   10266
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox AliquotaICMS 
            Height          =   225
            Left            =   2115
            TabIndex        =   29
            Top             =   2475
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   225
            Left            =   3960
            TabIndex        =   30
            Top             =   2400
            Width           =   1005
            _ExtentX        =   1773
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
      End
      Begin VB.ComboBox Ordenacao 
         Height          =   315
         ItemData        =   "GeracaoPedCompraOVOcx.ctx":0940
         Left            =   2280
         List            =   "GeracaoPedCompraOVOcx.ctx":0942
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   90
         Width           =   2325
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   3720
         Picture         =   "GeracaoPedCompraOVOcx.ctx":0944
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   7800
         Width           =   300
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   540
         Left            =   2565
         TabIndex        =   35
         Top             =   7125
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   953
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin VB.Label Concorrencia 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2565
         TabIndex        =   43
         Top             =   7785
         Width           =   1155
      End
      Begin VB.Label LabelConcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Concorrência:"
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
         Left            =   1305
         TabIndex        =   42
         Top             =   7845
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Financeira:"
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
         Left            =   4470
         TabIndex        =   41
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Label TaxaEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5970
         TabIndex        =   40
         Top             =   6780
         Width           =   1155
      End
      Begin VB.Label Label45 
         Caption         =   "Ordena por:"
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
         Left            =   1140
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label TotalItens 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2550
         TabIndex        =   38
         Top             =   6780
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total dos Itens:"
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
         Left            =   660
         TabIndex        =   37
         Top             =   6825
         Width           =   1845
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   1590
         TabIndex        =   36
         Top             =   7185
         Width           =   930
      End
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Comprador:"
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
      Left            =   6000
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   49
      Top             =   375
      Width           =   975
   End
   Begin VB.Label Comprador 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7155
      TabIndex        =   48
      Top             =   360
      Width           =   2370
   End
End
Attribute VB_Name = "GeracaoPedCompraOVOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOrcamento  As AdmEvento 'Incluido por Leo em 28/03/02
Attribute objEventoOrcamento.VB_VarHelpID = -1

'Variáveis Globais
Dim iAlterado As Integer
Dim objGridCotacoes As AdmGrid

Dim gcolItemConcorrencia As Collection
Dim gsOrdenacao As String
Dim asOrdenacao(2) As String
Dim asOrdenacaoString(2) As String

'Colunas do GridCotacoes
Dim iGrid_EscolhidoCot_Col As Integer
Dim iGrid_ProdutoCot_Col As Integer
Dim iGrid_DescProdutoCot_Col As Integer
Dim iGrid_CondPagtoCot_Col As Integer
Dim iGrid_QuantComprarMaxCot_Col As Integer
Dim iGrid_UMCot_Col As Integer
Dim iGrid_PrecoUnitarioCot_Col As Integer
Dim iGrid_ValorPresenteCot_Col As Integer
Dim iGrid_ValorItem_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_FornecedorCot_Col As Integer
Dim iGrid_FilialFornCot_Col As Integer
Dim iGrid_PedidoCot_Col As Integer
Dim iGrid_DataValidadeCot_Col As Integer
Dim iGrid_PrazoEntregaCot_Col As Integer
Dim iGrid_DataEntregaCot_Col As Integer
Dim iGrid_DataNecessidadeCot_Col As Integer
Dim iGrid_QuantidadeEntregaCot_Col As Integer
Dim iGrid_QuantComprarCot_Col As Integer
Dim iGrid_MotivoEscolhaCot_Col As Integer
Dim iGrid_DataCotacaoCot_Col As Integer
Dim iGrid_Moeda_Col As Integer
Dim iGrid_PrecoUnitario_RS_Col As Integer
Dim iGrid_TaxaForn_Col As Integer
Dim iGrid_CotacaoMoeda_Col As Integer

Public Sub Form_Load()

Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador
Dim lErro As Long
Dim lConcorrencia As Long
Dim iTipoTrib As Integer
Dim sDescricao As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load

    '###################################
    'Inserido por Wagner
    Call Formata_Controles
    '###################################
        
    Set objEventoOrcamento = New AdmEvento
    
    Set objGridCotacoes = New AdmGrid

    Set gcolItemConcorrencia = New Collection
    
    objComprador.sCodUsuario = gsUsuario

    'Verifica se gsUsuario é comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63669

    'Se gsUsuario nao é comprador==> erro
    If lErro = 50059 Then gError 63670

    objUsuario.sCodUsuario = objComprador.sCodUsuario

    'Lê o usuário
    lErro = CF("Usuario_Le", objUsuario)
    If lErro <> SUCESSO And lErro <> 36347 Then gError 63671

    'Se não encontrou o usuário ==> Erro
    If lErro = 36347 Then gError 63672

    'Coloca o Nome Reduzido do Comprador na tela
    Comprador.Caption = objUsuario.sNomeReduzido

    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError 108981

    'Preenche a combo de ordenacao
    Call Ordenacao_Carrega

    'Preenche a combo de MotivoEscolha
    lErro = Carrega_MotivoEscolha()
    If lErro <> SUCESSO Then gError 63795

    'Inicializa a máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoCot)
    If lErro <> SUCESSO Then gError 63676

    'Inicializa o GridCotacoes
    lErro = Inicializa_Grid_Cotacoes(objGridCotacoes)
    If lErro <> SUCESSO Then gError 63680

    'Coloca as Quantidades da tela no formato de Estoque
    QuantComprarMaxCot.Format = FORMATO_ESTOQUE
    QuantComprarCot.Format = FORMATO_ESTOQUE
    QuantidadeEntrega.Format = FORMATO_ESTOQUE

    'Coloca Taxa Financeira na tela
    TaxaEmpresa.Caption = Format(gobjCOM.dTaxaFinanceiraEmpresa, "Percent")

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 63669, 63671, 63673, 63674, 63675, 63676, 63677, 63680, 63795, 70496, 70497, 108980, 108981, 126946
            'Erros tratados nas rotinas chamadas

        Case 63670
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 63672
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuario.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160955)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoOrcamento = Nothing
    
    'libera as variaveis globais
    Set objGridCotacoes = Nothing
    
    Set gcolItemConcorrencia = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160956)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Cotacoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Cotacoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Cotacoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Taxa Forn.")
    objGridInt.colColuna.Add ("Cotação")
    objGridInt.colColuna.Add ("Preço Unitário (R$)")
    objGridInt.colColuna.Add ("Cond. Pagto")
    objGridInt.colColuna.Add ("Quant. Cotada")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Valor Presente (R$)")
    objGridInt.colColuna.Add ("Valor Item (R$)")
    objGridInt.colColuna.Add ("Alíquota IPI")
    objGridInt.colColuna.Add ("Alíquota ICMS")
    objGridInt.colColuna.Add ("Ped. Cotação")
    objGridInt.colColuna.Add ("Data Cotação")
    objGridInt.colColuna.Add ("Data Validade")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Data Necessidade")
    objGridInt.colColuna.Add ("Para Entrega")
    objGridInt.colColuna.Add ("Motivo da Escolha")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoCot.Name)
    objGridInt.colCampo.Add (ProdutoCot.Name)
    objGridInt.colCampo.Add (DescProdutoCot.Name)
    objGridInt.colCampo.Add (FornecedorCot.Name)
    objGridInt.colCampo.Add (FilialFornCot.Name)
    objGridInt.colCampo.Add (Moeda.Name)
    objGridInt.colCampo.Add (PrecoUnitarioCot.Name)
    objGridInt.colCampo.Add (TaxaForn.Name)
    objGridInt.colCampo.Add (Cotacao.Name)
    objGridInt.colCampo.Add (PrecoUnitarioReal.Name)
    objGridInt.colCampo.Add (CondPagtoCot.Name)
    objGridInt.colCampo.Add (QuantComprarMaxCot.Name)
    objGridInt.colCampo.Add (QuantComprarCot.Name)
    objGridInt.colCampo.Add (UMCot.Name)
    objGridInt.colCampo.Add (ValorPresenteCot.Name)
    objGridInt.colCampo.Add (ValorItem.Name)
    objGridInt.colCampo.Add (AliquotaIPI.Name)
    objGridInt.colCampo.Add (AliquotaICMS.Name)
    objGridInt.colCampo.Add (PedCotacao.Name)
    objGridInt.colCampo.Add (DataCotacao.Name)
    objGridInt.colCampo.Add (DataValidade.Name)
    objGridInt.colCampo.Add (PrazoEntrega.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (DataNecess.Name)
    objGridInt.colCampo.Add (QuantidadeEntrega.Name)
    objGridInt.colCampo.Add (MotivoEscolha.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoCot_Col = 1
    iGrid_ProdutoCot_Col = 2
    iGrid_DescProdutoCot_Col = 3
    iGrid_FornecedorCot_Col = 4
    iGrid_FilialFornCot_Col = 5
    iGrid_Moeda_Col = 6
    iGrid_PrecoUnitarioCot_Col = 7
    iGrid_TaxaForn_Col = 8
    iGrid_CotacaoMoeda_Col = 9
    iGrid_PrecoUnitario_RS_Col = 10
    iGrid_CondPagtoCot_Col = 11
    iGrid_QuantComprarMaxCot_Col = 12
    iGrid_QuantComprarCot_Col = 13
    iGrid_UMCot_Col = 14
    iGrid_ValorPresenteCot_Col = 15
    iGrid_ValorItem_Col = 16
    iGrid_AliquotaIPI_Col = 17
    iGrid_AliquotaICMS_Col = 18
    iGrid_PedidoCot_Col = 19
    iGrid_DataCotacaoCot_Col = 20
    iGrid_DataValidadeCot_Col = 21
    iGrid_PrazoEntregaCot_Col = 22
    iGrid_DataEntregaCot_Col = 23
    iGrid_DataNecessidadeCot_Col = 24
    iGrid_QuantidadeEntregaCot_Col = 25
    iGrid_MotivoEscolhaCot_Col = 26

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridCotacoes

    GridCotacoes.ColWidth(0) = 590
    
    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_COTACOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'    GridCotacoes.Width = 8295
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cotacoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Cotacoes:

    Inicializa_Grid_Cotacoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160958)

    End Select

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long, sProdutoFormatado As String, iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto, colSiglas As New Collection
Dim objClasseUM As New ClassClasseUM, iCodigo As Integer
Dim objUM As New ClassUnidadeDeMedida
Dim sUM As String, iIndice As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor, iSelecionado As Integer
Dim iProdutoPreenchido2 As Integer, sProdutoFormatado2 As String

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Fomata o Produto
    lErro = CF("Produto_Formata", GridCotacoes.TextMatrix(iLinha, iGrid_ProdutoCot_Col), sProdutoFormatado2, iProdutoPreenchido2)
    If lErro <> SUCESSO Then gError 63752
    
    'Pesquisa controle da coluna em questão
    Select Case objControl.Name
                
        'QuantComprarCot ou MotivoEscolha
        Case PrecoUnitarioCot.Name, QuantComprarCot.Name, MotivoEscolha.Name

            'Verifica se o Produto está preenchido
            If iProdutoPreenchido2 <> PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
                
                If objControl.Name = MotivoEscolha.Name And _
                   GridCotacoes.TextMatrix(iLinha, iGrid_MotivoEscolhaCot_Col) = MOTIVO_EXCLUSIVO_DESCRICAO Then
                   objControl.Enabled = False
                End If
                
            End If

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 63752, 63753, 63755
            'Erros tratados nas rotinas chamadas

        Case 63754
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 65593
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR", gErr, objFornecedor.sNomeReduzido, sProdutoFormatado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160959)

    End Select

    Exit Sub

End Sub

Private Sub Ordenacao_Carrega()
'preenche a combo de ordenacao e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordenação dos Bloqueios
    asOrdenacao(0) = "CotacaoProduto.Produto, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.CondPagtoPrazo"
    asOrdenacao(1) = "PedidoCotacao.Fornecedor,PedidoCotacao.Filial,CotacaoProduto.Produto,PedidoCotacao.CondPagtoPrazo"

    asOrdenacaoString(0) = "Produto"
    asOrdenacaoString(1) = "Fornecedor"

    'Carrega a Combobox Ordenacao
    For iIndice = 0 To 1

        Ordenacao.AddItem asOrdenacaoString(iIndice)
        Ordenacao.ItemData(Ordenacao.NewIndex) = iIndice

    Next

    'Seleciona a opção CodProduto + CondPagto + Fornecedor + Filial de seleção
    Ordenacao.ListIndex = 0

    gsOrdenacao = Ordenacao.Text

    Exit Sub

End Sub

Private Function Carrega_MotivoEscolha() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_MotivoEscolha

    'Lê o Código e o Nome de todo MotivoEscolha do BD
    lErro = CF("Cod_Nomes_Le", "Motivo", "Codigo", "Motivo", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 63796

    'Carrega a combo de Motivo Escolha com código e nome
    For Each objCodigoNome In colCodigoNome

        'Verifica se o MotivoEscolha é diferente de Exclusividade
        If objCodigoNome.iCodigo <> MOTIVO_EXCLUSIVO Then

            MotivoEscolha.AddItem objCodigoNome.sNome
            MotivoEscolha.ItemData(MotivoEscolha.NewIndex) = objCodigoNome.iCodigo

        End If

    Next

    Carrega_MotivoEscolha = SUCESSO

    Exit Function

Erro_Carrega_MotivoEscolha:

    Carrega_MotivoEscolha = gErr

    Select Case gErr

        Case 63681
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160962)

    End Select

    Exit Function

End Function

Private Sub Botao_Click(Index As Integer)

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgBox As VbMsgBoxResult
Dim objConcorrencia As New ClassConcorrencia
Dim colConcorrencia As New Collection

On Error GoTo Erro_BotaoExcluir_Click

    'Move dados da Concorrência para a memória
    objConcorrencia.lCodigo = StrParaLong(Concorrencia.Caption)
    objConcorrencia.iFilialEmpresa = giFilialEmpresa

    'Se nenhuma concorrência estiver selecionada, Erro
    If objConcorrencia.lCodigo = 0 Then gError 66913

    'Pede Confirmação para exclusão de Concorrência
    vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CONCORRENCIA", objConcorrencia.lCodigo)
    If vbMsgBox = vbNo Then Exit Sub

    GL_objMDIForm.MousePointer = vbHourglass

    'Exclui Concorrência
    lErro = CF("Concorrencia_Exclui", objConcorrencia)
    If lErro <> SUCESSO Then gError 66915

    Call Limpa_Tela_GeracaoPedCompraOV

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 66913
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CONCORRENCIAS_SELECIONADAS", gErr)

        Case 66915, 66998, 66999

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161078)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se o código da concorrencia esta preenchido
    If Len(Trim(Concorrencia.Caption)) = 0 Then gError 76084
    
    objConcorrencia.lCodigo = StrParaLong(Concorrencia.Caption)
    objConcorrencia.iFilialEmpresa = giFilialEmpresa
    
    'Lê a Concorrencia
    lErro = CF("Concorrencia_Le", objConcorrencia)
    If lErro <> SUCESSO And lErro <> 66788 Then gError 76079
    
    'Se não encontrou a concorrencia ==> erro
    If lErro = 66788 Then gError 76080
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Geracao Pedido Compra Avulsa", "CONCORTO.NumIntDoc = @NCONCORR", 1, "CONCORR", "NCONCORR", objConcorrencia.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76081
     
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 76079, 76081

        Case 76080
            Call Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_NAO_CADASTRADA", gErr, objConcorrencia.lCodigo)

        Case 76084
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONCORRENCIA_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160963)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 63697

    Call Limpa_Tela_GeracaoPedCompraOV

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 63697
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160964)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_GeracaoPedCompraOV()

Dim lErro As Long
Dim lConcorrencia As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_GeracaoPedCompraOV

    Call Limpa_Tela(Me)

    'Limpa os Grids da tela
    Call Grid_Limpa(objGridCotacoes)

    Call Calcula_TotalItens
    
    Set gcolItemConcorrencia = New Collection
    
    Concorrencia.Caption = ""
    
    Exit Sub

Erro_Limpa_Tela_GeracaoPedCompraOV:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160965)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridCotacoes
            Case GridCotacoes.Name

                lErro = Saida_Celula_GridCotacoes(objGridInt)
                If lErro <> SUCESSO Then gError 63700

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 63701

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 63699 To 63701
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160966)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridCotacoes(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridCotacoes que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCotacoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Produto
        Case iGrid_EscolhidoCot_Col
            lErro = Saida_Celula_Escolhido(objGridInt)
            If lErro <> SUCESSO Then gError 63737

        'UnidadeMed
        Case iGrid_QuantComprarCot_Col
            lErro = Saida_Celula_QuantComprarCot(objGridInt)
            If lErro <> SUCESSO Then gError 63738
        
        'Preço Unitário
        Case iGrid_PrecoUnitarioCot_Col
            lErro = Saida_Celula_PrecoUnitarioCot(objGridInt)
            If lErro <> SUCESSO Then gError 70481
        
        'Quantidade
        Case iGrid_MotivoEscolhaCot_Col
            lErro = Saida_Celula_MotivoEscolha(objGridInt)
            If lErro <> SUCESSO Then gError 63739

    End Select

    Saida_Celula_GridCotacoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridCotacoes:

    Saida_Celula_GridCotacoes = gErr

    Select Case gErr

        Case 63737 To 63739, 70481
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160968)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantComprarCot(objGridInt As AdmGrid) As Long
'Faz a saida de celula de QuantComprarCot

Dim lErro As Long
Dim dQuantidade As Double
Dim objCotItemConc As ClassCotacaoItemConc

On Error GoTo Erro_Saida_Celula_QuantComprarCot

     Set objGridInt.objControle = QuantComprarCot
    
    'Verifica se a QuantComprarCot esta preenchida
    If Len(Trim(QuantComprarCot.ClipText)) > 0 Then

        'Critica a quantidade
        lErro = Valor_Positivo_Critica(QuantComprarCot.Text)
        If lErro <> SUCESSO Then gError 63739

        dQuantidade = StrParaDbl(QuantComprarCot.Text)

        'Coloca a quantidade com o formato de estoque da tela
         QuantComprarCot.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63740
    
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    
    objCotItemConc.dQuantidadeComprar = dQuantidade
    GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar, "STANDARD")
    
    Call Calcula_TotalItens
    
    Saida_Celula_QuantComprarCot = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprarCot:

    Saida_Celula_QuantComprarCot = gErr

    Select Case gErr

        Case 63739, 63740
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160971)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitarioCot(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objCotItemConc As New ClassCotacaoItemConc
Dim dValorPresente As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Saida_Celula_PrecoUnitarioCot

    Set objGridInt.objControle = PrecoUnitarioCot

    'Se o Preço unitário estiver preenchido
    If Len(Trim(PrecoUnitarioCot.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(PrecoUnitarioCot.Text)
        If lErro <> SUCESSO Then gError 70482

    End If
        
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    
    objCotItemConc.dPrecoAjustado = StrParaDbl(PrecoUnitarioCot.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 70483

    'Se a condição de pagamento não for a vista
    If Codigo_Extrai(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_CondPagtoCot_Col)) <> COD_A_VISTA And PercentParaDbl(TaxaEmpresa.Caption) > 0 Then
        
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConc.sCondPagto)
        
        'Recalcula o Valor Presente
        lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
        If lErro <> SUCESSO Then gError 62736
        
        GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format(dValorPresente, gobjCOM.sFormatoPrecoUnitario)
        objCotItemConc.dValorPresente = dValorPresente
        
    ElseIf Codigo_Extrai(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_CondPagtoCot_Col)) = COD_A_VISTA Then
        GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format((StrParaDbl(PrecoUnitarioCot.Text)), gobjCOM.sFormatoPrecoUnitario) ' "Standard") 'Alterado por Wagner
        objCotItemConc.dValorPresente = dValorPresente
    End If
    
    GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar, "STANDARD")
    
    Call Calcula_TotalItens
    
    Saida_Celula_PrecoUnitarioCot = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitarioCot:

    Saida_Celula_PrecoUnitarioCot = gErr

    Select Case gErr

        Case 62736, 70482, 70483
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160972)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MotivoEscolha(objGridInt As AdmGrid) As Long
'Faz a saida de celula de MotivoEscolha

Dim lErro As Long
Dim iCodigo As Integer
Dim objCotItemConc As ClassCotacaoItemConc

On Error GoTo Erro_Saida_Celula_MotivoEscolha

    Set objGridInt.objControle = MotivoEscolha

    'Verifica se o MotivoEscolha está preenchido
    If Len(Trim(MotivoEscolha.Text)) > 0 Then

        'Verifica se MotivoEscolha não está selecionado
        If MotivoEscolha.ListIndex = -1 Then
                        
            If UCase(MotivoEscolha.Text) = UCase(MOTIVO_EXCLUSIVO_DESCRICAO) Then gError 62715
            
            'Seleciona o MotivoEscolha na combobox
            lErro = Combo_Item_Seleciona(MotivoEscolha)
            If lErro <> SUCESSO And lErro <> 12250 Then gError 63741

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63743

    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)

    objCotItemConc.sMotivoEscolha = GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_MotivoEscolhaCot_Col)

    Saida_Celula_MotivoEscolha = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoEscolha:

    Saida_Celula_MotivoEscolha = gErr

    Select Case gErr

        Case 62715
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_EXCLUSIVO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63741, 63743
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160973)

    End Select

    Exit Function

End Function

Private Sub Recarrega_Cotacoes(iLinha As Integer)

Dim lErro As Long
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objCotacaoItemConc As ClassCotacaoItemConc
    
    'Se não tiver nenhuma cotação selecionada sai da rotina
    Set objItemConcorrencia = gcolItemConcorrencia(iLinha)
    
    If objItemConcorrencia.colCotacaoItemConc.Count > 0 Then
    
        'Para cada cotação encontrada
        For Each objCotacaoItemConc In objItemConcorrencia.colCotacaoItemConc
            
            If objCotacaoItemConc.iSelecionada = DESMARCADO Then objCotacaoItemConc.dPrecoAjustado = objCotacaoItemConc.dPrecoUnitario
            
            'Todas as cotações podem aparecer
            objCotacaoItemConc.iSelecionada = MARCADO
            If objCotacaoItemConc.sMotivoEscolha = MOTIVO_EXCLUSIVO_DESCRICAO Then objCotacaoItemConc.sMotivoEscolha = ""
        
        Next
    ElseIf objItemConcorrencia.dQuantidade > 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_AUSENCIA_COTACOES_SELECAO")
    End If
    
    Exit Sub

End Sub

Private Function Saida_Celula_Escolhido(objGridInt As AdmGrid) As Long
'Faz a saida de celula de Escolhido

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Escolhido

    Set objGridInt.objControle = EscolhidoCot

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63738
    
    Saida_Celula_Escolhido = SUCESSO

    Exit Function

Erro_Saida_Celula_Escolhido:

    Saida_Celula_Escolhido = gErr

    Select Case gErr

        Case 63738
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160977)

    End Select

    Exit Function

End Function

Private Function Traz_Cotacoes_Tela(objProduto As ClassProduto, dQuantidade As Double, iLinha As Integer) As Long
'Traz para a tela as Cotacoes que envolvem o Produto passado com parametro

Dim lErro As Long, sProdutoFormatado As String
Dim iProdutoPreenchido As Integer, colCotacao As New Collection
Dim objItemConcorrencia As New ClassItemConcorrencia
Dim objItemCotItemConc As ClassCotacaoItemConc
Dim iIndice As Integer
Dim iTipoDestino As Integer
Dim lDestino As Long
Dim iFilialDestino As Integer

On Error GoTo Erro_Traz_Cotacoes_Tela

    Set colCotacao = New Collection
    
    lErro = Move_TipoDestino_Memoria(iTipoDestino, lDestino, iFilialDestino)
    If lErro <> SUCESSO And iTipoDestino <> -1 Then gError 62798
        
    If iTipoDestino > -1 Then
    
        'Lê as Cotacoes cujo Produto foi passado como parametro
        lErro = CF("Cotacoes_Produto_Le", colCotacao, objProduto, dQuantidade, iTipoDestino, lDestino, iFilialDestino)
        If lErro <> SUCESSO And lErro <> 63822 Then gError 68498
        
        Set objItemConcorrencia = gcolItemConcorrencia(iLinha)
        
        For iIndice = objItemConcorrencia.colCotacaoItemConc.Count To 1 Step -1
            objItemConcorrencia.colCotacaoItemConc.Remove iIndice
        Next
            
        For Each objItemCotItemConc In colCotacao
            objItemCotItemConc.dPrecoAjustado = objItemCotItemConc.dPrecoUnitario
            objItemConcorrencia.colCotacaoItemConc.Add objItemCotItemConc
        Next
            
        Call Escolher_Cotacoes(objItemConcorrencia, dQuantidade)
        
    End If
    
    Traz_Cotacoes_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Cotacoes_Tela:

    Traz_Cotacoes_Tela = gErr
    
    Select Case gErr

        Case 62798, 68498
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160978)

    End Select

    Exit Function

End Function

Private Sub BotaoPedCotacao_Click()
'Chama a tela PedidoCotacao a passando o código do PedidoCotacao como parametro

Dim objPedidoCotacao As New ClassPedidoCotacao
    
On Error GoTo Erro_BotaoPedCotacao_Click
    
    'Verifica se existe alguma linha do GridCotacoes selecionada
    If GridCotacoes.Row = 0 Then gError 89443

    'Coloca o código do Pedido de Cotacao selecionado
    objPedidoCotacao.lCodigo = StrParaLong(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_PedidoCot_Col))
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    'Chama a tela PedidoCotacao
    Call Chama_Tela("PedidoCotacao", objPedidoCotacao)

    Exit Sub

Erro_BotaoPedCotacao_Click:

    Select Case gErr
    
        Case 89443
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160980)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
'Gera o próximo número de Concorrencia

Dim lErro As Long
Dim lConcorrencia As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera o próximo código para Concorrencia
    lErro = CF("Concorrencia_Automatica", lConcorrencia)
    If lErro <> SUCESSO Then gError 76082

    'Coloca o código gerado na tela
    Concorrencia.Caption = lConcorrencia
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 76082
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160984)

    End Select

    Exit Sub

End Sub

Private Sub LabelConcorrencia_Click()
    '??? trazer a concorrencia
End Sub

Private Sub Ordenacao_Click()

Dim lErro As Long

On Error GoTo Erro_Ordenacao_Click

    If gsOrdenacao = "" Then Exit Sub

    If gsOrdenacao <> Ordenacao.Text Then
    
        gsOrdenacao = Ordenacao.Text
        
        'Devolve os elementos ordenados para o  GridCotacoes
        lErro = GridCotacoes_Preenche()
        If lErro <> SUCESSO Then gError 63809
        
    End If

    Exit Sub

Erro_Ordenacao_Click:

    Select Case gErr

        Case 63807 To 63809
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160990)

    End Select

    Exit Sub

End Sub

Sub Monta_Colecao_Campos_Cotacao(colCampos As Collection, iOrdenacao As Integer)

Dim objCotacaoItemConc As New ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia

    Select Case iOrdenacao

        Case 0

            colCampos.Add "sProduto"
            colCampos.Add "sCondPagto"
            colCampos.Add "sFornecedor"
            colCampos.Add "sFilial"

        Case 1

            colCampos.Add "sFornecedor"
            colCampos.Add "sFilial"
            colCampos.Add "sProduto"
            colCampos.Add "sCondPagto"

    End Select

End Sub

Private Function GridCotacoes_Preenche() As Long
'Preenche Grid de Cotações

Dim lErro As Long
Dim iIndiceMoeda As Integer
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim iIndice As Integer, iIndice2 As Integer
Dim colCampos As New Collection
Dim iCondPagto As Integer
Dim colGeracao As New Collection
Dim dValorPresente As Double
Dim colCotacaoSaida As New Collection
Dim sProdutoMascarado As String
Dim objCotItemConcAux As ClassCotacaoItemConcAux
Dim objItemCotItemConc As ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_GridCotacoes_Preenche
    
    Call Grid_Limpa(objGridCotacoes)
           
    For Each objItemConcorrencia In gcolItemConcorrencia

        'Coloca na coleção as cotações que aparecem na tela
         For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
                
            If objItemCotItemConc.iSelecionada = MARCADO Then
                Set objCotItemConcAux = New ClassCotacaoItemConcAux
                
                Set objCotItemConcAux.objCotacaoItemConc = objItemCotItemConc
                objCotItemConcAux.sCondPagto = objItemCotItemConc.sCondPagto
                objCotItemConcAux.sDescricao = objItemConcorrencia.sDescricao
                objCotItemConcAux.sFilial = objItemCotItemConc.sFilial
                objCotItemConcAux.sFornecedor = objItemCotItemConc.sFornecedor
                objCotItemConcAux.sProduto = objItemConcorrencia.sProduto
                objCotItemConcAux.dtDataNecessidade = objItemConcorrencia.dtDataNecessidade
            
                colGeracao.Add objCotItemConcAux
            End If
         Next
    Next
    
    'Carrega os campos base para a ordenação utilizados na rotina de ordenação
    Call Monta_Colecao_Campos_Cotacao(colCampos, Ordenacao.ListIndex)

    If colGeracao.Count > 0 Then
        lErro = Ordena_Colecao(colGeracao, colCotacaoSaida, colCampos)
        If lErro <> SUCESSO Then gError 63808
    End If
    
    Set colGeracao = colCotacaoSaida
    
    iIndice = 0
    
    For Each objCotItemConcAux In colGeracao

        iIndice = iIndice + 1
        GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col) = objCotItemConcAux.objCotacaoItemConc.iEscolhido

        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objCotItemConcAux.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 68358

        'Preenche o Produto com o ProdutoEnxuto
        ProdutoCot.PromptInclude = False
        ProdutoCot.Text = sProdutoMascarado
        ProdutoCot.PromptInclude = True
        
        GridCotacoes.TextMatrix(iIndice, iGrid_ProdutoCot_Col) = ProdutoCot.Text
        GridCotacoes.TextMatrix(iIndice, iGrid_DescProdutoCot_Col) = objCotItemConcAux.sDescricao
        GridCotacoes.TextMatrix(iIndice, iGrid_CondPagtoCot_Col) = objCotItemConcAux.objCotacaoItemConc.sCondPagto
        
        GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar)

        GridCotacoes.TextMatrix(iIndice, iGrid_UMCot_Col) = objCotItemConcAux.objCotacaoItemConc.sUMCompra
        GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitarioCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado, gobjCOM.sFormatoPrecoUnitario) ' "Standard")'Alterado por Wagner
        
        
        iCondPagto = Codigo_Extrai(objCotItemConcAux.objCotacaoItemConc.sCondPagto)
        
        'Se a condição de pagamento não for a vista
        If iCondPagto <> COD_A_VISTA And PercentParaDbl(TaxaEmpresa.Caption) > 0 Then
            
            Set objCondicaoPagto = New ClassCondicaoPagto
            objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConcAux.objCotacaoItemConc.sCondPagto)
            
            'Recalcula o Valor Presente
            lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
            If lErro <> SUCESSO Then gError 62733
            
            If objCotItemConcAux.objCotacaoItemConc.iMoeda <> MOEDA_REAL Then
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = dValorPresente * objCotItemConcAux.objCotacaoItemConc.dTaxa
            Else
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = dValorPresente
            End If
                
        Else
            
            If objCotItemConcAux.objCotacaoItemConc.iMoeda <> MOEDA_REAL Then
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario * objCotItemConcAux.objCotacaoItemConc.dTaxa
            Else
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario
            End If
                
        End If
                                          
        GridCotacoes.TextMatrix(iIndice, iGrid_ValorPresenteCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dValorPresente, "STANDARD")
        
        If objCotItemConcAux.objCotacaoItemConc.iMoeda <> MOEDA_REAL Then
            GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado * objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar * objCotItemConcAux.objCotacaoItemConc.dTaxa, "STANDARD")
        Else
            GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado * objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar, "STANDARD")
        End If
        
        GridCotacoes.TextMatrix(iIndice, iGrid_FornecedorCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFornecedor
        GridCotacoes.TextMatrix(iIndice, iGrid_FilialFornCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFilial
        GridCotacoes.TextMatrix(iIndice, iGrid_PedidoCot_Col) = objCotItemConcAux.objCotacaoItemConc.lPedCotacao
        If objCotItemConcAux.objCotacaoItemConc.dQuantEntrega > 0 Then GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeEntregaCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantEntrega)
        
        'Data da Cotacao
        If objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataCotacaoCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao, "dd/mm/yyyy")
        End If
    
''''        For iIndice2 = 0 To TipoTributacaoCot.ListCount - 1
''''            If objCotItemConcAux.objCotacaoItemConc.iTipoTributacao = TipoTributacaoCot.ItemData(iIndice2) Then
''''                GridCotacoes.TextMatrix(iIndice, iGrid_TipoTributacaoCot_Col) = TipoTributacaoCot.List(iIndice2)
''''                Exit For
''''            End If
''''        Next
        
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaIPI, "Percent")
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaICMS, "Percent")
        
        'Data de Validade
        If objCotItemConcAux.objCotacaoItemConc.dtDataValidade <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataValidadeCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dtDataValidade, "dd/mm/yyyy")
        End If

        'Prazo de Entrega
        If objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega <> 0 Then
            GridCotacoes.TextMatrix(iIndice, iGrid_PrazoEntregaCot_Col) = objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega
            GridCotacoes.TextMatrix(iIndice, iGrid_DataEntregaCot_Col) = Format(DateAdd("d", objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega, Date), "dd/mm/yyyy")
        End If

        'Data de Entrega
        If objCotItemConcAux.objCotacaoItemConc.dtDataEntrega <> DATA_NULA Then
        End If
                
        'Quantidade a comprar Máxima
        GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarMaxCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantCotada)

        'Motivo escolha
        GridCotacoes.TextMatrix(iIndice, iGrid_MotivoEscolhaCot_Col) = objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha
        
        If objCotItemConcAux.dtDataNecessidade <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataNecessidadeCot_Col) = Format(objCotItemConcAux.dtDataNecessidade, "dd/mm/yyyy")
        End If
        
        'Moeda
        For iIndiceMoeda = 0 To Moeda.ListCount - 1
            If Moeda.ItemData(iIndiceMoeda) = objCotItemConcAux.objCotacaoItemConc.iMoeda Then
                GridCotacoes.TextMatrix(iIndice, iGrid_Moeda_Col) = Moeda.List(iIndiceMoeda)
                Exit For
            End If
        Next
        
        'TaxaForn
        GridCotacoes.TextMatrix(iIndice, iGrid_TaxaForn_Col) = IIf(objCotItemConcAux.objCotacaoItemConc.dTaxa = 0, "", Format(objCotItemConcAux.objCotacaoItemConc.dTaxa, "#.0000"))
        
        If Moeda.ItemData(iIndiceMoeda) <> MOEDA_REAL Then
            
            'Cotacao
            objCotacaoMoeda.iMoeda = Moeda.ItemData(iIndiceMoeda)
            objCotacaoMoeda.dtData = gdtDataHoje
            
            lErro = CF("CotacaoMoeda_Le", objCotacaoMoeda)
            If lErro <> SUCESSO And lErro <> 80267 Then gError 108983
            
            If objCotacaoMoeda.dValor > 0 Then GridCotacoes.TextMatrix(iIndice, iGrid_CotacaoMoeda_Col) = Format(objCotacaoMoeda.dValor, "#.0000")
            
            'Preco unitario R$
            GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitario_RS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario * objCotItemConcAux.objCotacaoItemConc.dTaxa, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        Else
            'Preco unitario R$
            GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitario_RS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner

            
        End If
        
        objGridCotacoes.iLinhasExistentes = objGridCotacoes.iLinhasExistentes + 1
        
    Next

    Call Grid_Refresh_Checkbox(objGridCotacoes)
    
    Call Calcula_TotalItens
    
    Exit Function

Erro_GridCotacoes_Preenche:

    Select Case gErr

        Case 62733, 63808, 68358, 108983
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160991)

    End Select

    Exit Function

End Function

Private Sub GridCotacoes_Click()

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridCotacoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)
    End If

End Sub

Private Sub GridCotacoes_GotFocus()
    Call Grid_Recebe_Foco(objGridCotacoes)
End Sub

Private Sub GridCotacoes_EnterCell()
    Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)
End Sub

Private Sub GridCotacoes_LeaveCell()
    Call Saida_Celula(objGridCotacoes)
End Sub

Private Sub GridCotacoes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridCotacoes)

End Sub

Private Sub GridCotacoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCotacoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)
    End If
    
    Exit Sub
    
End Sub

Private Sub GridCotacoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridCotacoes)
End Sub

Private Sub GridCotacoes_RowColChange()
    Call Grid_RowColChange(objGridCotacoes)
End Sub

Private Sub GridCotacoes_Scroll()
    Call Grid_Scroll(objGridCotacoes)
End Sub

Private Sub Escolhido_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Escolhido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub Escolhido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub Escolhido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = EscolhidoCot
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantComprarCot_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantComprarCot_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub QuantComprarCot_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub QuantComprarCot_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = QuantComprarCot
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoUnitarioCot_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoUnitarioCot_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub PrecoUnitarioCot_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub PrecoUnitarioCot_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = PrecoUnitarioCot
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub DataValidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = DataValidade
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrazoEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrazoEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub PrazoEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub PrazoEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = PrazoEntrega
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = DataEntrega
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataNecess_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataNecess_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub DataNecess_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub DataNecess_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = DataNecess
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub QuantidadeEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub QuantidadeEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = QuantidadeEntrega
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MotivoEscolha_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MotivoEscolha_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub MotivoEscolha_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub MotivoEscolha_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = MotivoEscolha
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Pedidos de Compra para Orçamento de Venda"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoPedCompraOV"

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
Private Sub Comprador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Comprador, Source, X, Y)
End Sub

Private Sub Comprador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Comprador, Button, Shift, X, Y)
End Sub

Private Sub Label45_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label45, Source, X, Y)
End Sub

Private Sub Label45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label45, Button, Shift, X, Y)
End Sub

Private Sub TaxaEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TaxaEmpresa, Source, X, Y)
End Sub

Private Sub TaxaEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TaxaEmpresa, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Concorrencia_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Concorrencia, Source, X, Y)
End Sub

Private Sub Concorrencia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Concorrencia, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub BotaoGravaConcorrencia_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravaConcorrencia_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 63761
   
    Exit Sub

Erro_BotaoGravaConcorrencia_Click:

    Select Case gErr

        Case 63761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160995)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGeraPedidos_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGeraPedidos_Click

    'Grava a Geracao de Pedido de Compra
    lErro = Gravar_Pedidos()
    If lErro <> SUCESSO Then gError 63696

    iAlterado = 0

    Exit Sub

Erro_BotaoGeraPedidos_Click:

    Select Case gErr

        Case 63696
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 160996)

    End Select

    Exit Sub

End Sub

Function Gravar_Pedidos() As Long

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim colPedidoCompra As New Collection

On Error GoTo Erro_Gravar_Pedidos
    
    GL_objMDIForm.MousePointer = vbHourglass
            
    'Recolhe os dados da tela
    lErro = Move_Concorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63751

    'Atualiza a Concorrencia no Banco de Dados
    lErro = CF("Concorrencia_Grava", objConcorrencia)
    If lErro <> SUCESSO Then gError 63754
    
    'Carrega em colPedidoCompras os Pedidos de Compra gerados a partir de diferentes Fornecedores e FiliaisFornecedores
    lErro = Carrega_Dados_Pedidos(objConcorrencia, colPedidoCompra)
    If lErro <> SUCESSO Then gError 63753

    'Grava o Pedido de Compras
    lErro = CF("PedCompra_Concorrencia_Grava", objConcorrencia, colPedidoCompra)
    If lErro <> SUCESSO Then gError 63755

    '#####################################
    'Inserido por Wagner
    If colPedidoCompra.Count > 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_CODIGO_PEDCOMPRA_GRAVADO", colPedidoCompra.Item(1).lCodigo, colPedidoCompra.Item(colPedidoCompra.Count).lCodigo)
    End If
    '#####################################

    'Limpa a tela
    Call Limpa_Tela_GeracaoPedCompraOV

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Pedidos = SUCESSO

    Exit Function

Erro_Gravar_Pedidos:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Pedidos = gErr

    Select Case gErr

        Case 63751, 63753, 63754, 63755, 70499
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 160997)

    End Select

    Exit Function

End Function

Function Valida_Quantidade(objItemConcorrencia As ClassItemConcorrencia, iItem As Integer) As Long
'Verifica se os campos da tela foram preenchidos corretamente

Dim lErro As Long
Dim dQuantidade As Double
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim objCotItemConc As ClassCotacaoItemConc
Dim dQuantComprar As Double
Dim iTot As Integer

On Error GoTo Erro_Valida_Quantidade

    If objItemConcorrencia.colCotacaoItemConc.Count = 0 Then gError 63759
    
    iTot = 0

    objProduto.sCodigo = objItemConcorrencia.sProduto

    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 23080 Then gError 62712
    If lErro <> SUCESSO Then gError 62713 'não encontrou

    'Recolhe a quantidade do grid
    dQuantidade = objItemConcorrencia.dQuantidade

    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemConcorrencia.sUM, objProduto.sSiglaUMCompra, dFator)
    If lErro <> SUCESSO Then gError 62714

    dQuantidade = dQuantidade * dFator

    dQuantComprar = 0

    'Percorre as cotações
    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
        If objCotItemConc.iSelecionada = MARCADO And objCotItemConc.iEscolhido = MARCADO Then
            iTot = iTot + 1
            dQuantComprar = dQuantComprar + objCotItemConc.dQuantidadeComprar
            If objCotItemConc.dPrecoAjustado = 0 Then gError 70498
        End If
    Next
    
    If iTot = 0 Then gError 63759

    If Abs(Formata_Estoque(dQuantComprar - dQuantidade)) >= QTDE_ESTOQUE_DELTA Then gError 63811

    Valida_Quantidade = SUCESSO

    Exit Function

Erro_Valida_Quantidade:

    Valida_Quantidade = gErr

    Select Case gErr

        Case 62712, 62714
        
        Case 62713
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 63759
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_VINCULADO_ITEMCOTACAO", gErr, iItem)

        Case 63811
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOTACAO_DIFERENTE_QUANTCOMPRAR", gErr, objProduto.sCodigo)

        Case 70498
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOUNITARIO_ITEMCOTACAO_NAO_PREENCHIDO", gErr, iItem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 160998)

    End Select

    Exit Function
    
End Function

Private Function Move_Concorrencia_Memoria(objConcorrencia As ClassConcorrencia) As Long
'Recolhe os dados da tela e armazena em objConcorrencia

Dim lErro As Long
Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Concorrencia_Memoria
        
        
    objConcorrencia.iTipoDestino = TIPO_DESTINO_EMPRESA
    objConcorrencia.iFilialDestino = giFilialEmpresa
    
    objConcorrencia.dTaxaFinanceira = PercentParaDbl(TaxaEmpresa.Caption)
    
    'verifica se o código da concorrencia está preenchido
    If Len(Trim(Concorrencia.Caption)) = 0 Then gError 76083
    
    objConcorrencia.lCodigo = StrParaLong(Concorrencia.Caption)

    objUsuario.sNomeReduzido = Comprador.Caption

    'Lê o usuario a partir do nome reduzido
    lErro = CF("Usuario_Le_NomeRed", objUsuario)
    If lErro <> SUCESSO And lErro <> 57269 Then gError 63774
    If lErro = 57269 Then gError 63777

    objComprador.sCodUsuario = objUsuario.sCodUsuario

    'Lê o comprador a partir do codUsuario
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63820

    'Se não encontrou o comprador==>erro
    If lErro = 50059 Then gError 70490

    objConcorrencia.iComprador = objComprador.iCodigo
    objConcorrencia.iFilialEmpresa = giFilialEmpresa
    objConcorrencia.dtData = gdtDataAtual
    objConcorrencia.sDescricao = Descricao.Text

    'Move os itens da concorrência para a memória
    lErro = Move_ItensConcorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63776

    Move_Concorrencia_Memoria = SUCESSO

    Exit Function

Erro_Move_Concorrencia_Memoria:

    Move_Concorrencia_Memoria = gErr

    Select Case gErr

        Case 63746
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)

        Case 63749
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_GRID", gErr)

        Case 63777
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_INEXISTENTE", gErr, objUsuario.sNomeReduzido)
        
        Case 63820, 63774, 63775, 63776
            'Erros tratados nas rotinas chamadas

        Case 70490
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 70491
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 76083
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONCORRENCIA_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 160999)

    End Select

    Exit Function

End Function

Function Move_ItensConcorrencia_Memoria(objConcorrencia As ClassConcorrencia) As Long
'Move os dados dos Itens da Concorrência (GridProdutos1) para a memória

Dim lErro As Long
Dim iItem As Integer
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_Move_ItensConcorrencia_Memoria
            
    iItem = 0
    'Para cada item de concorrencia
    For Each objItemConcorrencia In gcolItemConcorrencia
        
        iItem = iItem + 1
        'verifica se a quantidade foi preenchida
        If objItemConcorrencia.dQuantidade = 0 Then gError 63750
        
        'valida a quantidade do item de concorrência
        lErro = Valida_Quantidade(objItemConcorrencia, iItem)
        If lErro <> SUCESSO Then gError 70492
    
    Next
    
    Set objConcorrencia.colItens = gcolItemConcorrencia

    Move_ItensConcorrencia_Memoria = SUCESSO

    Exit Function

Erro_Move_ItensConcorrencia_Memoria:

    Move_ItensConcorrencia_Memoria = gErr

    Select Case gErr

        Case 63750
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr)

        Case 70492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161000)

    End Select

    Exit Function

End Function

Function Carrega_Dados_Pedidos(objConcorrencia As ClassConcorrencia, colPedidoCompras As Collection) As Long
'Carrega em colPedidoCompras os Pedidos de Compra gerados a partir de diferentes Fornecedores e FiliaisFornecedores

Dim lErro As Long
Dim objCotItemConc As ClassCotacaoItemConc
Dim objPedidoCompra As ClassPedidoCompras
Dim objItemConcorrencia As ClassItemConcorrencia
Dim bAchou As Boolean
Dim iIndice As Integer
Dim objFornecedor As New ClassFornecedor
Dim objItemPC As ClassItemPedCompra
Dim lNumIntOriginal As Long
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colPedCompraGeral As New Collection
Dim colPedCompraExclu As New Collection
Dim objItemPedCotacao As ClassItemPedCotacao
Dim objItemCotacao As ClassItemCotacao
Dim colItensCotacao As New Collection
Dim dTotalItens As Double

On Error GoTo Erro_Carrega_Dados_Pedidos
    
    'Para cada item da concorrência
    For Each objItemConcorrencia In objConcorrencia.colItens
        
        If objItemConcorrencia.lFornecedor > 0 And objItemConcorrencia.iFilial > 0 Then
            Set colPedidoCompras = colPedCompraExclu
        Else
            Set colPedidoCompras = colPedCompraGeral
        End If
        
        'Para cada cotação utilizada
        For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc

            If objCotItemConc.iEscolhido = MARCADO Then
                
                'Lê o Fornecedor
                objFornecedor.sNomeReduzido = objCotItemConc.sFornecedor
                
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 63799
    
                'Se não encontrou ==> erro
                If lErro = 6681 Then gError 63800
                
                iIndice = 0
                bAchou = False
                                
                'Verifica se já foi criado pedido de compra com
                'o fornecedor, a Filial e a condPagto da cotação
                For Each objPedidoCompra In colPedidoCompras
                    iIndice = iIndice + 1
                    
                    If objPedidoCompra.lFornecedor = objFornecedor.lCodigo And _
                       objPedidoCompra.iFilial = Codigo_Extrai(objCotItemConc.sFilial) And _
                       objPedidoCompra.iCondicaoPagto = Codigo_Extrai(objCotItemConc.sCondPagto) Then
                       
                        bAchou = True
                        Exit For
                    End If
                Next
                
                'Se já existe pedido
                If bAchou Then
                    'seleciona o pedido
                    Set objPedidoCompra = colPedidoCompras(iIndice)
                'Senão
                Else
                    'Cria um novo Pedido de compras com as características na cotação
                    Set objPedidoCompra = New ClassPedidoCompras
                    
                    'Guarda o número do pedido de cotação do item de cotação
                    objPedidoCompra.lPedCotacao = objCotItemConc.lPedCotacao
                    
                    objPedidoCompra.iFilialEmpresa = giFilialEmpresa
                    objPedidoCompra.dtData = gdtDataAtual
                    objPedidoCompra.dtDataAlteracao = DATA_NULA
                    objPedidoCompra.dtDataBaixa = DATA_NULA
                    objPedidoCompra.dtDataEmissao = DATA_NULA
                    objPedidoCompra.dtDataEnvio = DATA_NULA
                    objPedidoCompra.dValorProdutos = 0
                    objPedidoCompra.dValorTotal = 0
                    objPedidoCompra.iComprador = objConcorrencia.iComprador
                    objPedidoCompra.iCondicaoPagto = Codigo_Extrai(objCotItemConc.sCondPagto)
                    objPedidoCompra.iFilial = Codigo_Extrai(objCotItemConc.sFilial)
                    objPedidoCompra.iFilialDestino = objConcorrencia.iFilialDestino
                    objPedidoCompra.iTipoDestino = objConcorrencia.iTipoDestino
                    objPedidoCompra.lFornCliDestino = objConcorrencia.lFornCliDestino
                    objPedidoCompra.lFornecedor = objFornecedor.lCodigo
                    objPedidoCompra.sTipoFrete = TIPO_FOB
                    objPedidoCompra.iMoeda = objCotItemConc.iMoeda
                    objPedidoCompra.dTaxa = objCotItemConc.dTaxa
                    
                    colPedidoCompras.Add objPedidoCompra
                End If
          
                'cria um novo item para o pedido de compras
                Set objItemPC = New ClassItemPedCompra
                      
                'Se o pedido de cotação utilizado no pedido não for o mesmo
                If objPedidoCompra.lPedCotacao <> objCotItemConc.lPedCotacao Then objPedidoCompra.lPedCotacao = 0
      
                objItemPC.dPrecoUnitario = objCotItemConc.dPrecoAjustado
                objItemPC.dQuantidade = objCotItemConc.dQuantidadeComprar
                objItemPC.dtDataLimite = objItemConcorrencia.dtDataNecessidade
                objItemPC.iStatus = ITEM_PED_COMPRAS_ABERTO
                objItemPC.iTipoOrigem = TIPO_ORIGEM_COTACAOITEMCONC
                objItemPC.sDescProduto = objItemConcorrencia.sDescricao
                objItemPC.sProduto = objItemConcorrencia.sProduto
                objItemPC.sUM = objCotItemConc.sUMCompra

                objItemPC.lNumIntOrigem = objCotItemConc.lNumIntDoc
                            
                objPedidoCompra.colItens.Add objItemPC
                
                'Adiciona o item de cotação na coleção de itens de cotacao
                lErro = colItensCotacao_Adiciona(objCotItemConc.lItemCotacao, colItensCotacao)
                If lErro <> SUCESSO Then gError 62726
            End If
        Next
    Next
   
    Set colPedidoCompras = New Collection

    'Gera uma única colecao de Pedidos de Compra, a partir das colecoes colPedCompraExclu e colPedCompraGeral já criadas
    lErro = PedidoCompra_Define_Colecao(colPedCompraExclu, colPedCompraGeral, colPedidoCompras)
    If lErro <> SUCESSO Then gError 76246
    
    'Aproveita os valores das cotações utilizadas
    'caso o pedido tenha sido gerado com itens da mesma cotação
    lErro = Atualiza_Valores_Pedido(colPedidoCompras, colItensCotacao)
    If lErro <> SUCESSO Then gError 62727
        
    Carrega_Dados_Pedidos = SUCESSO

    Exit Function

Erro_Carrega_Dados_Pedidos:

    Carrega_Dados_Pedidos = gErr

    Select Case gErr

        Case 63799, 70484, 62726, 62727, 76246
            'Erros tratados nas rotinas chamadas

        Case 63800, 70485
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161001)

    End Select

    Exit Function

End Function

Function colItensCotacao_Adiciona(lItemCotacao As Long, colItensCotacao As Collection) As Long

Dim objItemCotacao As ClassItemCotacao
Dim bAchou As Boolean
Dim lErro As Long

On Error GoTo Erro_colItensCotacao_Adiciona

    bAchou = False
    For Each objItemCotacao In colItensCotacao
        If objItemCotacao.lNumIntDoc = lItemCotacao Then
            bAchou = True
            Exit For
        End If
    Next
    
    If Not bAchou Then
        Set objItemCotacao = New ClassItemCotacao
        
        objItemCotacao.lNumIntDoc = lItemCotacao
        
        lErro = CF("ItemCotacao_Le", objItemCotacao)
        If lErro <> SUCESSO Then gError 62725
        
        colItensCotacao.Add objItemCotacao, CStr(objItemCotacao.lNumIntDoc)

    End If
    
    colItensCotacao_Adiciona = SUCESSO
    
    Exit Function

Erro_colItensCotacao_Adiciona:

    colItensCotacao_Adiciona = Err
    
    Select Case gErr
    
        Case 62725
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161002)
    
    End Select

    Exit Function

End Function

Function Atualiza_Valores_Pedido(colPedidoCompras As Collection, colItensCotacao As Collection) As Long
'Aproveita os valores das cotações utilizadas
'caso o pedido tenha sido gerado com itens da mesma cotação
         
Dim lErro As Long
Dim objItemPC As ClassItemPedCompra
Dim objItemCotacao As ClassItemCotacao
Dim objCotItemConc As ClassCotacaoItemConc
Dim objPedidoCompra As ClassPedidoCompras
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objItemConcorrencia As ClassItemConcorrencia
    
On Error GoTo Erro_Atualiza_Valores_Pedido

    'Atualiza o valor dos produtos no pedido de venda
    For Each objPedidoCompra In colPedidoCompras

        'Zera os acumuladores dos valores
        objPedidoCompra.dValorDesconto = 0
        objPedidoCompra.dValorFrete = 0
        objPedidoCompra.dValorIPI = 0
        objPedidoCompra.dValorProdutos = 0
        objPedidoCompra.dValorSeguro = 0

        'Se o pedido foi gerado com itens de um só ped Cotação
        If objPedidoCompra.lPedCotacao <> 0 Then

            objPedidoCotacao.lCodigo = objPedidoCompra.lPedCotacao
            objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
            
            'Lê o Pedido de Cotacao
            lErro = CF("PedidoCotacao_Le", objPedidoCotacao)
            If lErro <> SUCESSO And lErro <> 53670 Then gError 62728
            If lErro <> SUCESSO Then gError 62729 'Não encontrou
            
            objPedidoCompra.sTipoFrete = objPedidoCotacao.iTipoFrete
            
            'Para cada item de pedido de compra
            For Each objItemPC In objPedidoCompra.colItens
                
                'Busca nos itens de concorrencia os dados do item de cotação
                For Each objItemConcorrencia In gcolItemConcorrencia
                    
                    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
                        
                        'Se a cotação foi a utilizada pelo item de Pedido de Compras
                        If objItemPC.lNumIntOrigem = objCotItemConc.lNumIntDoc Then

                            'Guarda o número do item de cotação
                            Set objItemCotacao = colItensCotacao(CStr(objCotItemConc.lItemCotacao))
                                                 
                            objPedidoCompra.dOutrasDespesas = objPedidoCompra.dOutrasDespesas + (objItemCotacao.dOutrasDespesas * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorDesconto = objPedidoCompra.dValorDesconto + (objItemCotacao.dValorDesconto * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorFrete = objPedidoCompra.dValorFrete + (objItemCotacao.dValorFrete * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorSeguro = objPedidoCompra.dValorSeguro + (objItemCotacao.dValorSeguro * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objItemPC.dAliquotaICMS = objItemCotacao.dAliquotaICMS
                            objItemPC.dAliquotaIPI = objItemCotacao.dAliquotaIPI
                            objItemPC.dValorIPI = (objItemCotacao.dValorIPI * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorIPI = objPedidoCompra.dValorIPI + objItemPC.dValorIPI
                            objItemPC.lObservacao = objItemCotacao.lObservacao
                        End If
                    Next
                Next
            Next
        End If
        
        'Atualiza o valor dos produtos no Pedido de compras
        For Each objItemPC In objPedidoCompra.colItens
            objPedidoCompra.dValorProdutos = objPedidoCompra.dValorProdutos + (objItemPC.dPrecoUnitario * objItemPC.dQuantidade)
        Next
        
        objPedidoCompra.dValorTotal = objPedidoCompra.dValorFrete + objPedidoCompra.dValorIPI + objPedidoCompra.dValorProdutos + objPedidoCompra.dValorSeguro - objPedidoCompra.dValorDesconto
    Next
    
    Atualiza_Valores_Pedido = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Valores_Pedido:

    Atualiza_Valores_Pedido = gErr
    
    Select Case gErr
    
        Case 62728
    
        Case 62729
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161003)
            
    End Select
    
    Exit Function

End Function

Private Sub Localiza_ItemCotacao(objCotItemConc As ClassCotacaoItemConc, iLinha As Integer)
    
Dim sFornecedor As String
Dim sFilial As String
Dim sMotivo As String
Dim sProduto As String
Dim sCondPagto As String
Dim iIndice As Integer
Dim iItemConc As Integer
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objCotItemConc2 As ClassCotacaoItemConc
Dim iMoeda As Integer
Dim iProdutoPreenchido As Integer, sProdutoFormatado As String
    
    'Recolhe os campos que amarram  uma cotação na tela
    sMotivo = GridCotacoes.TextMatrix(iLinha, iGrid_MotivoEscolhaCot_Col)
    
    Call CF("Produto_Formata", GridCotacoes.TextMatrix(iLinha, iGrid_ProdutoCot_Col), sProdutoFormatado, iProdutoPreenchido)
    
    sCondPagto = GridCotacoes.TextMatrix(iLinha, iGrid_CondPagtoCot_Col)
    sFornecedor = GridCotacoes.TextMatrix(iLinha, iGrid_FornecedorCot_Col)
    sFilial = GridCotacoes.TextMatrix(iLinha, iGrid_FilialFornCot_Col)
    
    For iIndice = 0 To Moeda.ListCount - 1
        If Moeda.List(iIndice) = GridCotacoes.TextMatrix(iLinha, iGrid_Moeda_Col) Then
            iMoeda = Moeda.ItemData(iIndice)
            Exit For
        End If
    Next
    
    For Each objItemConcorrencia In gcolItemConcorrencia
        If sProdutoFormatado = objItemConcorrencia.sProduto Then Exit For
    Next
    
    'Busca dentro das cotações do item de concorrência a cotação em questão
    For Each objCotItemConc2 In objItemConcorrencia.colCotacaoItemConc
        
        If objCotItemConc2.sFornecedor = sFornecedor And _
            objCotItemConc2.sFilial = sFilial And objCotItemConc2.sCondPagto = sCondPagto And _
            objCotItemConc2.iMoeda = iMoeda Then
            
            Set objCotItemConc = objCotItemConc2
            Exit For
        
        End If
    Next
    
End Sub



Private Sub EscolhidoCot_Click()

Dim objCotItemConc As ClassCotacaoItemConc

    iAlterado = REGISTRO_ALTERADO
   
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    
    objCotItemConc.iEscolhido = GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_EscolhidoCot_Col)
    
    Call Calcula_TotalItens

End Sub

Private Sub EscolhidoCot_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub EscolhidoCot_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub EscolhidoCot_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = EscolhidoCot
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Escolher_Cotacoes(objItemConcorrencia As ClassItemConcorrencia, dQuantComprar As Double)
'recebe a coleção de Itens de cotação lida do BD e Escolhe para
'o usuário aquelas que possuem melhor preço ,ou melhor preco + prazo entrega
'como defaut
Dim dMelhorPreco As Double
Dim objCotItemConcMelhor As ClassCotacaoItemConc
Dim objCotItemConc As ClassCotacaoItemConc
Dim dQuantidade As Double
Dim dValorPresente As Double
Dim lErro As Long
Dim dTaxa As Double
Dim dValorPresenteReal As Double
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim iIndice As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Escolher_Cotacoes
    
    dMelhorPreco = 0
    dQuantidade = dQuantComprar
      
    'Se está amarrado com for e filial --> sai
    If objItemConcorrencia.lFornecedor > 0 And objItemConcorrencia.iFilial > 0 Then Exit Sub
        
    If objItemConcorrencia.colCotacaoItemConc.Count = 0 Then Exit Sub
    
    Set objCotItemConcMelhor = objItemConcorrencia.colCotacaoItemConc(1)
    
    For iIndice = 1 To objItemConcorrencia.colCotacaoItemConc.Count
        
        Set objCotItemConcMelhor = objItemConcorrencia.colCotacaoItemConc(iIndice)
    
        If objCotItemConcMelhor.iMoeda <> MOEDA_REAL Then
            If objCotItemConcMelhor.dTaxa > 0 Then
                dTaxa = objCotItemConcMelhor.dTaxa
                Exit For
            Else
                objCotacaoMoeda.iMoeda = objCotItemConcMelhor.iMoeda
                objCotacaoMoeda.dtData = gdtDataHoje
                
                lErro = CF("CotacaoMoeda_Le", objCotacaoMoeda)
                If lErro <> SUCESSO And lErro <> 80267 Then gError 108983
                If lErro = SUCESSO Then
                    dTaxa = objCotItemConcMelhor.dTaxa
                    Exit For
                End If
            End If
        Else
            dTaxa = 1
            Exit For
        End If
    Next
    
    dMelhorPreco = objCotItemConcMelhor.dPrecoUnitario * dTaxa
    
    Set objCondicaoPagto = New ClassCondicaoPagto
    objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConcMelhor.sCondPagto)
    
    'Recalcula o Valor Presente
    lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConcMelhor.dPrecoAjustado * dTaxa, PercentParaDbl(TaxaEmpresa.Caption), dValorPresenteReal, gdtDataAtual)
    If lErro <> SUCESSO Then gError 62733
    
    objCotItemConcMelhor.iSelecionada = MARCADO
    objCotItemConcMelhor.iEscolhido = MARCADO
    objCotItemConcMelhor.sMotivoEscolha = MOTIVO_MELHORPRECO_DESCRICAO
    
    'Para cada cotação do item
    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
        
        Set objCondicaoPagto = New ClassCondicaoPagto
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConc.sCondPagto)
        
        'Recalcula o Valor Presente
        lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
        If lErro <> SUCESSO Then gError 62733

        'Calcula o valor presente
        objCotItemConc.dValorPresente = dValorPresente

        If objCotItemConc.iMoeda <> MOEDA_REAL Then
            If objCotItemConc.dTaxa > 0 Then
                dTaxa = objCotItemConc.dTaxa
            Else
                objCotacaoMoeda.iMoeda = objCotItemConc.iMoeda
                objCotacaoMoeda.dtData = gdtDataHoje
                
                lErro = CF("CotacaoMoeda_Le", objCotacaoMoeda)
                If lErro <> SUCESSO And lErro <> 80267 Then gError 108983

                dTaxa = objCotItemConc.dTaxa
            End If
        Else
            dTaxa = 1
        End If
        
        dValorPresenteReal = dValorPresente * dTaxa
        
        'Se a Cotação for em Real ou se for em outra moeda para a qual _
         a Cotação esteja informada então pode-se analisar qual é a _
         melhor opção de preço convertendo todos para Real
        If ((objCotItemConc.iMoeda = MOEDA_REAL) Or (objCotItemConc.iMoeda <> MOEDA_REAL And dTaxa > 0)) Then

            'Se o valor presente é melhor que o menor preço até agora
            If (dValorPresenteReal < dMelhorPreco) Then
    
                objCotItemConcMelhor.sMotivoEscolha = ""
                objCotItemConcMelhor.iEscolhido = DESMARCADO
                objCotItemConcMelhor.iSelecionada = DESMARCADO
                
                'Guarda essa cotação como a de melhor preço
                dMelhorPreco = dValorPresenteReal
                
                Set objCotItemConcMelhor = objCotItemConc
                
                objCotItemConcMelhor.sMotivoEscolha = MOTIVO_MELHORPRECO_DESCRICAO
                objCotItemConcMelhor.iEscolhido = MARCADO
                objCotItemConcMelhor.iSelecionada = MARCADO
    
            'Se o valor for igual ao da cotação de melhor preço
            ElseIf dValorPresenteReal = dMelhorPreco Then
    
                If objCotItemConc.iPrazoEntrega <> 0 And objCotItemConcMelhor.iPrazoEntrega <> 0 Then
                    'Escolhe a cotação com o melhor prazo de entrega
                    If objCotItemConc.iPrazoEntrega < objCotItemConcMelhor.iPrazoEntrega Then
                                                
                        objCotItemConcMelhor.sMotivoEscolha = ""
                        objCotItemConcMelhor.iEscolhido = DESMARCADO
                        objCotItemConcMelhor.iSelecionada = DESMARCADO
                        
                        dMelhorPreco = objCotItemConc.dValorPresente
                        Set objCotItemConcMelhor = objCotItemConc
                        objCotItemConcMelhor.sMotivoEscolha = MOTIVO_PRECO_PRAZO_DESCRICAO
                        objCotItemConcMelhor.iEscolhido = MARCADO
                        objCotItemConcMelhor.iSelecionada = MARCADO
                    End If
                End If
            Else
                objCotItemConc.iEscolhido = DESMARCADO
            End If
        End If
    Next
    
    Exit Sub
    
Erro_Escolher_Cotacoes:

    Select Case gErr
    
        Case 62733
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161004)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub Calcula_TotalItens()

Dim dTotalItens As Double
Dim iIndice As Integer
    
    dTotalItens = 0
    
    For iIndice = 1 To objGridCotacoes.iLinhasExistentes
        If StrParaInt(GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col)) = MARCADO Then
            dTotalItens = dTotalItens + StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col))
        End If
    Next

    TotalItens.Caption = Format(dTotalItens, "STANDARD")
    
    Exit Sub

End Sub

Function PedidoCompra_Define_Colecao(colPedCompraExclu As Collection, colPedCompraGeral As Collection, colPedidoCompras As Collection) As Long
'A partir das colecoes de Pedidos de Compra Exclusivos e de Pedidos de Compra Não Exclusivos,
'define uma coleção única para todos os Pedidos de Compra criados

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim bProdutoIgual As Boolean
Dim objPCGeral As New ClassPedidoCompras
Dim objPCExclu As New ClassPedidoCompras
Dim objItemPCExclu As New ClassItemPedCompra
Dim objItemPCGeral As New ClassItemPedCompra
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_PedidoCompra_Define_Colecao

    'Verifica se existem Pedidos de Compra nas duas colecoes criadas
    If colPedCompraExclu.Count > 0 And colPedCompraGeral.Count > 0 Then
    
        bProdutoIgual = False
        For iIndice = colPedCompraExclu.Count To 1 Step -1
        
            Set objPCExclu = colPedCompraExclu.Item(iIndice)
            For Each objPCGeral In colPedCompraGeral
            
                'Verifica se os Pedidos tem o mesmo TipoDestino
                If objPCExclu.lFornecedor = objPCGeral.lFornecedor And objPCExclu.iFilial = objPCGeral.iFilial Then
                
                    For iIndice2 = objPCExclu.colItens.Count To 1 Step -1
                        
                        Set objItemPCExclu = objPCExclu.colItens.Item(iIndice2)
                        
                        For Each objItemPCGeral In objPCGeral.colItens
                        
                            'Verifica se o produto do Item Exclusivo está presente na colecao de Itens nao exclusivos
                            If objItemPCExclu.sProduto = objItemPCGeral.sProduto Then
                                bProdutoIgual = True
                                Exit For
                            End If
                        Next
                    Next
                    'Se nao encontrou produto igual nas colecoes de Itens pesquisadas
                    If bProdutoIgual = False Then
                        
                        For iIndice2 = objPCExclu.colItens.Count To 1 Step -1
                            'Adiciona o item exclusivo na colecao de itens nao exclusivos
                            objPCGeral.colItens.Add objPCExclu.colItens.Item(iIndice2)
                            'Remove o Item
                            objPCExclu.colItens.Remove (iIndice2)
                        Next
                        
                        If objPCExclu.lPedCotacao <> objPCGeral.lPedCotacao Then objPCGeral.lPedCotacao = 0
                        
                        'Remove o Pedido
                        colPedCompraExclu.Remove (iIndice)
                        
                    End If
                End If
            Next
        Next
    End If
    
    'Coloca todos os pedidos em uma única coleção
    For Each objPedidoCompra In colPedCompraExclu
        colPedidoCompras.Add objPedidoCompra
    Next
    For Each objPedidoCompra In colPedCompraGeral
        colPedidoCompras.Add objPedidoCompra
    Next
    
    PedidoCompra_Define_Colecao = SUCESSO
    
    Exit Function
    
Erro_PedidoCompra_Define_Colecao:

    PedidoCompra_Define_Colecao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161005)
            
    End Select
    
    Exit Function
    
End Function

Function Move_TipoDestino_Memoria(iTipoDestino, lDestino, iFilialDestino)

Dim objFornecedor As New ClassFornecedor
Dim lErro As Long

On Error GoTo Erro_Move_TipoDestino_Memoria

    iTipoDestino = TIPO_DESTINO_EMPRESA
    iFilialDestino = giFilialEmpresa

    Move_TipoDestino_Memoria = SUCESSO

    Exit Function
    
Erro_Move_TipoDestino_Memoria:

    Move_TipoDestino_Memoria = gErr

    Select Case gErr

        Case 63746, 63775, 70491, 63747
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161006)

    End Select

    Exit Function

End Function

'Function Atualiza_Cotacoes() As Long
'
'Dim lErro As Long
'Dim iItem As Integer
'Dim iPosicao As Integer
'Dim lDestino As Long
'Dim objProduto As ClassProduto
'Dim colProdutos As New Collection
'Dim iTipoDestino As Integer
'Dim iFilialDestino As Integer
'Dim objItemConcorrencia As ClassItemConcorrencia
'Dim iIndice As Integer
'
'On Error GoTo Erro_Atualiza_Cotacoes
'
'    lErro = Move_TipoDestino_Memoria(iTipoDestino, lDestino, iFilialDestino)
'    If lErro <> SUCESSO And iTipoDestino <> -1 Then gError 62809
'
'    For iItem = 1 To gcolItemConcorrencia.Count
'
'        Set objItemConcorrencia = gcolItemConcorrencia(iItem)
'
'        Set objItemConcorrencia.colCotacaoItemConc = New Collection
'
'        If iTipoDestino <> TIPO_DESTINO_AUSENTE Then
'            Call Localiza_Produto(objItemConcorrencia.sProduto, colProdutos, iPosicao)
'
'            If iPosicao = 0 Then
'
'                Set objProduto = New ClassProduto
'
'                objProduto.sCodigo = objItemConcorrencia.sProduto
'
'                lErro = CF("Produto_Le", objProduto)
'                If lErro <> SUCESSO And lErro <> 23080 Then gError 62810
'                If lErro <> SUCESSO Then gError 62811
'
'            Else
'                Set objProduto = colProdutos(iPosicao)
'            End If
'
'            lErro = Traz_Cotacoes_Tela(objProduto, objItemConcorrencia.dQuantidade, iItem)
'            If lErro <> SUCESSO Then gError 62812
'        End If
'    Next
'
'    For iIndice = 1 To gcolItemConcorrencia.Count
'        Call Recarrega_Cotacoes(iIndice)
'    Next
'
'    Call Indica_Melhores
'    Call GridCotacoes_Preenche
'
'    Atualiza_Cotacoes = SUCESSO
'
'    Exit Function
'
'Erro_Atualiza_Cotacoes:
'
'    Atualiza_Cotacoes = Err
'
'    Select Case gErr
'
'        Case 62809, 62810, 62812
'
'        Case 62811
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161007)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub Localiza_Produto(sProduto As String, colProdutos As Collection, iPosicao As Integer)

Dim objProduto As ClassProduto
Dim iIndice As Integer

    iPosicao = 0
    iIndice = 0
    
    For Each objProduto In colProdutos
        iIndice = iIndice + 1
        If objProduto.sCodigo = sProduto Then
            iPosicao = iIndice
            Exit Sub
        End If
    Next
            
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Gravar_Registro
    
    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(Orcamento.Text)) = 0 Then gError 94427
    
    objOrcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
        
    GL_objMDIForm.MousePointer = vbHourglass
       
    'Recolhe os dados da tela e armazena em objConcorrencia
    lErro = Move_Concorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63761

    'Insere ou Altera uma concorrencia no BD
    lErro = CF("Concorrencia_Grava_OrcVenda", objConcorrencia, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 63672

    Call Rotina_Aviso(vbOKOnly, "AVISO_CONCORRENCIA_GRAVADA", objConcorrencia.lCodigo)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
        
    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
    
    Select Case gErr

        Case 94427
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)

        Case 63756

        Case 63761, 63672
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161008)

    End Select

    Exit Function

End Function

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection
Dim iPosMoedaReal As Integer
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.sNome
        Moeda.ItemData(iIndice) = objMoeda.iCodigo
        
        iIndice = iIndice + 1
    
    Next
    
    Moeda.ListIndex = -1

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161011)
    
    End Select

End Function

Private Sub Indica_Melhores()
'Indica as melhores opcoes

Dim dMenorPreco As Double, bJaTemEscolhido As Boolean, dPrecoEmReal As Double
Dim objItemCotItemConc As ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia
Dim objItemCotItemConcAux As ClassCotacaoItemConc

On Error GoTo Erro_Indica_Melhores

    Call Grid_Refresh_Checkbox_Limpa(objGridCotacoes)
    
    For Each objItemConcorrencia In gcolItemConcorrencia
        
        'verifica se já tem algum item escolhido
        bJaTemEscolhido = False
        'Para cada produto da colecao ...
        For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
            If objItemCotItemConc.iEscolhido = MARCADO Then
                bJaTemEscolhido = True
                Exit For
            End If
        Next
        
        If bJaTemEscolhido = False Then
        
            dMenorPreco = 0
            
            Set objItemCotItemConcAux = New ClassCotacaoItemConc
            
            'Para cada produto da colecao ...
             For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
                
                'Se for para aparecer no grid ...
                If objItemCotItemConc.iSelecionada = MARCADO Then
                
                    If objItemCotItemConc.iMoeda <> MOEDA_REAL Then
                        dPrecoEmReal = objItemCotItemConc.dPrecoAjustado * objItemCotItemConc.dTaxa
                    Else
                        dPrecoEmReal = objItemCotItemConc.dPrecoAjustado
                    End If
                    
                    'Desmarca.
                    objItemCotItemConc.iEscolhido = DESMARCADO
                    
                    'Caso ainda nao tenhamos um menor preco => Menor = $$ do Primeiro item
                    If dMenorPreco = 0 Then
                        
                        dMenorPreco = dPrecoEmReal
                        
                        Set objItemCotItemConcAux = objItemCotItemConc
                        
                    End If
                    
                    'Se o preco for menor do que o menor preco ja encontrado
                    If dPrecoEmReal < dMenorPreco Then
                        
                        'Guarda o menor preco
                        dMenorPreco = dPrecoEmReal
                        
                        'Coloca o preco anterior como desmarcado
                        objItemCotItemConcAux.iEscolhido = DESMARCADO
                        
                        'Aponta para o novo candidato
                        Set objItemCotItemConcAux = objItemCotItemConc
                        
                    End If
                
                End If
                
            Next
            
            'Seleciona o Menor
            objItemCotItemConcAux.iEscolhido = MARCADO
        
        End If
        
    Next
    
    Call Grid_Refresh_Checkbox(objGridCotacoes)

    Exit Sub

Erro_Indica_Melhores:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161012)
    
    End Select

End Sub

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitarioCot.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoUnitarioReal.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
    End If

End Sub

Private Sub objEventoOrcamento_evSelecao(obj1 As Object)

Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_objEventoOrcamento_evSelecao

    Set objOrcamentoVenda = obj1

    Orcamento.Text = objOrcamentoVenda.lCodigo

    Me.Show
    
    Call BotaoOrcamento_Click
    
    Exit Sub
    
Erro_objEventoOrcamento_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub
    
End Sub

Public Sub Orcamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Orcamento_GotFocus()

    Call MaskEdBox_TrataGotFocus(Orcamento, iAlterado)

End Sub

Public Sub OrcamentoLabel_Click()

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As New Collection

    Call Chama_Tela("OrcamentoPedidoLista", colSelecao, objOrcamentoVenda, objEventoOrcamento)

End Sub

Public Sub BotaoVerOrcamento_Click()

Dim lErro As Long
Dim objOcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_BotaoVerPedido_Click

    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(Orcamento.Text)) = 0 Then gError 94427
    
    objOcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOcamentoVenda.iFilialEmpresa = giFilialEmpresa
        
    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("OrcamentoVenda_Le", objOcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 94429
    
    If lErro = 101232 Then gError 94430 'Não encontrou, erro.
    
    'Chama  a tela de Pedido de Venda passando o pedido de venda da tela
    Call Chama_Tela("OrcamentoVenda", objOcamentoVenda)
    
    Exit Sub

Erro_BotaoVerPedido_Click:

    Select Case gErr
    
        Case 94427
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)
            
        Case 94429
                
        Case 94430
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOcamentoVenda.lCodigo, objOcamentoVenda.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158083)

    End Select

    Exit Sub

End Sub

Public Sub BotaoOrcamento_Click()
             
Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_BotaoOrcamento_Click
    
    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(Orcamento.Text)) = 0 Then gError 94427
    
    objOrcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
        
    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("OrcamentoVenda_Le_Completo", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 94429
    
    If lErro = 101232 Then gError 94430 'Não encontrou, erro.
    
    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Orcamento_Tela(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94432

    Exit Sub

Erro_BotaoOrcamento_Click:

    Select Case gErr

        Case 94427
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)

        Case 94428
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)

        Case 94429, 94432

        Case 94425
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_PEDVENDA1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
        
        Case 94426
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_NFISCAL1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
        
        Case 94430
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case 94431
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE1", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158084)

    End Select

    Exit Sub

End Sub

Private Function Traz_Dados_Orcamento_Tela(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long, objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_Traz_Dados_Orcamento_Tela

    Set gcolItemConcorrencia = New Collection
    
    Set objConcorrencia.colItens = gcolItemConcorrencia
    
    lErro = CF("ItensConc_Le_GeracaoPCOV", objOrcamentoVenda, objConcorrencia)
    If lErro <> SUCESSO Then gError 184381
    
    Call Indica_Melhores
    Call GridCotacoes_Preenche

    If objConcorrencia.lCodigo <> 0 Then
    
        Concorrencia.Caption = CStr(objConcorrencia.lCodigo)
        Descricao.Text = objConcorrencia.sDescricao
        
    End If
    
    Traz_Dados_Orcamento_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Dados_Orcamento_Tela:

    Traz_Dados_Orcamento_Tela = gErr

    Select Case gErr

        Case 184381
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184379)

    End Select
    
    Exit Function

End Function

