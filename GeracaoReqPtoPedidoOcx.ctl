VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl GeracaoReqPtoPedidoOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8415
      Index           =   3
      Left            =   135
      TabIndex        =   43
      Top             =   585
      Visible         =   0   'False
      Width           =   16560
      Begin VB.CommandButton BotaoGeraReq 
         Caption         =   "Gera Requisição"
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
         Left            =   225
         TabIndex        =   60
         Top             =   3495
         Width           =   2265
      End
      Begin VB.Frame Frame6 
         Caption         =   "Local de Entrega"
         Height          =   1365
         Left            =   195
         TabIndex        =   48
         Top             =   1800
         Width           =   8295
         Begin VB.Frame Frame2 
            Caption         =   "Tipo"
            Height          =   585
            Index           =   1
            Left            =   225
            TabIndex        =   57
            Top             =   390
            Width           =   3870
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Empresa/Filial"
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
               Left            =   420
               TabIndex        =   59
               Top             =   225
               Value           =   -1  'True
               Width           =   1635
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Fornecedor"
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
               Index           =   1
               Left            =   2295
               TabIndex        =   58
               Top             =   225
               Width           =   1335
            End
         End
         Begin VB.Frame FrameTipoDestino 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   675
            Index           =   0
            Left            =   4515
            TabIndex        =   54
            Top             =   360
            Width           =   3645
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               Left            =   1230
               TabIndex        =   55
               Top             =   210
               Width           =   2160
            End
            Begin VB.Label Label6 
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
               Left            =   690
               TabIndex        =   56
               Top             =   255
               Width           =   465
            End
         End
         Begin VB.Frame FrameTipoDestino 
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   4515
            TabIndex        =   49
            Top             =   345
            Visible         =   0   'False
            Width           =   3645
            Begin VB.ComboBox FilialFornec 
               Height          =   315
               Left            =   1230
               TabIndex        =   50
               Top             =   345
               Width           =   2160
            End
            Begin MSMask.MaskEdBox Fornec 
               Height          =   300
               Left            =   1260
               TabIndex        =   51
               Top             =   15
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label Label21 
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
               Left            =   690
               TabIndex        =   53
               Top             =   405
               Width           =   465
            End
            Begin VB.Label FornecLabel 
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
               Height          =   195
               Left            =   150
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   52
               Top             =   60
               Width           =   1035
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Principais"
         Height          =   1275
         Left            =   195
         TabIndex        =   44
         Top             =   360
         Width           =   8295
         Begin VB.ComboBox FilialCompra 
            Height          =   315
            Left            =   1680
            TabIndex        =   46
            Top             =   540
            Width           =   2610
         End
         Begin VB.CheckBox EnviaCompras 
            Caption         =   "Envia para Compras"
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
            Left            =   5340
            TabIndex        =   45
            Top             =   585
            Width           =   2040
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Filial Compra:"
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
            Left            =   420
            TabIndex        =   47
            Top             =   585
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8445
      Index           =   1
      Left            =   210
      TabIndex        =   1
      Top             =   525
      Width           =   16560
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Produtos"
         Height          =   6345
         Index           =   0
         Left            =   420
         TabIndex        =   41
         Top             =   615
         Width           =   8985
         Begin VB.CommandButton BotaoDesmarcarTodosTipos 
            Caption         =   "Desmarcar Todos"
            Height          =   570
            Left            =   2145
            Picture         =   "GeracaoReqPtoPedidoOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   5640
            Width           =   1425
         End
         Begin VB.CommandButton BotaoMarcarTodosTipos 
            Caption         =   "Marcar Todos"
            Height          =   570
            Left            =   675
            Picture         =   "GeracaoReqPtoPedidoOcx.ctx":11E2
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5640
            Width           =   1425
         End
         Begin VB.Frame Frame9 
            Caption         =   "Produto"
            Height          =   1515
            Left            =   5565
            TabIndex        =   6
            Top             =   600
            Width           =   3255
            Begin MSMask.MaskEdBox CodProdutoDe 
               Height          =   315
               Left            =   780
               TabIndex        =   8
               Top             =   420
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodProdutoAte 
               Height          =   315
               Left            =   780
               TabIndex        =   10
               Top             =   990
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label ProdutoDeLabel 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               Left            =   375
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   7
               Top             =   465
               Width           =   315
            End
            Begin VB.Label ProdutoAteLabel 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   330
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   9
               Top             =   1020
               Width           =   360
            End
         End
         Begin VB.ListBox TipoProduto 
            Height          =   4785
            Left            =   840
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   630
            Width           =   4515
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Produto"
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
            Left            =   810
            TabIndex        =   2
            Top             =   390
            Width           =   1470
         End
      End
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   15630
      Picture         =   "GeracaoReqPtoPedidoOcx.ctx":21FC
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Fechar"
      Top             =   60
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   8490
      Index           =   2
      Left            =   210
      TabIndex        =   11
      Top             =   525
      Visible         =   0   'False
      Width           =   16455
      Begin VB.CommandButton BotaoAlmoxarifados 
         Caption         =   "Almoxarifados"
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
         Left            =   4800
         TabIndex        =   38
         Top             =   8010
         Width           =   1485
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
         Height          =   330
         Left            =   3015
         TabIndex        =   37
         Top             =   8010
         Width           =   1710
      End
      Begin VB.CommandButton BotaoDesmarcarTodosProdutos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   1485
         Picture         =   "GeracaoReqPtoPedidoOcx.ctx":237A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   7890
         Width           =   1410
      End
      Begin VB.CommandButton BotaoMarcarTodosProdutos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   60
         Picture         =   "GeracaoReqPtoPedidoOcx.ctx":355C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7890
         Width           =   1350
      End
      Begin VB.CommandButton BotaoProdutoFiliaisForn 
         Caption         =   "Produto x Fornecedores"
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
         Left            =   6345
         TabIndex        =   39
         Top             =   8010
         Width           =   2310
      End
      Begin VB.Frame FrameProdutos 
         Caption         =   "Produtos em Ponto de Pedido"
         Height          =   7770
         Left            =   90
         TabIndex        =   12
         Top             =   60
         Width           =   16320
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   225
            Left            =   3405
            TabIndex        =   42
            Top             =   3105
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            HideSelection   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6480
            MaxLength       =   255
            TabIndex        =   34
            Top             =   2895
            Width           =   2160
         End
         Begin VB.ComboBox FilialForn 
            Height          =   315
            Left            =   4590
            TabIndex        =   32
            Top             =   2865
            Width           =   1980
         End
         Begin VB.ComboBox TipoTributacao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "GeracaoReqPtoPedidoOcx.ctx":4576
            Left            =   1500
            List            =   "GeracaoReqPtoPedidoOcx.ctx":4578
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2700
            Width           =   2385
         End
         Begin MSMask.MaskEdBox QuantEmReq 
            Height          =   225
            Left            =   -30
            TabIndex        =   22
            Top             =   795
            Width           =   1545
            _ExtentX        =   2725
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
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1305
            MaxLength       =   50
            TabIndex        =   16
            Top             =   465
            Width           =   4000
         End
         Begin VB.ComboBox Exclusivo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "GeracaoReqPtoPedidoOcx.ctx":457A
            Left            =   6525
            List            =   "GeracaoReqPtoPedidoOcx.ctx":4584
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   3030
            Width           =   1215
         End
         Begin VB.CheckBox Selecionado 
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
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   1050
         End
         Begin MSMask.MaskEdBox IntervaloRessup 
            Height          =   225
            Left            =   5775
            TabIndex        =   27
            Top             =   885
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ConsumoMedio 
            Height          =   225
            Left            =   4485
            TabIndex        =   26
            Top             =   855
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SaldoAtual 
            Height          =   225
            Left            =   1305
            TabIndex        =   23
            Top             =   810
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
         Begin MSMask.MaskEdBox LoteEconomico 
            Height          =   225
            Left            =   6825
            TabIndex        =   28
            Top             =   795
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   450
            TabIndex        =   15
            Top             =   390
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PontoPedido 
            Height          =   225
            Left            =   2175
            TabIndex        =   24
            Top             =   840
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   -330
            TabIndex        =   29
            Top             =   3075
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   3270
            TabIndex        =   25
            Top             =   825
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox UnidadeMed 
            Height          =   225
            Left            =   2715
            TabIndex        =   17
            Top             =   390
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantEstoque 
            Height          =   225
            Left            =   3555
            TabIndex        =   18
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantEmPedido 
            Height          =   225
            Left            =   7425
            TabIndex        =   21
            Top             =   390
            Width           =   1260
            _ExtentX        =   2223
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
         Begin MSMask.MaskEdBox QuantReservada 
            Height          =   225
            Left            =   4650
            TabIndex        =   19
            Top             =   405
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSMask.MaskEdBox QuantEmpenhada 
            Height          =   225
            Left            =   6045
            TabIndex        =   20
            Top             =   405
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   225
            Left            =   1290
            TabIndex        =   30
            Top             =   3045
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos 
            Height          =   7320
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   15960
            _ExtentX        =   28152
            _ExtentY        =   12912
            _Version        =   393216
            Rows            =   12
            Cols            =   24
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8910
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   16770
      _ExtentX        =   29580
      _ExtentY        =   15716
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Geração"
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
Attribute VB_Name = "GeracaoReqPtoPedidoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis Globais
Dim iSelecaoAlterada As Integer
Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim gobjGeracaoPtoPedido As ClassGeracaoPtoPedido
Dim giTipoTributacao As Integer

Dim iFrameTipoDestinoAtual As Integer

Dim iFornecedorAlterado As Integer

'GridProdutos
Dim objGridProdutos As New AdmGrid
Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_QuantEstoque_Col As Integer
Dim iGrid_QuantReservada_Col As Integer
Dim iGrid_QuantEmpenhada_Col As Integer
Dim iGrid_QuantEmPedido_Col As Integer
Dim iGrid_QuantEmReq_Col As Integer
Dim iGrid_SaldoAtual_Col As Integer
Dim iGrid_PontoPedido_Col As Integer
Dim iGrid_QuantComprar_Col As Integer
Dim iGrid_ConsumoMedio_Col As Integer
Dim iGrid_IntervaloRessup_Col As Integer
Dim iGrid_LoteEconomico_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_ContaContabil_Col As Integer
Dim iGrid_TipoTributacao_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialFornecedor_Col As Integer
Dim iGrid_Exclusivo_Col As Integer
Dim iGrid_Observacao_Col As Integer

'Constante da Tela
Const TAB_Produtos = 2

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos dos Browses
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoFiliaisFornProduto As AdmEvento
Attribute objEventoFiliaisFornProduto.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxarifados As AdmEvento
Attribute objEventoAlmoxarifados.VB_VarHelpID = -1

Function Trata_Parametros(Optional gobjGeracaoPtoPedido As ClassGeracaoPtoPedido) As Long
    
    Trata_Parametros = SUCESSO
    
End Function

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim bCancel As Boolean
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    Set objEventoFornecedor = New AdmEvento
    Set objEventoFiliaisFornProduto = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoAlmoxarifados = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    
    'Carrega Tipos de Produto
    lErro = Carrega_TipoProduto()
    If lErro <> SUCESSO Then gError 67309
    
    'Carrega Tipos de Tributação
    lErro = Carrega_TipoTributacao()
    If lErro <> SUCESSO Then gError 67313
        
    'Inicializa Mascara o Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 67314
    
    'Inicializa Mascara o Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodProdutoDe)
    If lErro <> SUCESSO Then gError 67254
    
    'Inicializa Mascara o Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodProdutoAte)
    If lErro <> SUCESSO Then gError 67255
    
    'Inicializa Mascara da Conta Contábil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 67315
    
    Quantidade.Format = FORMATO_ESTOQUE
    ConsumoMedio.Format = FORMATO_ESTOQUE
    PontoPedido.Format = FORMATO_ESTOQUE
    QuantEstoque.Format = FORMATO_ESTOQUE
    QuantReservada.Format = FORMATO_ESTOQUE
    QuantEmpenhada.Format = FORMATO_ESTOQUE
    QuantEmPedido.Format = FORMATO_ESTOQUE
    QuantEmReq.Format = FORMATO_ESTOQUE
    SaldoAtual.Format = FORMATO_ESTOQUE
    ConsumoMedio.Format = FORMATO_ESTOQUE
    
    'Carrega as filiais Empresas na combo de Filial de compra
    lErro = Carrega_ComboFiliais(colCodigoDescricao)
    If lErro <> SUCESSO Then gError 67316
    
    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then gError 67317

    'Coloca FilialCompra Default na tela
    If objConfiguraCOM.iFilialCompra > 0 Then
        FilialCompra.Text = objConfiguraCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If
    Call FilialCompra_Validate(bCancel)
    
    'Seleciona em Filial Empresa, a Matriz como Default
    bCancel = False
    
    iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126943
    
    FilialEmpresa.Text = iFilialEmpresa
    Call FilialEmpresa_Validate(bCancel)
    
    Set objGridProdutos = New AdmGrid
    
    'Inicializa GridProdutos
    lErro = Inicializa_GridProdutos(objGridProdutos)
    If lErro <> SUCESSO Then gError 67318
         
   'Atualiza a variável global para controle de frames e seta um tipo Padrao
    iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 67254, 67255, 67309, 67313, 67314, 67315, 67316, 67317, 67318, 126493
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161494)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Libera variáveis globais
    Set objEventoFiliaisFornProduto = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoAlmoxarifados = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoFornecedor = Nothing
    
    Set objGridProdutos = Nothing
    Set gobjGeracaoPtoPedido = Nothing
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Function Carrega_TipoProduto() As Long
'Carrega a ListBox TipoProduto com tipos de produtos que possam ser comprados (Compras=1)

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoProduto

    'Le todos os Codigos e DescReduzida de tipos de produtos
    lErro = CF("TiposProduto_Le_Todos", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 67310

    For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na ListBox CondPagtos
        TipoProduto.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        TipoProduto.ItemData(TipoProduto.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    'Marca todos os Tipos de Produto
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = True
    Next

    Carrega_TipoProduto = SUCESSO

    Exit Function

Erro_Carrega_TipoProduto:

    Carrega_TipoProduto = gErr

    Select Case gErr

        Case 67310
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161495)

    End Select

    Exit Function

End Function

Function Carrega_TipoTributacao() As Long
'Carrega Tipos de Tributação

Dim lErro As Long
Dim colTributacao As New AdmColCodigoNome
Dim iIndice As Integer
Dim iTipoTrib As Integer

On Error GoTo Erro_Carrega_TipoTributacao

    'Lê os Tipos de Tributação associadas a Compras
    lErro = CF("TiposTributacaoCompras_Le", colTributacao)
    If lErro <> SUCESSO Then gError 67311
        
    'Lê o Tipo de Tributação Padrão
    lErro = CF("TipoTributacaoPadrao_Le", iTipoTrib)
    If lErro <> SUCESSO And lErro <> 66597 Then gError 67312
    
    'Carrega Tipos de Tributação
    For iIndice = 1 To colTributacao.Count
        TipoTributacao.AddItem colTributacao(iIndice).iCodigo & SEPARADOR & colTributacao(iIndice).sNome
    Next
    
    'Seleciona Tipo de Tributação default
    For iIndice = 0 To TipoTributacao.ListCount - 1
        If Codigo_Extrai(TipoTributacao.List(iIndice)) = iTipoTrib Then
            TipoTributacao.ListIndex = iIndice
            Exit For
        End If
    Next
            
    'Guarda tipo de tributação padrão
    giTipoTributacao = iTipoTrib
    
    Carrega_TipoTributacao = SUCESSO
    
    Exit Function
    
Erro_Carrega_TipoTributacao:

    Carrega_TipoTributacao = gErr
    
    Select Case gErr
        
        Case 67311, 67312
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161496)
        
    End Select
    
    Exit Function
    
End Function

Function Carrega_ComboFiliais(colCodigoDescricao As AdmColCodigoNome) As Long
'Carrega a Combo FilialCompra com as Filiais Empresa passada na colecao

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Carrega_ComboFiliais

    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 67450

    'Preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao

        If objCodigoNome.iCodigo <> 0 Then

            FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo

            FilialCompra.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialCompra.ItemData(FilialCompra.NewIndex) = objCodigoNome.iCodigo

        End If

    Next

    Carrega_ComboFiliais = SUCESSO

    Exit Function

Erro_Carrega_ComboFiliais:

    Carrega_ComboFiliais = gErr

    Select Case gErr

        Case 67450
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161497)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridProdutos(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Selecionado")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U. M.")
    objGridInt.colColuna.Add ("Em Estoque")
    objGridInt.colColuna.Add ("Reservada")
    objGridInt.colColuna.Add ("Empenhada")
    objGridInt.colColuna.Add ("Em Pedido")
    objGridInt.colColuna.Add ("Em Requisição")
    objGridInt.colColuna.Add ("Saldo Atual")
    objGridInt.colColuna.Add ("Ponto Pedido")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Consumo médio")
    objGridInt.colColuna.Add ("Int. de Ressup")
    objGridInt.colColuna.Add ("Lote Econômico")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Conta Contábil")
    objGridInt.colColuna.Add ("Tipo de Tributação")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusividade")
    objGridInt.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Selecionado.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (QuantEstoque.Name)
    objGridInt.colCampo.Add (QuantReservada.Name)
    objGridInt.colCampo.Add (QuantEmpenhada.Name)
    objGridInt.colCampo.Add (QuantEmPedido.Name)
    objGridInt.colCampo.Add (QuantEmReq.Name)
    objGridInt.colCampo.Add (SaldoAtual.Name)
    objGridInt.colCampo.Add (PontoPedido.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (ConsumoMedio.Name)
    objGridInt.colCampo.Add (IntervaloRessup.Name)
    objGridInt.colCampo.Add (LoteEconomico.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)
    objGridInt.colCampo.Add (TipoTributacao.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (FilialForn.Name)
    objGridInt.colCampo.Add (Exclusivo.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    
    'Colunas do Grid
    iGrid_Selecionado_Col = 1
    iGrid_Produto_Col = 2
    iGrid_DescProduto_Col = 3
    iGrid_UnidadeMed_Col = 4
    iGrid_QuantEstoque_Col = 5
    iGrid_QuantReservada_Col = 6
    iGrid_QuantEmpenhada_Col = 7
    iGrid_QuantEmPedido_Col = 8
    iGrid_QuantEmReq_Col = 9
    iGrid_SaldoAtual_Col = 10
    iGrid_PontoPedido_Col = 11
    iGrid_QuantComprar_Col = 12
    iGrid_ConsumoMedio_Col = 13
    iGrid_IntervaloRessup_Col = 14
    iGrid_LoteEconomico_Col = 15
    iGrid_Almoxarifado_Col = 16
    iGrid_ContaContabil_Col = 17
    iGrid_TipoTributacao_Col = 18
    iGrid_Fornecedor_Col = 19
    iGrid_FilialFornecedor_Col = 20
    iGrid_Exclusivo_Col = 21
    iGrid_Observacao_Col = 22

    'Grid do GridInterno
    objGridInt.objGrid = GridProdutos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PRODUTOS_PONTOPEDIDO + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 20

    'Largura da primeira coluna
    GridProdutos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'É proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridProdutos = SUCESSO

    Exit Function

End Function

Private Sub BotaoGeraReq_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objRequisicao As New ClassRequisicaoCompras
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_BotaoGeraReq_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se existe pelo menos uma linha selecionada no GridProdutos
    For iIndice = 1 To objGridProdutos.iLinhasExistentes
        If GridProdutos.TextMatrix(iIndice, iGrid_Selecionado_Col) = "1" Then
            Exit For
        End If
    Next
    
    'Se não, erro
    If iIndice > objGridProdutos.iLinhasExistentes Then gError 67402
    
    'Se FilialCompra não está preenchida, Erro
    If Len(Trim(FilialCompra.Text)) = 0 Then gError 67457
    
    'Verifica se o Grid foi preenchido
    If objGridProdutos.iLinhasExistentes = 0 Then gError 84575 '61971

    'Se o tipo destino for empresa
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then
    
        'Se a FilialEmpresa não estiver preenchida, erro
        If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 84576 '61972
    
        'Critica o preenchimento do Almoxarifado
        For iIndice = 1 To objGridProdutos.iLinhasExistentes
                
            'Formata o Produto
            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 84577 '67880
            
            objProduto.sCodigo = sProdutoFormatado
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 23080 Then gError 84578 '67881
            If lErro <> SUCESSO Then gError 84579 '67882
            
            'Se o produto possui Estoque e o almoxarifado não foi preenchido, erro
            If Not (objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) And Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 84580
                            
        Next
    
    'Se o tipo destino for Fornecedor
    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then
    
        'Se o Fornecedor não estiver preenchido, erro
        If Len(Trim(Fornec.Text)) = 0 Then gError 84582
        
        'Se a Filial do Fornecedor não estiver preenchida, erro
        If Len(Trim(FilialFornec.Text)) = 0 Then gError 84583
            
    End If
    
    'Move GridProdutos para a memória
    lErro = Move_GridProdutos_Memoria(objRequisicao)
    If lErro <> SUCESSO Then gError 67403
    
    'Gera Requisição
    lErro = CF("GeracaoRequisicaoPtoPedido_Grava", objRequisicao)
    If lErro <> SUCESSO Then gError 67404
        
    'Limpa a tela
    Call Limpa_Tela_GeracaoRePtoPedido
    
    'Informa o Código da Requisição gerada
    vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_REQUISICAOCOMPRA_GERADA", objRequisicao.lCodigo)
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoGeraReq_Click:

    Select Case gErr
    
        Case 67402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 67403, 67404
        
        Case 67457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", gErr)
        
        Case 84575
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDITENS_VAZIO", gErr)
        
        Case 84576
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)
        
        Case 84582
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO", gErr)
        
        Case 84583
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA", gErr)
                        
        Case 61977, 61978, 61979, 84577, 89156
        
        Case 62683
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENS_MESMO_LEQUE", gErr, iIndice, iIndice2)
        
        Case 84579, 84578 'Alterado por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
         Case 61981
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENS_MESMO_LEQUE", gErr, iIndice, iIndice2)
        
        Case 61980
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
            
        Case 67882
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161498)
    
    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
End Sub

Function Limpa_Tela_GeracaoRePtoPedido()

    'Função genérica que limpa a tela
    Call Limpa_Tela(Me)
    
    'Limpa GridProdutos
    Call Grid_Limpa(objGridProdutos)
    
    EnviaCompras.Value = vbUnchecked
    
    iSelecaoAlterada = REGISTRO_ALTERADO
    
End Function

Function Move_GridProdutos_Memoria(objRequisicao As ClassRequisicaoCompras) As Long

Dim lErro As Long
Dim objItemRC As ClassItemReqCompras
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim lCodigo As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_Move_GridProdutos_Memoria

    'Move os dados da tela para a memória
    objRequisicao.sDigitador = gsUsuario
    objRequisicao.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO
    objRequisicao.iFilialCompra = Codigo_Extrai(FilialCompra.Text)
    objRequisicao.iTipoTributacao = giTipoTributacao
    
    If EnviaCompras.Value = vbChecked Then
        objRequisicao.dtDataEnvio = gdtDataHoje
    Else
        objRequisicao.dtDataEnvio = DATA_NULA
    End If
    
    objRequisicao.dtData = gdtDataAtual
    objRequisicao.dtDataBaixa = DATA_NULA
    objRequisicao.iTipoDestino = TIPO_DESTINO_EMPRESA
    objRequisicao.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)
    objRequisicao.iFilialEmpresa = giFilialEmpresa
    objRequisicao.dtDataLimite = DATA_NULA
    
    'Gera código automático para a Requisição de Compras
    lErro = CF("RequisicaoCompras_Codigo_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 67449
    
    objRequisicao.lCodigo = lCodigo
    
    Set objRequisicao.colItens = New Collection
    
    'Move os Itens do Grid para a memória
    For iIndice = 1 To objGridProdutos.iLinhasExistentes
        
        'Se a linha do Grid foi selecionada
        If GridProdutos.TextMatrix(iIndice, iGrid_Selecionado_Col) = "1" Then
            
            'Se o Consumo Médido não foi preenchido, Erro
            If StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_ConsumoMedio_Col)) = 0 Then gError 67405
    
            'Se o Intervalo de Ressuprimento não foi preenchido, Erro
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_IntervaloRessup_Col))) = 0 Then gError 67406
            
            'Se a Quantidade a comprar não foi preenchida ou for igual a 0, erro
            If StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col)) = 0 Then gError 67407
            
            Set objItemRC = New ClassItemReqCompras
            objItemRC.dQuantidade = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col))
            
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) > 0 Then
            
                objAlmoxarifado.sNomeReduzido = GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
        
                'Lê dados do almoxarifado a partir do Nome Reduzido
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 67408
        
                'Se não econtrou o almoxarifado, erro
                If lErro = 25060 Then gError 67409
        
                If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 86170
        
                objItemRC.iAlmoxarifado = objAlmoxarifado.iCodigo
            
            End If
            
            objItemRC.iTipoTributacao = Codigo_Extrai(GridProdutos.TextMatrix(iIndice, iGrid_TipoTributacao_Col))
            
            'Formata o Produto
            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 67413
                    
            objItemRC.sProduto = sProdutoFormatado
            objItemRC.iStatus = STATUS_ABERTO
            objItemRC.sUM = GridProdutos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
            objItemRC.sObservacao = GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col)
            
            'Move o Código do Fornecedor
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Fornecedor_Col))) > 0 Then
                
                objFornecedor.sNomeReduzido = GridProdutos.TextMatrix(iIndice, iGrid_Fornecedor_Col)
                
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 67410
                If lErro = 6681 Then gError 67411
                
                objItemRC.lFornecedor = objFornecedor.lCodigo
                
                'Move o código da Filial do Fornecedor
                If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_FilialFornecedor_Col))) > 0 Then
                    objItemRC.iFilial = Codigo_Extrai(GridProdutos.TextMatrix(iIndice, iGrid_FilialFornecedor_Col))
                End If
                
                objFornecedorProdutoFF.lFornecedor = objItemRC.lFornecedor
                objFornecedorProdutoFF.iFilialForn = objItemRC.iFilial
                objFornecedorProdutoFF.sProduto = objItemRC.sProduto
                objFornecedorProdutoFF.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
                
                lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
                If lErro <> SUCESSO And lErro <> 54217 Then gError 86097
                If lErro <> SUCESSO Then gError 86098
                
                
            End If
        
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_ContaContabil_Col))) > 0 Then
            
                'Formata ContaContábil para BD
                lErro = CF("Conta_Formata", GridProdutos.TextMatrix(iIndice, iGrid_ContaContabil_Col), sContaFormatada, iContaPreenchida)
                If lErro <> SUCESSO Then gError 67412
            Else
                sContaFormatada = ""
            End If
            objItemRC.sContaContabil = sContaFormatada
            objItemRC.sDescProduto = GridProdutos.TextMatrix(iIndice, iGrid_DescProduto_Col)
            
            If GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col) = "Preferencial" Then
                objItemRC.iExclusivo = ITEM_FILIALFORNECEDOR_PREFERENCIAL
            ElseIf GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col) = "Exclusivo" Then
                objItemRC.iExclusivo = ITEM_FILIALFORNECEDOR_EXCLUSIVO
            End If
            
            'Adiciona o ItemRC na coleção da Requisição de Compras
            objRequisicao.colItens.Add objItemRC
        
        End If
        
    Next
    
    Move_GridProdutos_Memoria = SUCESSO

    Exit Function
    
Erro_Move_GridProdutos_Memoria:

    Move_GridProdutos_Memoria = gErr
    
    Select Case gErr
    
        Case 67405
            Call Rotina_Erro(vbOKOnly, "ERRO_CONSUMOMEDIO_NAO_PREENCHIDO", gErr, iIndice)
                    
        Case 67406
            Call Rotina_Erro(vbOKOnly, "ERRO_INTERVALORESSUP_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 67407
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 67408, 67410, 67412, 67449, 86097
        
        Case 67409
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)
        
        Case 67411
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
                    
        Case 86098
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
                    
        Case 86170
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161499)
    
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    'Se nenhuma linha foi selecionada, erro
    If GridProdutos.Row = 0 Then gError 67451

    'Se o Produto da linha selecionada não foi preenchido, erro
    If GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col) = "" Then gError 67452

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 67453

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 67451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 67452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 67453

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161500)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAlmoxarifados_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iFilialEmpresa As Integer

On Error GoTo Erro_BotaoALmoxarifados_Click

    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = False Then Exit Sub
    
    'Se nenhuma linha foi selecionada, erro
    If GridProdutos.Row = 0 Then gError 67454

    sCodProduto = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col)

    sProdutoFormatado = String(STRING_PRODUTO, 0)

    'Formata o Produto
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 67456
    
    iFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text)
    If Len(Trim(iFilialEmpresa)) = 0 Then gError 84611

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        'adiciona produto e FiliaEmpresa de Destino em colSelecao
        colSelecao.Add sProdutoFormatado
        colSelecao.Add iFilialEmpresa
        
        Call Chama_Tela("AlmoxarifadoFilialLista", colSelecao, objEstoqueProduto, objEventoAlmoxarifados)

    Else
        gError 67455
    End If

    Exit Sub

Erro_BotaoALmoxarifados_Click:

    Select Case gErr

        Case 84611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_ENTREGA_NAO_PREENCHIDA", Err)
        
        Case 67454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 67455
             lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 67456

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161501)

    End Select

    Exit Sub

End Sub



Private Sub FilialEmpresa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Click()

Dim lErro As Long
Dim iCodFilial As Integer

On Error GoTo Erro_FilialEmpresa_Click
    
    'Se nenhuma FilialEmpresa foi selecionada, sai da rotina
    If FilialEmpresa.ListIndex = -1 Then Exit Sub
    
    'Guarda o código da Filial
    iCodFilial = Codigo_Extrai(FilialEmpresa.Text)
    
    lErro = AlmoxarifadoPadrao_Preenche(iCodFilial)
    If lErro <> SUCESSO Then gError 84661 '(65562)
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub
    
Erro_FilialEmpresa_Click:
    
    Select Case gErr
        
        Case 84661 'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161502)
        
    End Select
    
    Exit Sub




End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim iCodFilial As Integer
Dim iIndice As Integer

On Error GoTo Erro_FilialEmpresa_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialEmpresa.Text)) = 0 Then
        For iIndice = 1 To objGridProdutos.iLinhasExistentes
            GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = ""
        Next
        Exit Sub
    End If

    'Verifica se é uma FilialEmpresa selecionada
    If FilialEmpresa.Text = FilialEmpresa.List(FilialEmpresa.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84557 '61922

    'Se não encontrou o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 84558

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 84559

        'coloca na tela
        FilialEmpresa.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

        lErro = AlmoxarifadoPadrao_Preenche(iCodigo)
        If lErro <> SUCESSO Then gError 86163

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 84560 '61925
        
    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True
    
    Select Case gErr

        Case 84557, 84558, 86163

        Case 84559
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialEmpresa.Text)
            
        Case 84560
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialEmpresa.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161503)

    End Select

    Exit Sub


End Sub

Private Sub FilialFornec_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornec_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornec_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        'Verifica preenchimento de Fornecedor
        If Len(Trim(Fornec.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornec, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 84562 '61915

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 84563 '61916

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornec, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornec, iCodFilial)

        ElseIf Len(Trim(Fornec.Text)) = 0 Then

            'Se Fornecedor não foi preenchido limpa a combo de Filiais
            FilialFornec.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 84562, 84563

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161504)

    End Select

    Exit Sub


End Sub

Private Sub FornecLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornec.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Public Sub objEventoAlmoxarifados_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto
Dim bCancel As Boolean

    Set objEstoqueProduto = obj1

    'Preenche campo Almoxarifado
    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta <> "" Then

        'Se a linha selecionada não está preenchida
        If GridProdutos.Row > objGridProdutos.iLinhasExistentes Then
            Me.Show
            Exit Sub
        End If
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 67458
        
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
        
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 67458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161505)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutoFiliaisForn_Click()

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoProdutoFiliaisForn_Click

    'Se nenhuma linha foi selecionada, Erro
    If GridProdutos.Row = 0 Then gError 67389

    sCodProduto = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col)

    'Formata o Produto
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 67390

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        colSelecao.Add Codigo_Extrai(FilialCompra.Text)
        
        Call Chama_Tela("FiliaisFornProdutoLista", colSelecao, objFornecedorProdutoFF, objEventoFiliaisFornProduto)
    Else
        gError 67391
    End If

    Exit Sub

Erro_BotaoProdutoFiliaisForn_Click:

    Select Case gErr

        Case 67389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 67390

        Case 67391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161506)

    End Select

    Exit Sub

End Sub

Private Sub CodProdutoDe_Change()
    
    iSelecaoAlterada = REGISTRO_ALTERADO

End Sub

Private Sub CodProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_CodProdutoDe_Validate
    
    'Se o código do produto não foi preenchido, sai da rotina
    If Len(Trim(CodProdutoDe.ClipText)) = 0 Then Exit Sub
    
    'Formata o Produto
    lErro = CF("Produto_Formata", CodProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 67259
    
    objProduto.sCodigo = sProdutoFormatado
        
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 67392
    
    'Se o Produto não está cadastrado, pergunta se deseja criar
    If lErro = 28030 Then gError 67393
    
    'Se o código do produto inicial é maior que o final, erro
    If CodProdutoDe.Text > CodProdutoAte.Text And Len(Trim(CodProdutoAte.ClipText)) > 0 Then gError 67394
    
    Exit Sub
    
Erro_CodProdutoDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 67392
        
        Case 67393
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Produto", objProduto)
            End If
        
        Case 67394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTODE_MAIOR_PRODUTOATE", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161507)
    
    End Select
    
    Exit Sub

End Sub

Private Sub CodProdutoAte_Change()

    iSelecaoAlterada = REGISTRO_ALTERADO
    
End Sub

Private Sub CodProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String

On Error GoTo Erro_CodProdutoAte_Validate
    
    'Se o código do produto não foi preenchido, sai da rotina
    If Len(Trim(CodProdutoAte.ClipText)) = 0 Then Exit Sub
        
    'Formata o Produto
    lErro = CF("Produto_Formata", CodProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 67259
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 67319
    
    'Se o Produto não está cadastrado, pergunta se deseja criar
    If lErro = 28030 Then gError 67320
    
    'Se o código do produto inicial é maior que o final, erro
    If CodProdutoDe.Text > CodProdutoAte.Text And Len(Trim(CodProdutoDe.ClipText)) > 0 Then gError 67395
    
    Exit Sub
    
Erro_CodProdutoAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 67259, 67319
        
        Case 67320
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Produto", objProduto)
            End If
        
        Case 67395
           lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTODE_MAIOR_PRODUTOATE", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161508)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub FilialCompra_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCompra_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialCompra.Text)) = 0 Then Exit Sub

    'Verifica se é uma FilialEmpresa selecionada
    If FilialCompra.Text = FilialCompra.List(FilialCompra.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCompra, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 67386

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 67385

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 67387

        'coloca na tela
        FilialCompra.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 67388

    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case gErr

        Case 67385, 67386
            
        Case 67387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialCompra.Text)

        Case 67388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialCompra.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161509)

    End Select

    Exit Sub

End Sub



Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TipoDestino_Click(Index As Integer)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TipoDestino_Click

    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame correspondente a Index visivel
    FrameTipoDestino(Index).Visible = True
    
    'Torna Frame atual invisivel
    FrameTipoDestino(iFrameTipoDestinoAtual).Visible = False
    
    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

    'Se o Destino da não é a própria empresa
    If Index <> TIPO_DESTINO_EMPRESA Then

        'Limpa os almoxarifados e os Ccls do GridProdutos
        For iIndice = 1 To objGridProdutos.iLinhasExistentes
            GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = ""
        Next
    Else

        'Seleciona a Filial Empresa na combo
        Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)
        Call FilialEmpresa_Click

    End If

    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_TipoDestino_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161510)

    End Select

    Exit Sub

End Sub

Private Sub TipoProduto_Click()

    iSelecaoAlterada = REGISTRO_ALTERADO

End Sub

Private Sub BotaoMarcarTodosProdutos_Click()

Dim iIndice As Integer

    'Marca todas as linhas do Grid de Produtos
    For iIndice = 1 To objGridProdutos.iLinhasExistentes
        GridProdutos.TextMatrix(iIndice, iGrid_Selecionado_Col) = "1"
    Next
    
    Call Grid_Refresh_Checkbox(objGridProdutos)
    
End Sub

Private Sub BotaoDesmarcarTodosProdutos_Click()

Dim iIndice As Integer

    'Desmarca todas as linhas do Grid de Produtos
    For iIndice = 1 To objGridProdutos.iLinhasExistentes
        GridProdutos.TextMatrix(iIndice, iGrid_Selecionado_Col) = "0"
    Next
    
    Call Grid_Refresh_Checkbox(objGridProdutos)

End Sub

Private Sub BotaoMarcarTodosTipos_Click()

Dim iIndice As Integer

    'Marca todos os Tipos de Produtos
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = True
    Next
    
End Sub

Private Sub BotaoDesmarcarTodosTipos_Click()

Dim iIndice As Integer

    'Desmarca todos os Tipos de Produtos
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = False
    Next

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub ProdutoDeLabel_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoDeLabel_Click

    'Verifica se Produto está preenchido
    If Len(Trim(CodProdutoDe.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", CodProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 67396

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objProduto.sCodigo = sProdutoFormatado
        Else
            objProduto.sCodigo = ""
        End If

    End If

    'Chama a tela de browse
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub
    
Erro_ProdutoDeLabel_Click:

    Select Case gErr

        Case 67396
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161511)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoAteLabel_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoAteLabel_Click

    'Verifica se Produto está preenchido
    If Len(Trim(CodProdutoAte.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", CodProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 67397

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objProduto.sCodigo = sProdutoFormatado
        Else
            objProduto.sCodigo = ""
        End If

    End If

    'Chama a tela de browse
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub
    
Erro_ProdutoAteLabel_Click:

    Select Case gErr

        Case 67397
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161512)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    CodProdutoDe.PromptInclude = False
    CodProdutoDe.Text = objProduto.sCodigo
    CodProdutoDe.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161513)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    CodProdutoAte.PromptInclude = False
    CodProdutoAte.Text = objProduto.sCodigo
    CodProdutoAte.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161514)

    End Select

    Exit Sub

End Sub

Public Sub objEventoFiliaisFornProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_objEventoFiliaisFornProd_evSelecao

    Set objFornecedorProdutoFF = obj1

    'Lê o Nome Reduzido do Fornecedor
    objFornecedor.lCodigo = objFornecedorProdutoFF.lFornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO Then gError 67398

    'Preenche campo Fornecedor
    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
    Fornecedor.Text = objFornecedor.sNomeReduzido
    
    'Lê coleção de códigos e nomes da Filial do Fornecedor
    lErro = CF("FornecedorProdutoFF_Le_FilialForn", objFornecedorProdutoFF.sProduto, objFornecedor.lCodigo, Codigo_Extrai(FilialCompra), colCodigoNome)
    If lErro <> SUCESSO Then gError 67399
    
    'Se não encontrou nenhuma Filial, erro
    If colCodigoNome.Count = 0 Then gError 67400
    
    If iCodFilial > 0 Then

        For iIndice = 1 To colCodigoNome.Count
            If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                Exit For
            End If
        Next
    
        If iIndice = colCodigoNome.Count Then gError 67401
    
    ElseIf iCodFilial = 0 Then
        iCodFilial = colCodigoNome.Item(1).iCodigo
    End If

    For iIndice = 1 To colCodigoNome.Count
        If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_FilialFornecedor_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
            Exit For
        End If
    Next
            
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col))) = 0 Then
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col) = "Preferencial"
    End If

    Me.Show

    Exit Sub

Erro_objEventoFiliaisFornProd_evSelecao:

    Select Case gErr

        Case 67398, 67399, 67400, 67401

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161515)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
       
       'Torna Frame correspondente ao Tab selecionado visivel
       Frame1(TabStrip1.SelectedItem.Index).Visible = True
       'Torna Frame atual visivel
       Frame1(iFrameAtual).Visible = False
       'Armazena novo valor de iFrameAtual
       iFrameAtual = TabStrip1.SelectedItem.Index

       'Se o Tab selecionado foi o de produtos
       If TabStrip1.SelectedItem.Index = TAB_Produtos And iSelecaoAlterada = REGISTRO_ALTERADO Then
                                        
            'Trata o Tab de Produtos
            lErro = Trata_TabProdutos()
            If lErro <> SUCESSO Then gError 67321
            
            iSelecaoAlterada = 0
       
       End If

    End If

    Exit Sub
    
Erro_TabStrip1_Click:
            
    Select Case gErr
    
        Case 67321
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161516)
    
    End Select
    
    Exit Sub
    
End Sub

Function Trata_TabProdutos() As Long

Dim lErro As Long
Dim colProdutoPtoPedido As New Collection
Dim iIndice As Integer
Dim sProdutoDe As String
Dim sProdutoAte As String
Dim objTipoProduto As ClassTipoDeProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_TabProdutos
    
    Set gobjGeracaoPtoPedido = New ClassGeracaoPtoPedido
    Set gobjGeracaoPtoPedido.colTipoProduto = New Collection
    
    'Obtem os tipos de Produtos selecionados
    For iIndice = 0 To TipoProduto.ListCount - 1
        If TipoProduto.Selected(iIndice) = True Then
            
            Set objTipoProduto = New ClassTipoDeProduto
            
            objTipoProduto.iTipo = TipoProduto.ItemData(iIndice)
            objTipoProduto.sDescricao = TipoProduto.List(iIndice)
            
            gobjGeracaoPtoPedido.colTipoProduto.Add objTipoProduto
        
        End If
    
    Next
    
    'Verifica se existe algum Tipo de Produto selecionado
    If gobjGeracaoPtoPedido.colTipoProduto.Count = 0 Then gError 74867
    
    'Formata o Produto
    lErro = CF("Produto_Formata", CodProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 67252
    
    'Guarda o Produto Inicial e o Final
    gobjGeracaoPtoPedido.sCodProdutoDe = sProdutoFormatado
    
    'Formata o Produto
    lErro = CF("Produto_Formata", CodProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 67256
    
    gobjGeracaoPtoPedido.sCodProdutoAte = sProdutoFormatado
    
    'Lê os Produtos que estão em Ponto de Pedido
    lErro = CF("ProdutoFilial_Le_PontoPedido", gobjGeracaoPtoPedido)
    If lErro <> SUCESSO Then gError 67322
    
    'Preenche o Grid de Produtos
    lErro = GridProdutos_Preenche()
    If lErro <> SUCESSO Then gError 67323
    
    Trata_TabProdutos = SUCESSO

    Exit Function

Erro_Trata_TabProdutos:
    
    Trata_TabProdutos = gErr
    
    Select Case gErr
    
        Case 67252, 67256, 67322, 67323
        
        Case 74867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_TIPOPRODUTO_SELECIONADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161517)
    
    End Select
    
    Exit Function
    
End Function

Function GridProdutos_Preenche() As Long
'Preenche o Grid de Produtos

Dim lErro As Long
Dim sProdutoMascarado As String
Dim objProdutoPtoPedido As ClassProdutoPtoPedido
Dim iIndice As Integer
Dim dSaldoAtual As Double
Dim dQuantEmPedidoCompra As Double
Dim dQuantEmReqCompra As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iCont As Integer
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim dQuantComprar As Double
Dim dSaldoDisponivel As Double
Dim sContaEnxuta As String
Dim dQuantEmpenhada As Double
Dim dQuantEmEstoque As Double
Dim dQuantReservada As Double
Dim iAlmoxarifadoPadrao As Integer
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_GridProdutos_Preenche
    
    Call Grid_Limpa(objGridProdutos)
            
    'Para cada Produto em Ponto de Pedido
    For Each objProdutoPtoPedido In gobjGeracaoPtoPedido.colProdutoPtoPedido
        
        'Calcula a quantidade em transferência
        lErro = CF("SaldoAtualProdComprado", objProdutoPtoPedido.objProduto.sCodigo, dSaldoAtual, dQuantReservada, dQuantEmReqCompra, dQuantEmPedidoCompra, dSaldoDisponivel, dQuantEmpenhada)
        If lErro <> SUCESSO Then gError 67381
                
        'Se a Quantidade em Ponto de Pedido for maior ou igual ao saldo atual
        If objProdutoPtoPedido.objProdutoFilial.dPontoPedido >= dSaldoAtual Then
        
            'Incluido por Leo em 06/02/02
        If iIndice = 100 Then
             iIndice = iIndice
        End If
            'Verifica se existem mais Produtos em colProdutoPtoPedido do que o permitido no grid
            If iIndice >= objGridProdutos.objGrid.Rows - 1 Then

                'Aumenta o nº de linhas permitidas no grid p/ igual ao nº de elemntos da coleção que vai o preencher
                objGridProdutos.objGrid.Rows = objGridProdutos.objGrid.Rows + 100
                Call Grid_Inicializa(objGridProdutos)

            End If

            'Leo até aqui
            
            'Mascara o Produto - Alterado por Wagner
            lErro = Mascara_RetornaProdutoTela(objProdutoPtoPedido.objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 67324
        
            iIndice = iIndice + 1
            
            GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
            GridProdutos.TextMatrix(iIndice, iGrid_DescProduto_Col) = objProdutoPtoPedido.objProduto.sDescricao
            GridProdutos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objProdutoPtoPedido.objProduto.sSiglaUMEstoque
            
            dQuantEmEstoque = dSaldoDisponivel + dQuantReservada
            GridProdutos.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(dQuantEmEstoque)
            GridProdutos.TextMatrix(iIndice, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantReservada)
            GridProdutos.TextMatrix(iIndice, iGrid_QuantEmpenhada_Col) = Formata_Estoque(dQuantEmpenhada)
            GridProdutos.TextMatrix(iIndice, iGrid_QuantEmPedido_Col) = Formata_Estoque(dQuantEmPedidoCompra)
            GridProdutos.TextMatrix(iIndice, iGrid_QuantEmReq_Col) = Formata_Estoque(dQuantEmReqCompra)
            GridProdutos.TextMatrix(iIndice, iGrid_SaldoAtual_Col) = Formata_Estoque(dSaldoAtual)
            GridProdutos.TextMatrix(iIndice, iGrid_PontoPedido_Col) = Formata_Estoque(objProdutoPtoPedido.objProdutoFilial.dPontoPedido)
                        
            dQuantComprar = (objProdutoPtoPedido.objProdutoFilial.dConsumoMedio / 30) * objProdutoPtoPedido.objProdutoFilial.iIntRessup + (objProdutoPtoPedido.objProdutoFilial.dPontoPedido - dSaldoAtual)
            
            If dQuantComprar > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col) = Formata_Estoque(dQuantComprar)
            If objProdutoPtoPedido.objProdutoFilial.dConsumoMedio > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_ConsumoMedio_Col) = Formata_Estoque(objProdutoPtoPedido.objProdutoFilial.dConsumoMedio)
            If objProdutoPtoPedido.objProdutoFilial.iIntRessup > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_IntervaloRessup_Col) = objProdutoPtoPedido.objProdutoFilial.iIntRessup
            GridProdutos.TextMatrix(iIndice, iGrid_LoteEconomico_Col) = Formata_Estoque(objProdutoPtoPedido.objProdutoFilial.dLoteEconomico)
            
            If TipoDestino(TIPO_DESTINO_EMPRESA).Value Then
                If Codigo_Extrai(FilialEmpresa.Text) > 0 Then
                    
                    lErro = CF("AlmoxarifadoPadrao_Le", Codigo_Extrai(FilialEmpresa), objProdutoPtoPedido.objProduto.sCodigo, iAlmoxarifadoPadrao)
                    If lErro <> SUCESSO And lErro <> 23796 Then gError 86167
                    If iAlmoxarifadoPadrao <> 0 Then
                
                        'Lê o Almoxarifado e coloca seu nome Reduzido no Grid
                        objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
                        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25056 Then gError 67375
                       'Se não encontrou, erro
                        If lErro = 25056 Then gError 67376
                    
                        GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
                    End If
                End If
            End If
            
            If Len(Trim(objProdutoPtoPedido.objEstoqueProduto.sContaContabil)) > 0 Then
            
                lErro = Mascara_RetornaContaEnxuta(objProdutoPtoPedido.objEstoqueProduto.sContaContabil, sContaEnxuta)
                If lErro <> SUCESSO Then gError 67458
                
                ContaContabil.PromptInclude = False
                ContaContabil.Text = sContaEnxuta
                ContaContabil.PromptInclude = True
                
                GridProdutos.TextMatrix(iIndice, iGrid_ContaContabil_Col) = ContaContabil.Text
            End If
            
            'Tipo de Tributação
            For iCont = 0 To TipoTributacao.ListCount - 1
                If Codigo_Extrai(TipoTributacao.List(iCont)) = giTipoTributacao Then
                    GridProdutos.TextMatrix(iIndice, iGrid_TipoTributacao_Col) = TipoTributacao.List(iCont)
                    Exit For
                End If
            Next
                
            If Codigo_Extrai(FilialCompra.Text) <> 0 Then
                
                If Codigo_Extrai(FilialCompra) = giFilialEmpresa Then
                    Set objProdutoFilial = objProdutoPtoPedido.objProdutoFilial
                Else
                    Set objProdutoFilial = New ClassProdutoFilial
                    
                    objProdutoFilial.sProduto = objProdutoPtoPedido.objProduto.sCodigo
                    objProdutoFilial.iFilialEmpresa = Codigo_Extrai(FilialCompra)
                    
                    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                    If lErro <> SUCESSO And lErro <> 28261 Then gError 86168
                    
''''                    'Se não encontrou ==> Erro
''''                    If lErro = 28261 Then gError 86169
                
                End If
                
                'Coloca Nome Reduzido do Fornecedor no Grid
                If objProdutoFilial.lFornecedor > 0 Then
        
                    objFornecedor.lCodigo = objProdutoFilial.lFornecedor
                    lErro = CF("Fornecedor_Le", objFornecedor)
                    If lErro <> SUCESSO And lErro <> 12729 Then gError 67377
                    If lErro = 12729 Then gError 67378
        
                    GridProdutos.TextMatrix(iIndice, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
                               
                    'Coloca Filial do Fornecedor no Grid
                    If objProdutoFilial.iFilialForn > 0 Then
                        
                        objFilialFornecedor.iCodFilial = objProdutoFilial.iFilialForn
                        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
                        
                        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                        If lErro <> SUCESSO And lErro <> 12929 Then gError 67379
                        If lErro = 12929 Then gError 67380
                        
                        GridProdutos.TextMatrix(iIndice, iGrid_FilialFornecedor_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
                    End If
                            
                End If
                If objProdutoFilial.lFornecedor > 0 Then
            
                    'Preenche Exclusivo
                    For iCont = 0 To Exclusivo.ListCount - 1
                        If Exclusivo.ItemData(iCont) = FORNECEDOR_PREFERENCIAL Then
                            GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col) = Exclusivo.List(iCont)
                            Exit For
                        End If
                    Next
                    
                End If
            
            End If
            
            GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col) = ""
            
        End If
        
    Next
    
    objGridProdutos.iLinhasExistentes = iIndice
    
    'Se nenhum produto lido está em Ponto de Pedido, Erro
    If iIndice = 0 Then gError 67440
    
    GridProdutos_Preenche = SUCESSO
    
    Exit Function

Erro_GridProdutos_Preenche:

    GridProdutos_Preenche = gErr
    
    Select Case gErr
    
        Case 67324, 67375, 67377, 67379, 67381, 67382, 67383, 86167, 86168
        
        Case 86169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
        
        Case 67376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
        
        Case 67378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
        
        Case 67380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
            
        Case 67440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PRODUTO_SELECIONADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161518)
    
    End Select
    
    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'Quant. Comprar
            Case iGrid_QuantComprar_Col
                lErro = Saida_Celula_QuantComprar(objGridInt)
                If lErro <> SUCESSO Then gError 67325
            
            'Consumo Medio
            Case iGrid_ConsumoMedio_Col
                lErro = Saida_Celula_ConsumoMedio(objGridInt)
                If lErro <> SUCESSO Then gError 67326
            
            'Intervalo de Ressuprimento
            Case iGrid_IntervaloRessup_Col
                lErro = Saida_Celula_IntervaloRessup(objGridInt)
                If lErro <> SUCESSO Then gError 67327
                        
            'Almoxarifado
            Case iGrid_Almoxarifado_Col
                lErro = Saida_Celula_Almoxarifado(objGridInt)
                If lErro <> SUCESSO Then gError 67329
                    
            'ContaContabil
            Case iGrid_ContaContabil_Col
                lErro = Saida_Celula_ContaContabil(objGridInt)
                If lErro <> SUCESSO Then gError 67330
            
            'Tipo Tributação
            Case iGrid_TipoTributacao_Col
                lErro = Saida_Celula_TipoTributacao(objGridInt)
                If lErro <> SUCESSO Then gError 67331
            
            'Fornecedor
            Case iGrid_Fornecedor_Col
                lErro = Saida_Celula_Fornecedor(objGridInt)
                If lErro <> SUCESSO Then gError 67332
                
            'Filial Fornecedor
            Case iGrid_FilialFornecedor_Col
                lErro = Saida_Celula_FilialForn(objGridInt)
                If lErro <> SUCESSO Then gError 67333
                
            'Exclusivo
            Case iGrid_Exclusivo_Col
                lErro = Saida_Celula_Exclusivo(objGridInt)
                If lErro <> SUCESSO Then gError 67334
            
            'Observação
            Case iGrid_Observacao_Col
                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 67335
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 67336

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 67325 To 67336
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161519)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantComprar(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dIntRessup As Double
Dim dPontoPedido As Double
Dim dSaldoAtual As Double
Dim dConsumoMedio As Double

On Error GoTo Erro_Saida_Celula_QuantComprar

    Set objGridInt.objControle = Quantidade

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 67337

        dQuantidade = CDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)

    End If

    'Se a Quantidade foi alterada
    If StrParaDbl(Quantidade.Text) <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_QuantComprar_Col)) Then
    
        'Se o Consumo médio for maior que zero, recalcula o intervalo de Ressuprimento
        If StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ConsumoMedio_Col)) > 0 Then
            
            dPontoPedido = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PontoPedido_Col))
            dSaldoAtual = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SaldoAtual_Col))
            dConsumoMedio = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ConsumoMedio_Col))
            dQuantidade = StrParaDbl(Quantidade.Text)
            
            dIntRessup = (dQuantidade - (dPontoPedido - dSaldoAtual)) / dConsumoMedio
            
            If dIntRessup > 0 Then
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_IntervaloRessup_Col) = Formata_Estoque(dIntRessup)
            Else
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_IntervaloRessup_Col) = ""
            End If
        
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67338
        
    Saida_Celula_QuantComprar = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprar:

    Saida_Celula_QuantComprar = gErr

    Select Case gErr

        Case 67337, 67338
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161520)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ConsumoMedio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ConsumoMédio do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dConsumoMedido As Double
Dim dIntRessup As Double
Dim dPontoPedido As Double
Dim dSaldoAtual As Double
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_ConsumoMedio

    Set objGridInt.objControle = ConsumoMedio

    'Se o ConsumoMedio estiver preenchida
    If Len(Trim(ConsumoMedio.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(ConsumoMedio.Text)
        If lErro <> SUCESSO Then gError 67339

        dConsumoMedido = CDbl(ConsumoMedio.Text)

        'Coloca o valor Formatado na tela
        ConsumoMedio.Text = Formata_Estoque(dConsumoMedido)

    End If
            
    'Se o Consumo médio foi alterado
    If StrParaDbl(ConsumoMedio.Text) <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ConsumoMedio_Col)) Then
    
        'Recalcula a quantidade a comprar
        dPontoPedido = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PontoPedido_Col))
        dSaldoAtual = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SaldoAtual_Col))
        dConsumoMedido = StrParaDbl(ConsumoMedio.Text)
        dIntRessup = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_IntervaloRessup_Col))
        
        dQuantidade = (dConsumoMedido * dIntRessup) + (dPontoPedido - dSaldoAtual)
        
        If dQuantidade > 0 Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_QuantComprar_Col) = Formata_Estoque(dQuantidade)
        Else
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_QuantComprar_Col) = ""
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67340
        
    Saida_Celula_ConsumoMedio = SUCESSO

    Exit Function

Erro_Saida_Celula_ConsumoMedio:

    Saida_Celula_ConsumoMedio = gErr

    Select Case gErr

        Case 67339, 67340
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161521)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_IntervaloRessup(objGridInt As AdmGrid) As Long
'Faz a crítica da célula IntervaloRessup do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dConsumoMedido As Double
Dim dIntRessup As Double
Dim dPontoPedido As Double
Dim dSaldoAtual As Double
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_IntervaloRessup

    Set objGridInt.objControle = IntervaloRessup

    'Se o IntervaloRessup estiver preenchida
    If Len(Trim(IntervaloRessup.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(IntervaloRessup.Text)
        If lErro <> SUCESSO Then gError 67341

        dIntRessup = CDbl(IntervaloRessup.Text)

        'Coloca o valor Formatado na tela
        IntervaloRessup.Text = Formata_Estoque(dIntRessup)

    End If
            
    'Se o Intervalo de Ressuprimento foi alterado
    If StrParaDbl(IntervaloRessup.Text) <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_IntervaloRessup_Col)) Then
    
        'Recalcula a quantidade a comprar
        dPontoPedido = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PontoPedido_Col))
        dSaldoAtual = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SaldoAtual_Col))
        dConsumoMedido = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ConsumoMedio_Col))
        dIntRessup = StrParaDbl(IntervaloRessup.Text)
        
        dQuantidade = (dConsumoMedido * dIntRessup) + (dPontoPedido - dSaldoAtual)
        
        If dQuantidade > 0 Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_QuantComprar_Col) = Formata_Estoque(dQuantidade)
        Else
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_QuantComprar_Col) = ""
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67342
        
    Saida_Celula_IntervaloRessup = SUCESSO

    Exit Function

Erro_Saida_Celula_IntervaloRessup:

    Saida_Celula_IntervaloRessup = gErr

    Select Case gErr

        Case 67341, 67342
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161522)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Almoxarifado do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado foi preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then
        'Lê o Almoxarifado
        
        'Formata o Produto
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 84568 '67343
        
        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 84564 '67344
        If lErro = 25130 Then gError 84565 '67345
        If lErro = 25135 Then gError 84566 '67346
        If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 84567 '86099
        
        
'''        objEstoqueProduto.sProduto = sProdutoFormatado
'''        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84569 '67347

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 84568, 84564, 84569
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84565
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 84566
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

         Case 84567
            
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161523)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaEnxuta As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabil

    Set objGridProdutos.objControle = ContaContabil

    'Se a Conta Contábil foi preenchida
    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabil.ClipText, objPlanoConta, MODULO_COMPRAS)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 67348

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaEnxuta = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
            If lErro <> SUCESSO Then gError 67349

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_COMPRAS)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 67350

            'conta não cadastrada
            If lErro = 5700 Then gError 67351

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67352

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = gErr

    Select Case gErr

        Case 67348, 67350, 67352
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 67349
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 67351
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161524)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoTributacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo de Tributação que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TipoTributacao

    Set objGridInt.objControle = TipoTributacao

    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TipoTributacao_Col) = TipoTributacao.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67355

    Saida_Celula_TipoTributacao = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoTributacao:

    Saida_Celula_TipoTributacao = gErr

    Select Case gErr

        Case 67355
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161525)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Fornecedor(objGridInt As AdmGrid) As Long
'faz a critica da celula fornecedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iFilialEmpresa As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sFornecedor As String

On Error GoTo Erro_Saida_Celula_Fornecedor

    Set objGridInt.objControle = Fornecedor

    sFornecedor = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col)
    
    'Se o fornecedor foi preenchido
    If Len(Trim(Fornecedor.ClipText)) > 0 Then

        'Verifica se o fornecedor está cadastrado
        lErro = TP_Fornecedor_Grid(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 25611 And lErro <> 25613 And lErro <> 25616 And lErro <> 25619 Then gError 67356

        'Fornecedor não cadastrado
        'Nome Reduzido
        If lErro = 25611 Then gError 67359

        'Codigo
        If lErro = 25613 Then gError 67360

        'CGC/CPF
        If lErro = 25616 Or lErro = 25619 Then gError 67361

        If sFornecedor <> objFornecedor.sNomeReduzido Then
        
            'Formata o Produto
            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 67358
            
            iFilialEmpresa = Codigo_Extrai(FilialCompra)
            
            'Lê coleção de códigos e nomes da Filial do Fornecedor
            lErro = CF("FornecedorProdutoFF_Le_FilialForn", sProdutoFormatado, objFornecedor.lCodigo, iFilialEmpresa, colCodigoNome)
            If lErro <> SUCESSO Then gError 61635
    
            'Se não encontrou nenhuma Filial, erro
            If colCodigoNome.Count = 0 Then gError 67362
    
            If iCodFilial > 0 Then
    
                For iIndice = 1 To colCodigoNome.Count
                    If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                        Exit For
                    End If
                Next
    
                If iIndice = colCodigoNome.Count Then gError 67363
    
            ElseIf iCodFilial = 0 Then
                iCodFilial = colCodigoNome.Item(1).iCodigo
            End If
    
            For iIndice = 1 To colCodigoNome.Count
                If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_FilialFornecedor_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
                    Exit For
                End If
            Next
        
        End If
    
        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col))) = 0 Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col) = "Preferencial"
        End If
        
    Else
        
        'Limpa a Filial e Exclusividade Correspondente
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_FilialFornecedor_Col) = ""
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67357

    Saida_Celula_Fornecedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Fornecedor:

    Saida_Celula_Fornecedor = gErr

    Select Case gErr

        Case 67356, 67357, 67358
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 67359 'Fornecedor com Nome Reduzido %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", Fornecedor.Text)
            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 67360 'Fornecedor com código %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 67361 'Fornecedor com CGC/CPF %s não encontado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_3", Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 67362
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR_FILIALCOMPRA", gErr, objFornecedor.sNomeReduzido, sProdutoFormatado)
            Fornecedor.Text = ""
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 67363
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", gErr, iCodFilial, objFornecedor.sNomeReduzido, sProdutoFormatado)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161526)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialForn(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sNomeFilial As String

On Error GoTo Erro_Saida_Celula_FilialForn

    Set objGridInt.objControle = FilialForn
    
    'Verifica se a filial foi preenchida
    If Len(Trim(FilialForn.Text)) > 0 Then

        'Verifica se não é uma filial selecionada
        If Not FilialForn.Text = FilialForn.List(FilialForn.ListIndex) Then

            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialForn, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 67364
    
            'Se nao encontra o ítem com o código informado
            If lErro = 6730 Then
    
                'Verifica se o Fornecedor foi preenchido
                If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col))) = 0 Then gError 67369
    
                lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 67368
                
                sFornecedor = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col)
                objFornecedorProdutoFF.iFilialForn = iCodigo
                objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa
                objFornecedorProdutoFF.sProduto = sProdutoFormatado
    
                'Pesquisa se existe filial com o codigo extraido
                lErro = CF("FornecedorProdutoFF_Le_NomeRed", sFornecedor, sNomeFilial, objFornecedorProdutoFF)
                If lErro <> SUCESSO And lErro <> 61780 Then gError 67365
    
                'Se não encontrou a Filial do Fornecedor
                If lErro = 61780 Then
    
                    'Lê FilialFornecedor do BD
                    objFilialFornecedor.iCodFilial = iCodigo
                    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
                    If lErro <> SUCESSO And lErro <> 18272 Then gError 67366
    
                    'Se não encontrou, pergunta se deseja criar
                    If lErro = 18272 Then
                        gError 67370
                    
                    'Se encontrou, erro
                    Else
                        gError 67372
                    End If
                
                'Se encontrou a Filial do Fornecedor
                Else
    
                    'coloca na tela
                    FilialForn.Text = iCodigo & SEPARADOR & sNomeFilial
    
                End If
    
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 67371

        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError (67367)

    Saida_Celula_FilialForn = SUCESSO
    
    Exit Function

Erro_Saida_Celula_FilialForn:

    Saida_Celula_FilialForn = gErr
    
    Select Case gErr

        Case 67364, 67365, 67366, 67367, 67368
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 67369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_FORNECEDOR_NAO_PREENCHIDO", gErr, GridProdutos.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 67370

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then

                objFornecedor.sNomeReduzido = Fornecedor.Text

                'Lê Fornecedor no BD
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)

                'Se achou o Fornecedor --> coloca o codigo em objFilialFornecedor
                If lErro = SUCESSO Then objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 67371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_NAO_ENCONTRADA_ASSOCIADA", gErr, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 67372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", gErr, objFilialFornecedor.iCodFilial, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161527)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Exclusivo(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Exclusivo do grid que está deixando de ser a corrente
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Exclusivo

    Set objGridInt.objControle = Exclusivo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67373

    Saida_Celula_Exclusivo = SUCESSO

    Exit Function

Erro_Saida_Celula_Exclusivo:

    Saida_Celula_Exclusivo = gErr

    Select Case gErr

        Case 67373
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161528)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Observacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67374

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 67374
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161529)

    End Select

    Exit Function

End Function

'Tratamento do GridProdutos
Private Sub GridProdutos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
    End If

End Sub

Private Sub GridProdutos_EnterCell()

    Call Grid_Entrada_Celula(objGridProdutos, iAlterado)

End Sub

Private Sub GridProdutos_GotFocus()

    Call Grid_Recebe_Foco(objGridProdutos)

End Sub

Private Sub GridProdutos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
    End If

End Sub

Private Sub GridProdutos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridProdutos_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos)

    Exit Sub

Erro_GridProdutos_KeyDown:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161530)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos_LeaveCell()

    Call Saida_Celula(objGridProdutos)

End Sub

Private Sub GridProdutos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridProdutos)
End Sub

Private Sub GridProdutos_RowColChange()

    Call Grid_RowColChange(objGridProdutos)

End Sub

Private Sub GridProdutos_Scroll()

    Call Grid_Scroll(objGridProdutos)

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ConsumoMedio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConsumoMedio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub ConsumoMedio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub ConsumoMedio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = ConsumoMedio
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub IntervaloRessup_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IntervaloRessup_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub IntervaloRessup_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub IntervaloRessup_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = IntervaloRessup
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub ContaContabil_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = ContaContabil
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub Fornecedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Fornecedor
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FilialForn_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialForn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub FilialForn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub FilialForn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = FilialForn
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Exclusivo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Exclusivo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub Exclusivo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub Exclusivo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Exclusivo
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TipoTributacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTributacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub TipoTributacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = TipoTributacao
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

'Entra com Intervalo Ressup e Consumo Medio de cada Produto,
'mas deixa alterar. Alterando, reflete na Quant Comprar.
'Se alterar Quant Comprar, reflete no Intervalo Ressup.
'ConsumoMedio e IntervaloRessup alterados aqui NÃO refletem no BD.

'Digitador é quem está digitando.
'Requisitante vai ser AUTO (automático).
'Destino é a própria FilialEmpresa (essa tela só entra a nível de Filial)

'SE colocar DataEnvio então já fica enviado p/ SetorCompras

'**** inicio do trecho a ser copiado *****

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
            
        If Me.ActiveControl Is CodProdutoDe Then
            Call ProdutoDeLabel_Click
        ElseIf Me.ActiveControl Is CodProdutoAte Then
            Call ProdutoAteLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call BotaoProdutoFiliaisForn_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoAlmoxarifados_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call BotaoPlanoConta_Click
        End If
        
    End If
    
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Requisições - Ponto de Pedido"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoReqPtoPedido"

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

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub


Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)
'Rotina que habilita a entrada na celula

Dim lErro As Long
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor
Dim iCodigo As Integer
Dim iSelecionado As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If GridProdutos.Row = 0 Then Exit Sub


    iSelecionado = StrParaInt(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Selecionado_Col))

    'Verifica se produto está preenchido
    sCodProduto = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 61848

    Select Case objControl.Name
        
        Case Quantidade.Name, ContaContabil.Name, Fornecedor.Name, Observacao.Name, IntervaloRessup.Name, ContaContabil.Name, ConsumoMedio.Name

            If iSelecionado = GRID_CHECKBOX_ATIVO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case Exclusivo.Name
        
            If iSelecionado = GRID_CHECKBOX_ATIVO And Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
                
        Case TipoTributacao.Name
            
            If iSelecionado = GRID_CHECKBOX_ATIVO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        Case FilialForn.Name
        
            If iSelecionado = GRID_CHECKBOX_INATIVO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
                
                'Se o Fornecedor não está preenchido
                If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col))) = 0 Then
                    
                    'Desabilita combo de Filiais
                    objControl.Enabled = False
                
                Else
                    
                    objFornecedor.sNomeReduzido = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col)
                    
                    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                    If lErro <> SUCESSO And lErro <> 6681 Then gError (65590)
                    If lErro = 6681 Then gError (65591)
                    
                    lErro = CF("FornecedorProdutoFF_Le_FilialForn", sProdutoFormatado, objFornecedor.lCodigo, giFilialEmpresa, colCodigoNome)
                    If lErro <> SUCESSO Then gError (66592)
                    
                    If colCodigoNome.Count = 0 Then gError (65593)
                
                    If Len(Trim(FilialForn.Text)) = 0 Then
                        iCodigo = colCodigoNome.Item(1).iCodigo
                    Else
                        iCodigo = Codigo_Extrai(FilialForn.Text)
                    End If

                    FilialForn.Clear
                    
                    Call CF("Filial_Preenche", FilialForn, colCodigoNome)
                    Call CF("Filial_Seleciona", FilialForn, iCodigo)
                
                End If
                
            End If
            
        Case Almoxarifado.Name

            'Verifica se o detino é a empresa
            If iFrameTipoDestinoAtual <> TIPO_DESTINO_EMPRESA Or Len(Trim(FilialEmpresa.Text)) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 61848, 61849, 61850, 65590, 66592

        Case 65591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case 65593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR", gErr, objFornecedor.sNomeReduzido, sProdutoFormatado)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161531)

    End Select

    Exit Sub

End Sub


Private Sub ProdutoDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoDeLabel, Source, X, Y)
End Sub

Private Sub ProdutoDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoDeLabel, Button, Shift, X, Y)
End Sub

Private Sub ProdutoAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoAteLabel, Source, X, Y)
End Sub

Private Sub ProdutoAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoAteLabel, Button, Shift, X, Y)
End Sub

Public Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornec.Text = objFornecedor.lCodigo
    Call Fornec_Validate(bCancel)

    Me.Show

End Sub

Function AlmoxarifadoPadrao_Preenche(iCodFilial As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim iAlmoxarifadoPadrao As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoPadrao_Preenche
    
    'Para cada linha do Grid
    For iIndice = 1 To objGridProdutos.iLinhasExistentes
        
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    
            objProduto.sCodigo = sProdutoFormatado
            
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 84570 '65558)
    
            'Se não encontrou o Produto --> Erro
            If lErro = 28030 Then gError 84571 '(65560)
            
            'Se o Produto possui Estoque
            If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                'Lê dados do seu almoxarifado Padrão
                lErro = CF("AlmoxarifadoPadrao_Le", iCodFilial, objProduto.sCodigo, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then gError 84572 '(61710)
    
                'Se encontrou
                If lErro = SUCESSO And iAlmoxarifadoPadrao <> 0 Then
    
                    objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
    
                    'Lê os dados do Almoxarifado a partir do código passado
                    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25056 Then gError 84573 '(65559)
            
                    'Se não encontrou, erro
                    If lErro = 25056 Then gError 84574 '(65561)
            
                    'Coloca o Nome Reduzido na Coluna Almoxarifado
                    GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
                Else
                    GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = ""
                End If
                
            
            End If
        
        End If
        
    Next
    
    AlmoxarifadoPadrao_Preenche = SUCESSO
    
    Exit Function
    
Erro_AlmoxarifadoPadrao_Preenche:

    AlmoxarifadoPadrao_Preenche = gErr
    
    Select Case gErr
        
        Case 84570, 84573, 84572
        
        Case 84571
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 84574
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161532)
    
    End Select
    
    Exit Function
    
End Function


'Function SaldoAtualProdComprado(sCodProduto As String, dSaldoAtual As Double, Optional dQuantReservada As Double, Optional dQuantEmReq As Double, Optional dQuantEmPed As Double, Optional dSaldoDisponivel As Double, Optional dQuantEmpenhada As Double)
''Calcula o Saldo Atual para o produto comprado passado por parâmetro.
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objEstoqueProduto As New ClassEstoqueProduto
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_SaldoAtualProdComprado
'
'    'Lê a quantidade do produto em Requisição
'    lErro = CF("Produto_Le_QuantRequisitada", sCodProduto, dQuantEmReq)
'    If lErro <> SUCESSO Then gError 84584
'
'    objEstoqueProduto.sProduto = sCodProduto
'    'Lê o estoque do produto para a filial
'    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
'    If lErro <> SUCESSO And lErro <> 52979 Then gError 84585
'
'    'Calcula a quantidade do produto em Pedido de compras
'    lErro = CF("Produto_Le_QuantEmPedidoCompra", sCodProduto, dQuantEmPed)
'    If lErro <> SUCESSO Then gError 84586
'
'    'Calcula o Saldo Atual do produto
'    With objEstoqueProduto
'        dSaldoDisponivel = .dQuantDisponivel + .dQuantRecIndl
'        dQuantEmpenhada = .dQuantEmpenhada
'        dQuantReservada = .dQuantReservada
'        dSaldoAtual = objEstoqueProduto.dQuantReservada + dSaldoDisponivel + dQuantEmPed + dQuantEmReq - .dQuantEmpenhada
'    End With
'
'    SaldoAtualProdComprado = SUCESSO
'
'    Exit Function
'
'Erro_SaldoAtualProdComprado:
'
'    SaldoAtualProdComprado = gErr
'
'    Select Case gErr
'
'        Case 84584, 84585, 84586
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161533)
'
'    End Select
'
'    Exit Function
'
'End Function
