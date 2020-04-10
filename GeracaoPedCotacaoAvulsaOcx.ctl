VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl GeracaoPedCotacaoAvulsaOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8295
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   705
      Width           =   16545
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
         Height          =   375
         Left            =   375
         TabIndex        =   13
         Top             =   7695
         Width           =   1350
      End
      Begin VB.ListBox TipoProduto 
         Height          =   960
         Left            =   3855
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   270
         Width           =   2775
      End
      Begin VB.CommandButton BotaoMarcarTodosTipos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Index           =   0
         Left            =   6780
         Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodosTipos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Index           =   0
         Left            =   6780
         Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   705
         Width           =   1425
      End
      Begin VB.Frame Frame4 
         Caption         =   "Produtos"
         Height          =   6195
         Left            =   315
         TabIndex        =   7
         Top             =   1320
         Width           =   15945
         Begin VB.TextBox DescricaoProd 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3855
            MaxLength       =   50
            TabIndex        =   9
            Top             =   4095
            Width           =   7000
         End
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   9375
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   4095
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   7965
            TabIndex        =   11
            Top             =   4140
            Width           =   1305
            _ExtentX        =   2302
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
            Left            =   2295
            TabIndex        =   8
            Top             =   4065
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos 
            Height          =   5340
            Left            =   360
            TabIndex        =   12
            Top             =   480
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   9419
            _Version        =   393216
            Rows            =   10
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Label Comprador 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1515
         TabIndex        =   3
         Top             =   375
         Width           =   2145
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
         Left            =   390
         TabIndex        =   2
         Top             =   405
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   3855
         TabIndex        =   65
         Top             =   30
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8250
      Index           =   2
      Left            =   165
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   16545
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   660
         Left            =   210
         Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7485
         Width           =   2400
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   660
         Left            =   2715
         Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":3216
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   7485
         Width           =   2400
      End
      Begin VB.CommandButton BotaoFornecedores 
         Caption         =   "Filial Fornecedor ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   6090
         Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":43F8
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7485
         Width           =   2175
      End
      Begin VB.ComboBox Ordenacao 
         Height          =   315
         ItemData        =   "GeracaoPedCotacaoAvulsaOcx.ctx":519A
         Left            =   2100
         List            =   "GeracaoPedCotacaoAvulsaOcx.ctx":519C
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   60
         Width           =   2325
      End
      Begin VB.Frame Frame15 
         Caption         =   "Fornecedores"
         Height          =   6930
         Left            =   210
         TabIndex        =   17
         Top             =   435
         Width           =   16290
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6765
            MaxLength       =   100
            TabIndex        =   32
            Top             =   3225
            Width           =   1740
         End
         Begin MSMask.MaskEdBox PrazoEntrega 
            Height          =   225
            Left            =   2055
            TabIndex        =   28
            Top             =   3000
            Width           =   1410
            _ExtentX        =   2487
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
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   21
            Top             =   360
            Width           =   4000
         End
         Begin VB.CheckBox Escolhido 
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
            Left            =   480
            TabIndex        =   19
            Top             =   285
            Width           =   960
         End
         Begin MSMask.MaskEdBox TipoFrete 
            Height          =   225
            Left            =   7800
            TabIndex        =   25
            Top             =   420
            Visible         =   0   'False
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFornecedor 
            Height          =   225
            Left            =   5625
            TabIndex        =   23
            Top             =   360
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
         Begin MSMask.MaskEdBox UltimaCotacao 
            Height          =   225
            Left            =   6840
            TabIndex        =   24
            Top             =   225
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataUltimaCotacao 
            Height          =   225
            Left            =   75
            TabIndex        =   26
            Top             =   3015
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Prod 
            Height          =   225
            Left            =   1125
            TabIndex        =   20
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
         Begin MSMask.MaskEdBox SaldoTitulos 
            Height          =   225
            Left            =   1530
            TabIndex        =   59
            Top             =   3780
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CondicaoPagto 
            Height          =   225
            Left            =   75
            TabIndex        =   60
            Top             =   3795
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   30
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataUltimaCompra 
            Height          =   225
            Left            =   1125
            TabIndex        =   27
            Top             =   3135
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Forn 
            Height          =   225
            Left            =   3870
            TabIndex        =   22
            Top             =   315
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridFornecedores 
            Height          =   6390
            Left            =   105
            TabIndex        =   18
            Top             =   390
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   11271
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox QuantPedida 
            Height          =   225
            Left            =   4380
            TabIndex        =   30
            Top             =   3015
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   225
            Left            =   5475
            TabIndex        =   31
            Top             =   3180
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UMCompra 
            Height          =   225
            Left            =   3360
            TabIndex        =   29
            Top             =   3180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
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
         Left            =   960
         TabIndex        =   15
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame14"
      Height          =   8205
      Index           =   3
      Left            =   210
      TabIndex        =   36
      Top             =   720
      Visible         =   0   'False
      Width           =   16485
      Begin VB.CommandButton BotaoGeraPedidos 
         Caption         =   "Gera Pedidos de Cotação"
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
         Left            =   195
         TabIndex        =   52
         Top             =   6330
         Width           =   3690
      End
      Begin VB.Frame Frame2 
         Caption         =   "Destino"
         Height          =   2055
         Left            =   180
         TabIndex        =   45
         Top             =   4095
         Width           =   6435
         Begin VB.Frame FrameDestino 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   675
            Index           =   0
            Left            =   645
            TabIndex        =   61
            Top             =   1290
            Width           =   3645
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               Left            =   1245
               TabIndex        =   51
               Top             =   120
               Width           =   2160
            End
            Begin VB.Label FilialEmpresaLabel 
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
               TabIndex        =   50
               Top             =   195
               Width           =   465
            End
         End
         Begin VB.Frame FrameDestino 
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   570
            TabIndex        =   62
            Top             =   1290
            Visible         =   0   'False
            Width           =   3630
            Begin MSMask.MaskEdBox Fornecedor 
               Height          =   300
               Left            =   1290
               TabIndex        =   64
               Top             =   0
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.ComboBox FilialForn 
               Height          =   315
               Left            =   1275
               TabIndex        =   63
               Top             =   360
               Width           =   2160
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
               Height          =   195
               Left            =   165
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   66
               Top             =   60
               Width           =   1035
            End
            Begin VB.Label FilialLabel 
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
               Left            =   705
               TabIndex        =   67
               Top             =   405
               Width           =   465
            End
         End
         Begin VB.CheckBox Destino 
            Caption         =   "Seleciona Destino"
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
            Left            =   285
            TabIndex        =   46
            Top             =   270
            Width           =   2205
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo"
            Height          =   585
            Left            =   240
            TabIndex        =   47
            Top             =   570
            Width           =   4305
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Filial Empresa"
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
               Left            =   480
               TabIndex        =   48
               Top             =   240
               Width           =   1515
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Fornecedor"
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
               Index           =   1
               Left            =   2535
               TabIndex        =   49
               Top             =   240
               Width           =   1335
            End
         End
      End
      Begin VB.CommandButton BotaoImprimePedidos 
         Caption         =   "Gera e Imprime Pedidos de Cotação "
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
         Left            =   195
         TabIndex        =   53
         Top             =   6930
         Width           =   3690
      End
      Begin VB.ListBox CondPagtos 
         Height          =   6690
         Left            =   7110
         TabIndex        =   55
         Top             =   885
         Width           =   2985
      End
      Begin VB.Frame Frame5 
         Caption         =   "Geração"
         Height          =   3465
         Left            =   150
         TabIndex        =   37
         Top             =   420
         Width           =   6435
         Begin MSMask.MaskEdBox CondPagto 
            Height          =   315
            Left            =   2175
            TabIndex        =   42
            Top             =   735
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2955
            Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":519E
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Numeração Automática"
            Top             =   345
            Width           =   300
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   2085
            Left            =   2175
            TabIndex        =   44
            Top             =   1155
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   3678
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   "_"
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto A Prazo:"
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
            Left            =   315
            TabIndex        =   41
            Top             =   795
            Width           =   1785
         End
         Begin VB.Label Cotacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2175
            TabIndex        =   39
            Top             =   345
            Width           =   795
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
            Left            =   1170
            TabIndex        =   43
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Nº Geração:"
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
            Left            =   1035
            TabIndex        =   38
            Top             =   375
            Width           =   1065
         End
      End
      Begin VB.Label Label40 
         Caption         =   "Condições de Pagamento"
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
         Left            =   7155
         TabIndex        =   54
         Top             =   615
         Width           =   2265
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   15600
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   30
      Width           =   1170
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   90
         Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":5288
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "GeracaoPedCotacaoAvulsaOcx.ctx":57BA
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8745
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   16770
      _ExtentX        =   29580
      _ExtentY        =   15425
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fornecedores"
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
Attribute VB_Name = "GeracaoPedCotacaoAvulsaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Condicao de pagamentos:
'Só pode ter duas
'Uma é a vista
'A outra o usuário escolhe

Private WithEvents objEventoProdutos As AdmEvento
Attribute objEventoProdutos.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoBotaoProdutos As AdmEvento
Attribute objEventoBotaoProdutos.VB_VarHelpID = -1

'Grid Produtos
Dim objGridProdutos As AdmGrid
Dim iGrid_Detalhe_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProd_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quantidade_Col As Integer

'Grid Fornecedores
Dim objGridFornecedores As AdmGrid
Dim iGrid_Escolhido_Col As Integer
Dim iGrid_Prod_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_Forn_Col As Integer
Dim iGrid_FilialFornecedor_Col As Integer
Dim iGrid_UltimaCotacao_Col As Integer
Dim iGrid_TipoFrete_Col As Integer
Dim iGrid_DataUltimaCotacao_Col As Integer
Dim iGrid_DataUltimaCompra_Col As Integer
Dim iGrid_PrazoEntrega_Col As Integer
Dim iGrid_UMCompra_Col As Integer
Dim iGrid_QuantPedida_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_CondicaoPagto_Col As Integer
Dim iGrid_SaldoTitulos_Col As Integer
Dim iGrid_Observacao_Col As Integer

'Variáveis Globais
Dim giAlterado As Integer
Dim giFrameAtual As Integer
Dim giFrameDestinoAtual As Integer
Dim gsOrdenacao As String
Dim asOrdenacao(1) As String
Dim asOrdenacaoString(1) As String
Dim giFornecedorAlterado As Integer
Dim iAlterado As Integer
Dim gobjCotacao As ClassCotacao
Dim gcolPedidoCotacao As Collection


Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todas CheckBox do GridFornecedores

Dim iLinha As Integer
Dim objBaixaPedCompra As New ClassBaixaPedCompra

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridFornecedores.iLinhasExistentes

        'Desmarca na tela a linha em questão
        GridFornecedores.TextMatrix(iLinha, iGrid_Escolhido_Col) = GRID_CHECKBOX_INATIVO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridFornecedores)

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodosTipos_Click(Index As Integer)
'Desmarca todas as checkbox da ListBox TipoProduto

Dim iIndice As Integer

    'Percorre todas as checkbox de TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1

        'Desmarca na tela o tipo de produto em questão
        TipoProduto.Selected(iIndice) = False

    Next

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoFornecedores_Click()

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoFornecedores_Click

    'Verifica se a linha do GridFornecedores é uma linha existente
    If GridFornecedores.Row > 0 And GridFornecedores.Row <= objGridFornecedores.iLinhasExistentes Then

        'Coloca Codigo do Fornecedor e da Filial em objFilialFornecedor
        objFilialFornecedor.iCodFilial = Codigo_Extrai(GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_FilialFornecedor_Col))

        objFornecedor.sNomeReduzido = GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_Forn_Col)

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 68341

        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

        'Chama a tela FiliaisFornecedores
        Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)

    'Se a linha do GridFornecedores não for uma linha existente==>erro
    ElseIf GridFornecedores.Row > objGridFornecedores.iLinhasExistentes Then gError 63099

    End If

    Exit Sub

Erro_BotaoFornecedores_Click:

    Select Case gErr

        Case 68341
            'Erro tratado na rotina chamada

        Case 63099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_FORN_LINHA_NAO_SELECIONADA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161279)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGeraPedidos_Click()
'Gera os Pedidos de Cotacao Avulsa

Dim lErro As Long

On Error GoTo Erro_BotaoGeraPedidos_Click

    'Gera o Pedido de Cotacao Avulsa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 63106
    
    Exit Sub

Erro_BotaoGeraPedidos_Click:

    Select Case gErr

        Case 63106
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161280)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimePedidos_Click()
'Imprime os Pedidos de Cotacao Avulsa

Dim lErro As Long
Dim objPedidoCotacao As ClassPedidoCotacao

On Error GoTo Erro_BotaoImprimePedidos_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 76063
    
    If gobjCotacao Is Nothing Then gError 63148

   'Imprime os Pedidos de Cotacao
    lErro = PedidosCotacao_Imprime(gobjCotacao)
    If lErro <> SUCESSO Then gError 63149

    'Para cada Pedido de Cotação da coleção de pedidos
    For Each objPedidoCotacao In gcolPedidoCotacao

        'Atualiza data de emissao no BD para a data atual
        lErro = CF("PedidoCotacao_Atualiza_DataEmissao", objPedidoCotacao)
        If lErro <> SUCESSO And lErro <> 56348 Then gError 89860

    Next

    Call Limpa_Tela_Cotacao
    
    Exit Sub

Erro_BotaoImprimePedidos_Click:

    Select Case gErr

        Case 63148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PED_COTACAO_NAO_GERADO", gErr)

        Case 63149, 76063, 89860
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161281)

    End Select

    Exit Sub

End Sub

Function PedidosCotacao_Imprime(gobjCotacao As ClassCotacao) As Long
'Chama a impressao de pedidos de cotacao

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String, sBuffer As String
Dim lErro As Long

On Error GoTo Erro_PedidosCotacao_Imprime

    lErro = objRelatorio.ExecutarDireto("Geracao de Pedido de Cotacao Avulsa", "COTACAO.NumIntDoc=@NCOTACAO", 1, "COTACAO", "NCOTACAO", gobjCotacao.lNumIntDoc)
    If lErro <> SUCESSO Then gError 63200

    PedidosCotacao_Imprime = SUCESSO

    Exit Function

Erro_PedidosCotacao_Imprime:

    PedidosCotacao_Imprime = gErr

    Select Case gErr

        Case 63200
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161282)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 63107

    'Limpa a tela
    Call Limpa_Tela_Cotacao

    Set gobjCotacao = Nothing

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 63107
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161283)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_Cotacao()
'Limpa os campos da Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_Cotacao

    Call Limpa_Tela(Me)

    'Desseleciona todos os tipos de produtos da listbox TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = False
    Next

    'Limpa os Grids da tela
    Call Grid_Limpa(objGridProdutos)
    Call Grid_Limpa(objGridFornecedores)

    'Limpa os outros campos da tela
    Destino.Value = vbUnchecked
    FilialForn.Clear
    Cotacao.Caption = ""
    

    Exit Sub

Erro_Limpa_Tela_Cotacao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161284)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todas CheckBox do GridFornecedores

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridFornecedores.iLinhasExistentes

        'Marca na tela a linha em questão
        GridFornecedores.TextMatrix(iLinha, iGrid_Escolhido_Col) = GRID_CHECKBOX_ATIVO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridFornecedores)

    Exit Sub

End Sub

Private Sub BotaoMarcarTodosTipos_Click(Index As Integer)
'Marca todas as checkbox da ListBox TipoProduto

Dim iIndice As Integer

    'Percorre todas as checkbox de TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1

        'Marca na tela o bloqueio em questão
        TipoProduto.Selected(iIndice) = True

    Next

End Sub

Private Sub BotaoProdutos_Click()
'Chama a tela ProdutoCompraTipoLista, de acordo com giFilialEmpresa e colSelecao

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim colTipoProduto As Collection
Dim iIndice As Integer
Dim iSelecao As Integer
Dim colCampoValor As New AdmColCampoValor
Dim vParametro As Variant
Dim sSelecao As String
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se existe alguma linha do GridProdutos selecionada
    If GridProdutos.Row <> 0 Then

'        'Verifica se o produto da linha selecionada está preenchido
'        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) > 0 Then
'
'            'Passa o codigo do produto para o formato do BD
'            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError 63095
'
'            'Coloca o código formatado no objProduto
'            objProduto.sCodigo = sProdutoFormatado
'        End If

        '###############################################
        'Inserido por Wagner 05/05/06
        If Me.ActiveControl Is Produto Then
            sProduto1 = Produto.Text
        Else
            sProduto1 = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col)
        End If
        
        lErro = CF("Produto_Formata", sProduto1, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 177421
        
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
        
        objProduto.sCodigo = sProdutoFormatado
        '###############################################
   
        
        'Monta FiltroSQL com a colecao de TiposProduto escolhida na ListBox
        For iIndice = 0 To TipoProduto.ListCount - 1

            'Verifica se o TipoProduto foi selecionado
            If TipoProduto.Selected(iIndice) = True Then

                If iSelecao = 0 Then sSelecao = "Tipo = ?"
                iSelecao = iSelecao + 1

                'Adiciona o Filtro
                colSelecao.Add (TipoProduto.ItemData(iIndice))

                If iSelecao > 1 Then sSelecao = sSelecao & " OR Tipo = ?"
                
            End If
            
        Next
            
        'Verifica se nenhum tipo de produto foi selecionado
        If iSelecao = 0 Then gError 74869
        
        'Chama a tela ProdutoCompraTipoLista
        Call Chama_Tela("ProdutoCompraTipoLista", colSelecao, objProduto, objEventoBotaoProdutos, sSelecao)

    End If

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 63095, 177421
            'Erro tratado na rotina chamada

        Case 74869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_SELECIONADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161285)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCotacao As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtem o proximo codigo de cotacao disponivel para esta FilialEmpresa
    lErro = CF("Cotacao_Automatica", lCotacao)
    If lErro <> SUCESSO Then gError 68090

    'Coloca o Código de Cotacao obtido na tela
    Cotacao.Caption = lCotacao

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 68090
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161286)
    
    End Select

    Exit Sub


End Sub

Private Sub CondPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CondPagto_Validate

    'Verifica se CondPagto foi preenchida
    If Len(Trim(CondPagto.Text)) = 0 Then Exit Sub

    If CondPagto.Text = CondPagtos.Text Then Exit Sub

    lErro = List_Seleciona(CondPagto, CondPagtos, iCodigo)
    If lErro <> SUCESSO And lErro <> 25635 And lErro <> 25636 Then gError 63124


    'Se não encontrar Codigo
    If lErro = 25635 Or lErro = SUCESSO Then
    
            If iCodigo = COD_A_VISTA Then Error 62704

            objCondicaoPagto.iCodigo = iCodigo

            'Lê Condicao Pagamento no BD
            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError 63125
            If lErro = 19205 Then gError 63126

            'Testa se pode ser usada em Contas a Pagar
            If objCondicaoPagto.iEmPagamento = 0 Then gError 63127

            'Coloca na Tela
            CondPagto.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida

    End If

    'Nao encontrou o valor que era String
    If lErro = 25636 Then gError 63128

    Exit Sub

Erro_CondPagto_Validate:

    Cancel = True

    Select Case gErr
        
        Case 62704
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_NAO_DISPONIVEL", Err)
        
        Case 63124, 63125, 63250 'Tratado na Rotina chamada

        Case 63126
            'Avisa de deseja criar nova Condicao de Pagamento
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 63127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case 63128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondPagto.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161287)

    End Select

    Exit Sub

End Sub

Private Sub CondPagtos_DblClick()

Dim lErro As Long

On Error GoTo Erro_CondPagtos_DblClick

    'Coloca o texto de CondPagtos em CondPagto
    CondPagto.Text = CondPagto_Extrai(CondPagtos)
    Call CondPagto_Validate(bSGECancelDummy)

    Exit Sub

Erro_CondPagtos_DblClick:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161288)

    End Select

    Exit Sub

End Sub


Private Sub DescricaoProd_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub DescricaoProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub DescricaoProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = DescricaoProd
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Destino_Click()
'Habilita ou desabilita os componentes do FrameDestino dependendo se Destino está marcado ou não

Dim lErro As Long
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Destino_Click

    'Verifica se Destino está desmarcado
    If Destino.Value = vbUnchecked Then
        For iIndice = 0 To 1
            'Desabilita todos os Tipos de Destino da tela
            TipoDestino(iIndice).Enabled = False
        Next
        'Desabilita os componentes dos Frames Destino()
        FilialEmpresa.Enabled = False
        FilialEmpresa.Text = ""
        FilialEmpresaLabel.Enabled = False
        Fornecedor.Text = ""
        FilialForn.Clear
        Fornecedor.Enabled = False
        FilialForn.Enabled = False
        FilialLabel.Enabled = False
        FornecedorLabel.Enabled = False

    'Se Destino estiver marcado
    Else
        For iIndice = 0 To 1
            'Habilita todos os Tipos de Destino
            TipoDestino(iIndice).Enabled = True
        Next

        TipoDestino(TIPO_DESTINO_EMPRESA).Value = True
        
        FilialEmpresa.Enabled = True
        FilialEmpresaLabel.Enabled = True
        'Se nenhuma FilialEmpresa estiver selecionada
        If FilialEmpresa.ListIndex = -1 Then
            FilialEmpresa.Text = giFilialEmpresa
            FilialEmpresa_Validate (bCancel)
        End If

        Fornecedor.Enabled = True
        FilialForn.Enabled = True
        FornecedorLabel.Enabled = True
        FilialLabel.Enabled = True

    End If

    Exit Sub

Erro_Destino_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161289)

    End Select

    Exit Sub

End Sub

Private Sub Escolhido_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Escolhido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub Escolhido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub Escolhido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = Escolhido
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Gravar_Registro() As Long
'Grava o Pedido de Cotacao Avulsa

Dim lErro As Long
Dim iLinha As Integer
Dim sProduto As String
Dim iIndice As Integer
Dim objCotacao As New ClassCotacao
Dim colPedidoCotacao As Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Cotacao.Caption)) = 0 Then gError 74941
    
    'Verifica se nao existe linha preenchida no GridPRodutos
    If objGridProdutos.iLinhasExistentes = 0 Then gError 63142

    'Para cada linha preenchida do GridProdutos
    For iLinha = 1 To objGridProdutos.iLinhasExistentes

        'Verifica se a Quantidade está preenchida
        If Len(Trim(GridProdutos.TextMatrix(iLinha, iGrid_Quantidade_Col))) = 0 Then gError 63143

    Next
    
    For iLinha = 1 To objGridFornecedores.iLinhasExistentes

        'Verifica se pelo menos um fornecedor foi escohido
        If GridFornecedores.TextMatrix(iLinha, iGrid_Escolhido_Col) <> 0 Then Exit For
        
        'Se nenhum foi selecionado, erro.
        If iLinha = objGridFornecedores.iLinhasExistentes Then gError 84933
    
    Next

    
    
    'Percorre todas as linhas do GridProdutos
    For iLinha = 1 To objGridProdutos.iLinhasExistentes

        sProduto = GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col)

        For iIndice = 1 To objGridFornecedores.iLinhasExistentes

            'Verifica se o Produto existente em GridProdutos tambem faz parte do GridFornecedores
            If sProduto = GridFornecedores.TextMatrix(iIndice, iGrid_Prod_Col) Then

                'Verifica se para o produto em questao foi escolhido um fornecedor
                If Len(Trim(GridFornecedores.TextMatrix(iIndice, iGrid_Forn_Col))) = 0 Then gError 63170
            End If
        Next
    Next

    'Verifica se a checkbox Destino está marcada
    If Destino.Value = vbChecked Then

        'Verifica se o TipoDestino =FilialEmpresa está marcado
        If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then

            'Verifica se FilialEmpresa está preenchida
            If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 63171
        End If

        'Verifica se TipoDestino=Fornecedor está marcado
        If TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then

            'Verifica se Fornecedor está preenchido
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 63172
            
            'Verifica se FilialForn está preenchido
            If Len(Trim(FilialForn.Text)) = 0 Then gError 63173

        End If
        
    End If

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63176

    'Grava a cotacao
    lErro = CF("Cotacao_Grava", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63177

    Set gobjCotacao = objCotacao
    Set gcolPedidoCotacao = colPedidoCotacao

    Call Limpa_Tela_Cotacao

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 63142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_PRODUTOS_VAZIO", gErr)

        Case 63143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_COTAR_NAO_PREENCHIDA", gErr, GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col))

        Case 63170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_FORNECEDOR_ESCOLHIDO", gErr, sProduto)

        Case 63171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_PREENCHIDA", gErr)

        Case 63172
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 63173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 63176, 63177
            'Erros tratados nas rotinas chamadas

        Case 74941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_GERACAO_NAO_PREENCHIDO", gErr)
            
        Case 84933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_SELECIONADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161290)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function
Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresa_Validate

    'Se a FilialEmpresa tiver sido selecionada ==> sai da rotina
    If FilialEmpresa.ListIndex <> -1 Then Exit Sub

        'Verifica se FilialEmpresa foi preenchida
        If Len(Trim(FilialEmpresa.Text)) > 0 Then

            'Tenta selecionar a FilialEmpresa na combo FilialEmpresa
            lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 63100

            'Se nao encontra o ítem com o código informado
            If lErro = 6730 Then

                'preeenche objFilialEmpresa
                objFilialEmpresa.iCodFilial = iCodigo

                'Le a FilialEmpresa
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 63101

                'Se nao encontrou => erro
                If lErro = 27378 Then gError 63102

                If lErro = SUCESSO Then

                    'Coloca na tela o codigo e o nome da FilialEmpresa
                    FilialEmpresa.Text = objFilialEmpresa.lCodEmpresa & SEPARADOR & objFilialEmpresa.sNome

                End If

            End If

            'Se nao encontrou e nao era codigo
            If lErro = 6731 Then gError 63103

        End If

    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True

    Select Case gErr

        Case 63100, 63101
            'Erros tratados nas rotinas chamadas

        Case 63102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 63103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA1", gErr, FilialEmpresa.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161291)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objComprador As New ClassComprador
Dim lCotacao As Long
Dim objUsuario As New ClassUsuario
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load

    giFrameAtual = 1

    Set objEventoBotaoProdutos = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoProdutos = New AdmEvento 'Inserido por Wagner

    Set objGridProdutos = New AdmGrid
    Set objGridFornecedores = New AdmGrid

    'Inicializa o GridProdutos
    lErro = Inicializa_Grid_Produtos(objGridProdutos)
    If lErro <> SUCESSO Then gError 63082

    'Inicializa o GridFornecedores
    lErro = Inicializa_Grid_Fornecedores(objGridFornecedores)
    If lErro <> SUCESSO Then gError 63083

    'Carrega a listbox CondPagtos
    'lErro = Carrega_CondicaoPagamento()
    lErro = CF("Carrega_CondicaoPagamento", CondPagtos, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 63084

    'Carrega a listbox TipoProduto
    lErro = Carrega_TipoProduto()
    If lErro <> SUCESSO Then gError 63085

    'Preenche a combo de ordenacao
    Call Ordenacao_Carrega

    'Inicializa a máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 63086

    'Inicializa a máscara de Prod
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Prod)
    If lErro <> SUCESSO Then gError 63087

    'Coloca as Quantidades da tela no formato de Estoque
    Quantidade.Format = FORMATO_ESTOQUE
    QuantPedida.Format = FORMATO_ESTOQUE
    QuantRecebida.Format = FORMATO_ESTOQUE

    objComprador.sCodUsuario = gsUsuario

    'Verifica se gsUsuario é comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63088
    'Se gsUsuario nao é comprador==> erro
    If lErro = 50059 Then gError 63089

    objUsuario.sCodUsuario = objComprador.sCodUsuario

    lErro = CF("Usuario_Le", objUsuario)
    If lErro <> SUCESSO And lErro <> 36347 Then gError 63251
    If lErro = 36347 Then gError 63252

    'Coloca o Nome Reduzido do Comprador na tela
    Comprador.Caption = objUsuario.sNomeReduzido

    'Carrega a combo FilialEmpresa
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 63091

    Destino.Value = vbChecked

    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True

    'Coloca FiliaEmpresa Default na Tela
    iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126945
    
    FilialEmpresa.Text = iFilialEmpresa
    
    Call FilialEmpresa_Validate(bSGECancelDummy)

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 63082 To 63088, 126945
            'Erros tratados nas rotinas chamadas

        Case 63089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 63090, 63091, 63251
            'Erros tratados nas rotinas chamadas

        Case 63252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuario.sCodUsuario)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161292)

    End Select

    Exit Sub

End Sub

Private Sub Ordenacao_Carrega()
'preenche a combo de ordenacao e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordenação dos Bloqueios
    asOrdenacao(0) = "CotacaoProduto.Produto,CotacaoProduto.Fornecedor,CotacaoProduto.Filial"
    asOrdenacao(1) = "CotacaoProduto.Fornecedor,CotacaoProduto.Filial,CotacaoProduto.Produto"

    asOrdenacaoString(0) = "Produto"
    asOrdenacaoString(1) = "Fornecedor"

    'Carrega a Combobox Ordenacao
    For iIndice = 0 To 1

        Ordenacao.AddItem asOrdenacaoString(iIndice)
        Ordenacao.ItemData(Ordenacao.NewIndex) = iIndice

    Next

    'Seleciona a opção Produto + Fornecedor + Filial de seleção
    Ordenacao.ListIndex = 0

    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    'libera as variaveis globais
    Set objEventoProdutos = Nothing
    Set objEventoBotaoProdutos = Nothing
    Set objEventoFornecedor = Nothing

    Set objGridProdutos = Nothing
    Set objGridFornecedores = Nothing
    
    Set gobjCotacao = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161293)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialEmpresa() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o Código e o Nome de toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 63094

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 63094
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161294)

    End Select

    Exit Function

End Function
Function Fornecedores_Adiciona(objProduto As ClassProduto) As Long
'Adiciona o Fornecedor em colFornecedor

Dim lErro As Long
Dim colFilFornFilEmp As Collection
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialFornecedorEstatistica As New ClassFilialFornecedorEst
Dim colFilFornEstatistica As New Collection
Dim objFornecedor As New ClassFornecedor
Dim colFornecedor As New Collection
Dim iFilial As Integer
Dim iFilialAtual As Integer
Dim iLinha As Integer
Dim lFornecedorAnterior As Long
Dim colFornecedorProdutoFF As New Collection
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim lForn As Long
Dim lFornAtual As Long
Dim colFilialFornecedor As New Collection
Dim objFornecedorProdFF As New ClassFornecedorProdutoFF
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_Fornecedores_Adiciona

    'Le os registros de FornecedorProdutoFF
    lErro = CF("FornecedoresProdutoFF_Le", colFornecedorProdutoFF, objProduto)
    If lErro <> SUCESSO And lErro <> 63156 Then gError 63109
    
    For Each objFornecedorProdutoFF In colFornecedorProdutoFF
        
        objFornecedorProdutoFF.dtDataUltimaCotacao = DATA_NULA
        objFornecedorProdutoFF.dtDataPedido = DATA_NULA
        objFornecedorProdutoFF.dtDataReceb = DATA_NULA
        objFornecedorProdutoFF.dtDataUltimaCompra = DATA_NULA
        
        objFornecedorProdutoFF.sProduto = objProduto.sCodigo
        
        If lErro = SUCESSO Then

            'retorna os dados da ultima cotacao da FilialEmpresa(maior data de pedido de cotação) do produto/Fornecedor/FilialForn passado como parametro.
            lErro = CF("UltimaCotacao_Le_FornecedorProduto", objFornecedorProdutoFF)
            If lErro <> SUCESSO Then gError 89346
        
            'retorna os dados da ultima compra global do produto/Fornecedor/FilialForn passado como parametro.
            lErro = CF("UltimaCompra_Le_FornecedorProduto", objFornecedorProdutoFF)
            If lErro <> SUCESSO Then gError 89349
    
            'retorna os dados do ultimo item de pedido de compra com status fechado para o produto/Fornecedor/FilialForn passado como parametro.
            lErro = CF("UltimoItemPedCompraFechado_Le", objFornecedorProdutoFF)
            If lErro <> SUCESSO Then gError 89354
              
            'retorna a Quantidade em Unidade de Compra do Produto Pedido pela Filial em questão e ainda não entregue para o produto/Fornecedor/FilialForn passado como parametro.
            lErro = CF("QuantProdutoPedAbertos_Le", objFornecedorProdutoFF)
            If lErro <> SUCESSO Then gError 89368
    
            'Calcula o Tempo de Ressuprimento
            lErro = CF("TempoRessup_Calcula", objFornecedorProdutoFF)
            If lErro <> SUCESSO Then gError 89370
       
        End If

    Next
    
    'Se não encontrou registro em FornecedorProdutoFF
    If lErro = 63156 Then gError 63157

    For Each objFornecedorProdutoFF In colFornecedorProdutoFF

        lFornecedorAnterior = objFilialFornecedor.lCodFornecedor
        
        Set objFilialFornecedor = New ClassFilialFornecedor
        
        objFilialFornecedor.iCodFilial = objFornecedorProdutoFF.iFilialForn
        objFilialFornecedor.lCodFornecedor = objFornecedorProdutoFF.lFornecedor
        
        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 63215
        
        colFilialFornecedor.Add objFilialFornecedor
        
        'Se o fornecedor nao foi o ultimo lido
        If lFornecedorAnterior <> objFilialFornecedor.lCodFornecedor Then

            Set objFornecedor = New ClassFornecedor
            objFornecedor.lCodigo = objFilialFornecedor.lCodFornecedor
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 63216
            If lErro = 12729 Then gError 63274

            'Adiciona em colFornecedor
            colFornecedor.Add objFornecedor

        End If

    Next

    iFilial = objFilFornFilEmp.iCodFilial
    lForn = objFilFornFilEmp.lCodFornecedor

    For Each objFornecedorProdutoFF In colFornecedorProdutoFF

        iFilialAtual = objFornecedorProdutoFF.iFilialForn
        lFornAtual = objFornecedorProdutoFF.lFornecedor
        iLinha = objGridFornecedores.iLinhasExistentes + 1
        If iFilialAtual <> iFilial Or lForn <> lFornAtual Then

            'Preenche o Grid de Fornecedores
            lErro = GridFornecedores_Preenche(objProduto, objFornecedorProdutoFF, colFilialFornecedor, colFornecedor, iLinha)
            If lErro <> SUCESSO Then gError 63272
            iFilial = iFilialAtual
            lForn = lFornAtual

        End If

    Next

    Fornecedores_Adiciona = SUCESSO

    Exit Function

Erro_Fornecedores_Adiciona:

    Fornecedores_Adiciona = gErr

    Select Case gErr

        Case 63109, 63215, 63216, 63272
            'Erros tratados nas rotinas chamadas

        Case 63157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_FORNECEDOR", gErr, objProduto.sCodigo)

        Case 63274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

    End Select

    Exit Function

End Function

Private Function GridFornecedores_Preenche(objProduto As ClassProduto, objFornecedorProdutoFF As ClassFornecedorProdutoFF, colFilialFornecedor As Collection, colFornecedor As Collection, iLinha As Integer) As Long
'Preenche o GridFornecedores

Dim sProdutoEnxuto As String
Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_GridFornecedores_Preenche

    'Mascara o Produto
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError 63292

    Prod.PromptInclude = False
    Prod.Text = sProdutoEnxuto
    Prod.PromptInclude = True

    GridFornecedores.TextMatrix(iLinha, iGrid_Prod_Col) = Prod.Text

    GridFornecedores.TextMatrix(iLinha, iGrid_DescProduto_Col) = objProduto.sDescricao
    
    For Each objFornecedor In colFornecedor
        'Busca o Fornecedor para preencher o GridFornecedores
        
        If objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo Then
            GridFornecedores.TextMatrix(iLinha, iGrid_Forn_Col) = objFornecedor.sNomeReduzido
        End If
    Next
    
    For Each objFilialFornecedor In colFilialFornecedor
        If objFornecedorProdutoFF.iFilialForn = objFilialFornecedor.iCodFilial Then
        
            objFilialFornecedor.iCodFilial = objFornecedorProdutoFF.iFilialForn
            objFilialFornecedor.lCodFornecedor = objFornecedorProdutoFF.lFornecedor

            'Lê a Filial do Fornecedor
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 63315
            If lErro = 12929 Then gError 63317

            GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornecedor_Col) = objFornecedorProdutoFF.iFilialForn & SEPARADOR & objFilialFornecedor.sNome
            GridFornecedores.TextMatrix(iLinha, iGrid_Observacao_Col) = objFilialFornecedor.sObservacao
        
        End If
    Next
    
    
    If objFornecedorProdutoFF.dtDataUltimaCompra <> DATA_NULA Then GridFornecedores.TextMatrix(iLinha, iGrid_DataUltimaCompra_Col) = objFornecedorProdutoFF.dtDataUltimaCompra
    If objFornecedorProdutoFF.dtDataUltimaCotacao <> DATA_NULA Then GridFornecedores.TextMatrix(iLinha, iGrid_DataUltimaCotacao_Col) = objFornecedorProdutoFF.dtDataUltimaCotacao
    GridFornecedores.TextMatrix(iLinha, iGrid_UMCompra_Col) = objProduto.sSiglaUMCompra
    GridFornecedores.TextMatrix(iLinha, iGrid_SaldoTitulos_Col) = objFornecedor.dSaldoTitulos
    GridFornecedores.TextMatrix(iLinha, iGrid_UltimaCotacao_Col) = objFornecedorProdutoFF.dUltimaCotacao
    GridFornecedores.TextMatrix(iLinha, iGrid_CondicaoPagto_Col) = objFornecedorProdutoFF.sCondPagto
    
    If objFornecedorProdutoFF.dTempoRessup <> 0 Then
        GridFornecedores.TextMatrix(iLinha, iGrid_PrazoEntrega_Col) = objFornecedorProdutoFF.dTempoRessup
    End If
    
    If objFornecedor.iCondicaoPagto <> 0 Then

        objCondicaoPagto.iCodigo = objFornecedor.iCondicaoPagto

        'Lê a CondicaoPagto informada
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 62845

        GridFornecedores.TextMatrix(iLinha, iGrid_CondicaoPagto_Col) = objFornecedor.iCondicaoPagto & SEPARADOR & objCondicaoPagto.sDescReduzida

    End If

    'Preenche o TipoFrete
    If objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_FOB Then

        GridFornecedores.TextMatrix(iLinha, iGrid_TipoFrete_Col) = "FOB"

    ElseIf objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_CIF Then

        GridFornecedores.TextMatrix(iLinha, iGrid_TipoFrete_Col) = "CIF"

    End If


    GridFornecedores.TextMatrix(iLinha, iGrid_QuantPedida_Col) = Formata_Estoque(objFornecedorProdutoFF.dQuantPedida)
    GridFornecedores.TextMatrix(iLinha, iGrid_QuantRecebida_Col) = Formata_Estoque(objFornecedorProdutoFF.dQuantRecebida)
    GridFornecedores.TextMatrix(iLinha, iGrid_UltimaCotacao_Col) = Formata_Estoque(objFornecedorProdutoFF.dQuantUltimaCotacao)
    
    'Atualiza o número de Linhas existentes do GridFornecedores
    objGridFornecedores.iLinhasExistentes = objGridFornecedores.iLinhasExistentes + 1

    GridFornecedores_Preenche = SUCESSO

    Exit Function

Erro_GridFornecedores_Preenche:

    Select Case gErr

        Case 63292, 63315, 63845
            'Erros tratados nas rotinas chamadas

        Case 63317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFornecedor.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161295)

    End Select

    Exit Function

End Function

Private Sub FilialForn_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FilialForn_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialForn.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialForn.ListIndex >= 0 Then Exit Sub

    'Tenta selecionar na combo de FilialForn
    lErro = Combo_Seleciona(FilialForn, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 63136

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 63137

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 63138

        'Se nao existir
        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = sFornecedor

            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 63139

            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            gError 63140

        End If

        'Coloca na tela o código e o nome da FilialForn
        FilialForn.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 63141

    Exit Sub

Erro_FilialForn_Validate:

    Cancel = True

    Select Case gErr

        Case 63136, 63138, 63139 'Tratados nas Rotinas chamadas

        Case 63140
            'Pergunta se deseja criar nova filial para o fornecedor em questao
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela FiliaisFornecedores
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 63137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 63141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialForn.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161296)

    End Select

    Exit Sub


End Sub
Private Sub DescProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantPedida_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantPedida_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub QuantPedida_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub QuantPedida_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = QuantPedida
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub SaldoTitulos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SaldoTitulos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub SaldoTitulos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub SaldoTitulos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = SaldoTitulos
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub CondicaoPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub CondicaoPagto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub CondicaoPagto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = CondicaoPagto
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantRecebida_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantRecebida_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub QuantRecebida_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub QuantRecebida_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = QuantRecebida
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMCompra_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMCompra_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub UMCompra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub UMCompra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = UMCompra
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrazoEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrazoEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub PrazoEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub PrazoEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = PrazoEntrega
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataUltimaCompra_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataUltimaCompra_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub DataUltimaCompra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub DataUltimaCompra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = DataUltimaCompra
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataUltimaCotacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub DataUltimaCotacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub DataUltimaCotacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = DataUltimaCotacao
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoFrete_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub TipoFrete_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub TipoFrete_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = TipoFrete
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub UltimaCotacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UltimaCotacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub UltimaCotacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub UltimaCotacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = UltimaCotacao
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialFornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornecedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub FilialFornecedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub FilialFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = FilialFornecedor
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Forn_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Forn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub Forn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub Forn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = Forn
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se Fornecedor foi alterado
    If giFornecedorAlterado = 0 Then Exit Sub

    'Verifica se o Fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Le o Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 63129

        'Le as Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then gError 63130

        'Preenche a combo FilialForn
        Call CF("Filial_Preenche", FilialForn, colCodigoNome)

        'Seleciona a filial na combo de FilialForn
        Call CF("Filial_Seleciona", FilialForn, iCodFilial)

    End If

    'Se o Fornecedor nao estiver preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a combo FilialForn
        FilialForn.Clear

    End If

    giFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 63129, 63130
            'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161297)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentes As Integer
Dim sProduto As String
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_GridProdutos_KeyDown


    'Guarda o número de linhas existentes
    iLinhasExistentes = objGridProdutos.iLinhasExistentes

    'Guarda o código do Produto da linha atual
    sProduto = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col)

    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos)

    'Verifica se o número de linhas existentes diminuiu
    If iLinhasExistentes > objGridProdutos.iLinhasExistentes Then

        For iIndice = 1 To objGridFornecedores.iLinhasExistentes

            'Verifica se o produto que foi excluido aparece no GridFornecedores
            If (sProduto = GridFornecedores.TextMatrix(iIndice, iGrid_Prod_Col)) Then

                'Exclui do GridFornecedores as linhas correspondentes ao Produto excluido
                Call Grid_Exclui_Linha(objGridFornecedores, iIndice)
                iIndice = iIndice - 1

            End If

        Next

    End If

    Exit Sub

Erro_GridProdutos_KeyDown:

    Select Case gErr

        Case 63115
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161298)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSiglas As New Collection
Dim objClasseUM As New ClassClasseUM
Dim objUM As New ClassUnidadeDeMedida
Dim sUM As String
Dim iIndice As Integer
Dim sProduto As String
Dim lFornecedor As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim iCodFilial As Integer
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa controle da coluna em questão
    Select Case objControl.Name

        'Produto
        Case Produto.Name
            'Verifca se o Produto está preenchido
            If Len(Trim(GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If
        
        'Quantidade ou Fornecedor
        Case Quantidade.Name, Fornecedor.Name
            'Verifica se o Produto está preenchido
            If Len(Trim(GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col))) = 0 Then
                objControl.Enabled = False
                Exit Sub
            Else
                objControl.Enabled = True
            End If

        'FilialForn
        Case FilialForn.Name

            'Verifica se o Produto está preenchido
            If Len(Trim(GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_Produto_Col))) = 0 Then
                objControl.Enabled = False
                Exit Sub
            Else
                objControl.Enabled = True
            End If

            'Limpa a combo FiliForn
            FilialForn.Clear

            'Verifica se FilialForn está preenchido nesta mesma linha
            If Len(Trim(GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_FilialFornecedor_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            sProduto = GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_Prod_Col)

            objFornecedor.sNomeReduzido = GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_Forn_Col)

            'Le o Fornecedor a partir do NomeReduzido passado como parametro
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 And lErro <> 6684 Then gError 63249

            lFornecedor = objFornecedor.lCodigo

            'Le código e nome das Filiais
            lErro = CF("FornecedorProdutoFF_Le_FilForn", sProduto, lFornecedor, colCodigoNome)
            If lErro <> SUCESSO And lErro <> 63219 Then gError 63161

            'Se não encontrou nenhuma filial ==>erro
            If lErro <> 63219 Then gError 63162

            'Preenche a combo FilialForn
            Call CF("Filial_Preenche", FilialForn, colCodigoNome)

            'Seleciona FilialForn na combobox
            Call CF("Filial_Seleciona", FilialForn, iCodFilial)

        Case PrazoEntrega.Name
        
            If Len(Trim(GridFornecedores.TextMatrix(iLinha, iGrid_Prod_Col))) > 0 Then
                PrazoEntrega.Enabled = True
            Else
                PrazoEntrega.Enabled = False
            End If
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 63116, 63158, 63160, 63161, 63249
            'Erros tratados nas rotinas chamadas

        Case 63159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 63162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR", gErr, objFornecedor.sNomeReduzido, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161299)

    End Select

    Exit Sub

End Sub


Private Sub GridFornecedores_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFornecedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFornecedores, iAlterado)
    End If

End Sub

Private Sub GridFornecedores_GotFocus()
    Call Grid_Recebe_Foco(objGridFornecedores)
End Sub

Private Sub GridFornecedores_EnterCell()
    Call Grid_Entrada_Celula(objGridFornecedores, iAlterado)
End Sub

Private Sub GridFornecedores_LeaveCell()
    Call Saida_Celula(objGridFornecedores)
End Sub

Private Sub GridFornecedores_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridFornecedores)
End Sub

Private Sub GridFornecedores_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFornecedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFornecedores, iAlterado)
    End If

End Sub

Private Sub GridFornecedores_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridFornecedores)
End Sub

Private Sub GridFornecedores_RowColChange()
    Call Grid_RowColChange(objGridFornecedores)
End Sub

Private Sub GridFornecedores_Scroll()
    Call Grid_Scroll(objGridFornecedores)
End Sub

Private Sub GridProdutos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
    End If

End Sub

Private Sub GridProdutos_GotFocus()
    Call Grid_Recebe_Foco(objGridProdutos)
End Sub

Private Sub GridProdutos_EnterCell()
    Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
End Sub

Private Sub GridProdutos_LeaveCell()
    Call Saida_Celula(objGridProdutos)
End Sub


Private Sub GridProdutos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
    End If

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

Private Function Saida_Celula_GridProdutos(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridProdutos que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProdutos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 63153

        'UnidadeMed
        Case iGrid_UM_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 63154

        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 63155

        Case iGrid_DescricaoProd_Col
            lErro = Saida_Celula_Descricao(objGridInt)
            If lErro <> SUCESSO Then gError 63254

    End Select

    Saida_Celula_GridProdutos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridProdutos:

    Saida_Celula_GridProdutos = gErr

    Select Case gErr

        Case 63153 To 63155

        Case 63254

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161300)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridFornecedores(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridFornecedores que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridFornecedores

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'PrazoEntrega
        Case iGrid_PrazoEntrega_Col
            
            Set objGridInt.objControle = PrazoEntrega

            'Verifica se PrazoEntrega esta preeenchida
            If Len(Trim(PrazoEntrega.ClipText)) > 0 Then

                'Critica PrazoEntrega
                lErro = Valor_Positivo_Critica(PrazoEntrega.Text)
                If lErro <> SUCESSO Then gError 72518

            
            End If

    End Select

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 72519

    Saida_Celula_GridFornecedores = SUCESSO

    Exit Function

Erro_Saida_Celula_GridFornecedores:

    Saida_Celula_GridFornecedores = gErr

    Select Case gErr

        Case 72518, 72519
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161301)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'Faz a saida de celula de Descricao

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = DescricaoProd

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63255

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 63255
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161302)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridProdutos
            Case GridProdutos.Name

                lErro = Saida_Celula_GridProdutos(objGridInt)
                If lErro <> SUCESSO Then gError 63150

            Case GridFornecedores.Name
                lErro = Saida_Celula_GridFornecedores(objGridInt)
                If lErro <> SUCESSO Then gError 72517
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 63152

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 63150 To 63152
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161303)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a saida de célula de Produto

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim bProdutoPresente As Boolean
Dim iIndice As Integer
Dim iLinha As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim sProd As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se o Produto está preenchido
    If Len(Trim(Produto.Text)) > 0 Then

        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 63256

        objProduto.sCodigo = sProdutoFormatado

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Verifica se o Produto existe e pode ser Comprado
            lErro = CF("Produto_Critica_Compra", Produto.Text, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25605 Then gError 63163

            'Se o produto não existir ==> erro
            If lErro = 25605 Then gError 63164

            sProd = Produto.Text

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 63165

            'Preenche o Produto com o ProdutoEnxuto
            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            'Verifica se já existe o Produto em outra linha do Grid
            For iIndice = 1 To objGridProdutos.iLinhasExistentes
                If iIndice <> GridProdutos.Row Then
                    If GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 63166
                End If
            Next

            'Verifica se a linha atual do Grid é maior que o número de linhas existentes
            If GridProdutos.Row > objGridProdutos.iLinhasExistentes Then

                'Adiciona o Fornecedor lido em colFornecedor
                lErro = Fornecedores_Adiciona(objProduto)
                If lErro <> SUCESSO Then gError 63168
                
                'Incrementa o número de linhas existentes do GridProdutos
                objGridProdutos.iLinhasExistentes = objGridProdutos.iLinhasExistentes + 1

            End If
            
            'Preenche a UM de Compra e a Descricao do Produto
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 63167

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63169

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 63163, 63165, 63167, 63168, 63256
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63164
            'Pergunta se deseja criar produto
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de Produtos
                Call Chama_Tela("Produto", objProduto)
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If


        Case 63166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_GRID_PRODUTOS", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63169
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a saida de celula de Quantidade

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Quantidade

     Set objGridInt.objControle = Quantidade

    'Verifica se a quantidade esta preeenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        'Critica a quantidade
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 63121

        dQuantidade = StrParaDbl(Quantidade.Text)

        'Coloca a quantidade com o formato de estoque da tela
         Quantidade.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63122

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 63121, 63122
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161304)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a saida de celula de UM

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UM

    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UM_Col) = UM.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63123

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 63123
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161305)

    End Select

    Exit Function

End Function
Private Sub Col_Detalhe_Click()
'Chama a tela de Produto para o produto da linha GridProdutos escolhida

Dim objProduto As New ClassProduto
Dim lErro As Long

On Error GoTo Erro_Col_Detalhe_Click

    'Verifica se o produto está preenchido
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) > 0 Then

        'Passa o codigo do Produto para objProduto
        objProduto.sCodigo = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col)

        'Chama a tela de Produto
        Call Chama_Tela("Produto", objProduto)

    End If

    Exit Sub

Erro_Col_Detalhe_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161306)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO Then gError 63108

    For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na ListBox CondPagtos
        TipoProduto.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        TipoProduto.ItemData(TipoProduto.NewIndex) = objCod_DescReduzida.iCodigo

    Next


    'Marca todos os TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = True
    Next

    Carrega_TipoProduto = SUCESSO

    Exit Function

Erro_Carrega_TipoProduto:

    Carrega_TipoProduto = gErr

    Select Case gErr

        Case 63108
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161307)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_CondicaoPagamento() As Long
''Carrega Listbox CondPagtos com as condicoes de pagamento usadas em contas a pagar (EmPagamento=1)
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As New AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Le todos os Codigos e DescReduzida da tabela CondicoesPagto com a condicao EmPagamento = 1 e coloca na colecao colCod_DescReduzida
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 63092
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Verifica se a CondPagto é diferente de "À Vista"
'        If objCod_DescReduzida.iCodigo <> CONDPAGTO_VISTA Then
'
'            'Adiciona novo item na ListBox CondPagtos
'            CondPagtos.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'            CondPagtos.ItemData(CondPagtos.NewIndex) = objCod_DescReduzida.iCodigo
'
'        End If
'
'    Next
'
'    Carrega_CondicaoPagamento = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CondicaoPagamento:
'
'    Carrega_CondicaoPagamento = gErr
'
'    Select Case gErr
'
'        Case 63092
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161308)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Inicializa_Grid_Produtos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("A Cotar")

    'campos de edição do grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProd.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quantidade.Name)

    'indica onde estao situadas as colunas do grid
    'iGrid_Detalhe_Col = 1
    iGrid_Produto_Col = 1
    iGrid_DescricaoProd_Col = 2
    iGrid_UM_Col = 3
    iGrid_Quantidade_Col = 4

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PRODUTOS_COTACAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 14

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridProdutos.Width = 7000
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos:

    Inicializa_Grid_Produtos = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161309)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Fornecedores(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Fornecedores

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Fornecedores

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Selecionado")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Última Cotação")
    objGridInt.colColuna.Add ("Frete")
    objGridInt.colColuna.Add ("Data Cotação")
    objGridInt.colColuna.Add ("Última Compra")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Unidade Med.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    objGridInt.colColuna.Add ("Quant. Recebida")
    objGridInt.colColuna.Add ("Condição Pagto")
    objGridInt.colColuna.Add ("Saldo Tit. a Pagar")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (Escolhido.Name)
    objGridInt.colCampo.Add (Prod.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (Forn.Name)
    objGridInt.colCampo.Add (FilialFornecedor.Name)
    objGridInt.colCampo.Add (UltimaCotacao.Name)
    objGridInt.colCampo.Add (TipoFrete.Name)
    objGridInt.colCampo.Add (DataUltimaCotacao.Name)
    objGridInt.colCampo.Add (DataUltimaCompra.Name)
    objGridInt.colCampo.Add (PrazoEntrega.Name)
    objGridInt.colCampo.Add (UMCompra.Name)
    objGridInt.colCampo.Add (QuantPedida.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (CondicaoPagto.Name)
    objGridInt.colCampo.Add (SaldoTitulos.Name)
    objGridInt.colCampo.Add (Observacao.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Escolhido_Col = 1
    iGrid_Prod_Col = 2
    iGrid_DescProduto_Col = 3
    iGrid_Forn_Col = 4
    iGrid_FilialFornecedor_Col = 5
    iGrid_UltimaCotacao_Col = 6
    iGrid_TipoFrete_Col = 7
    iGrid_DataUltimaCotacao_Col = 8
    iGrid_DataUltimaCompra_Col = 9
    iGrid_PrazoEntrega_Col = 10
    iGrid_UMCompra_Col = 11
    iGrid_QuantPedida_Col = 12
    iGrid_QuantRecebida_Col = 13
    iGrid_CondicaoPagto_Col = 14
    iGrid_SaldoTitulos_Col = 15
    iGrid_Observacao_Col = 16

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridFornecedores

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_FORNECEDORES_COTACAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR


    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Fornecedores = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Fornecedores:

    Inicializa_Grid_Fornecedores = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161310)

    End Select

    Exit Function

End Function

Function Trata_Parametros()

    Trata_Parametros = SUCESSO

End Function

Private Sub Fornecedor_Change()

    giFornecedorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FornecedorLabel_Click()
'Chama a tela FornecedorLista

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Coloca o Fornecedor que está na tela no objFornecedor
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub



Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o nome reduzido do Fornecedor na tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Fornecedor_Validate (bCancel)

    Me.Show

End Sub

Private Sub objEventoBotaoProdutos_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim bCancel As Boolean
Dim sProdutoEnxuto As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_objEventoBotaoProdutos_evSelecao

    Set objProduto = obj1

    'Verifica se existe alguma linha do GridProdutos selecionada
    If GridProdutos.Row <> 0 Then

        'Verifica se o produto desta linha não está preenchido
        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) = 0 Then

            'Retorna o produto enxuto
            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 63096

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            'Verifica se já existe o Produto em outra linha do Grid
            For iIndice = 1 To objGridProdutos.iLinhasExistentes
                If iIndice <> GridProdutos.Row Then
                    If GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 74939
                End If

            Next

            'Verifica se a linha atual do grid é menor que o numero de linhas existentes
            If GridProdutos.Row > objGridProdutos.iLinhasExistentes Then

                'Preenche o GridFornecedores
                lErro = Fornecedores_Adiciona(objProduto)
                If lErro <> SUCESSO Then gError 63098

                'Aumenta o número de linhas existentes do GridProdutos
                objGridProdutos.iLinhasExistentes = objGridProdutos.iLinhasExistentes + 1

            End If

            'Alterado por Wagner
            If Not (Me.ActiveControl Is Produto) Then
                'Preenche o produto no GridProdutos
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col) = Produto.Text
    
                'Preenche a Unidade de Medida de Compra e a Descricao do Produto no GridProdutos
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 63097
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoBotaoProdutos_evSelecao:

    Select Case gErr

        Case 63096, 63097, 63098
            'Erro tratado na rotina chamada

        Case 74939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_GRID_PRODUTOS", gErr, Produto.Text, iIndice)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161311)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long
'Preenche a Unidade de Medida de Compra e a Descricao do Produto

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_ProdutoLinha_Preenche

    'Preenche unidade de medida e descricao do produto
    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_UM_Col) = objProduto.sSiglaUMCompra
    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_DescricaoProd_Col) = objProduto.sDescricao

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161312)

    End Select

    Exit Function

End Function

Private Function Move_Cotacao_Memoria(objCotacao As ClassCotacao) As Long
'Recolhe os dados da tela que não pertencem aos grids para a memória

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objFornecedor As New ClassFornecedor
Dim objUsuario As New ClassUsuario
Dim iIndice As Integer
Dim iCodigo  As Integer

On Error GoTo Erro_Move_Cotacao_Memoria

    objUsuario.sNomeReduzido = Comprador.Caption
    'Lê o usuario a partir do nome reduzido
    lErro = CF("Usuario_Le_NomeRed", objUsuario)
    If lErro <> SUCESSO And lErro <> 57269 Then gError 63276
    If lErro = 57269 Then gError 63277

    objComprador.sCodUsuario = objUsuario.sCodUsuario

    'Lê o comprador a partir do codUsuario
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63226

    'Se não encontrou o comprador==>erro
    If lErro = 50059 Then gError 63227

    objCotacao.iComprador = objComprador.iCodigo
    objCotacao.iFilialEmpresa = giFilialEmpresa
    objCotacao.sDescricao = Descricao.Text
    objCotacao.dtData = gdtDataAtual
    objCotacao.lCodigo = StrParaLong(Cotacao.Caption)
    iCodigo = COD_A_VISTA
    objCotacao.colCondPagtos.Add (iCodigo)
    
    If Len(Trim(CondPagto.Text)) > 0 Then
        iCodigo = Codigo_Extrai(CondPagto.Text)
        objCotacao.colCondPagtos.Add (iCodigo)
    End If

    'Frame Tipo
    If Destino.Value = vbChecked Then

        If TipoDestino(TIPO_DESTINO_EMPRESA) = True Then
            objCotacao.iTipoDestino = TIPO_DESTINO_EMPRESA
            objCotacao.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)

        ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR) = True Then
            objCotacao.iTipoDestino = TIPO_DESTINO_FORNECEDOR
            objCotacao.iFilialDestino = Codigo_Extrai(FilialForn.Text)

            objFornecedor.sNomeReduzido = Fornecedor.Text

            'Le o Fornecedor
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 63228
            If lErro = 6681 Then gError 63229
            objCotacao.lFornCliDestino = objFornecedor.lCodigo

        End If
    Else
        objCotacao.iTipoDestino = TIPO_DESTINO_AUSENTE
        objCotacao.iFilialDestino = 0
        objCotacao.lFornCliDestino = 0
    End If

    Move_Cotacao_Memoria = SUCESSO

    Exit Function

Erro_Move_Cotacao_Memoria:

    Move_Cotacao_Memoria = gErr

    Select Case gErr


        Case 63226, 63228, 63276
            'Erros tratados nas rotinas chamadas

        Case 63227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 63229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case 63277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_INEXISTENTE", gErr, objUsuario.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161313)

    End Select

    Exit Function

End Function

Function Move_GridFornecedores_Memoria(objCotacao As ClassCotacao, colPedidoCotacao As Collection) As Long
'Recolhe os dados do GridFornecedores

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer
Dim iLinha2 As Integer
Dim objPedidoCotacao As ClassPedidoCotacao
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemPedCotacao As ClassItemPedCotacao
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_GridFornecedores_Memoria

    'Inicializa a coleção de Pedidos de Cotação
    Set colPedidoCotacao = New Collection
    
    'Para cada linha do Grid de Fornecedores
    For iLinha = 1 To objGridFornecedores.iLinhasExistentes
                
        'Se a linha foi marcada
        If GridFornecedores.TextMatrix(iLinha, iGrid_Escolhido_Col) = GRID_CHECKBOX_ATIVO Then
        
            'Lê o Fornecedor do Grid de Fornecedores
            objFornecedor.sNomeReduzido = GridFornecedores.TextMatrix(iLinha, iGrid_Forn_Col)
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 63626
            
            'Se não encontrou o Fornecedor, erro
            If lErro = 6681 Then gError 70513
            
            'Verifica se já existe um Pedido de Cotação para o Fornecedor e Filial da linha do Grid de Fornecedores em questão
            For iIndice = 1 To colPedidoCotacao.Count
                        
                'Se encontrou um Pedido de cotação com o mesmo Fornecedor e Filial
                If colPedidoCotacao(iIndice).lFornecedor = objFornecedor.lCodigo And colPedidoCotacao(iIndice).iFilial = Codigo_Extrai(GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornecedor_Col)) Then
                    Exit For
                End If
            
            Next
                
            'Se não encontrou o Pedido de Cotação
            If iIndice > colPedidoCotacao.Count Then
            
                'Cria um novo Pedido de Cotação
                Set objPedidoCotacao = New ClassPedidoCotacao
                
                objPedidoCotacao.lFornecedor = objFornecedor.lCodigo
                objPedidoCotacao.iFilial = Codigo_Extrai(GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornecedor_Col))
                                                                                                
                'Tipo de Frete
                If GridFornecedores.TextMatrix(iLinha, iGrid_TipoFrete_Col) = "CIF" Then
                    objPedidoCotacao.iTipoFrete = TIPO_CIF
                ElseIf GridFornecedores.TextMatrix(iLinha, iGrid_TipoFrete_Col) = "FOB" Then
                    objPedidoCotacao.iTipoFrete = TIPO_FOB
                End If
                    
                'Se existe uma Condicao de Pagamento à Prazo
                If objCotacao.colCondPagtos.Count > 1 Then
                
                    objPedidoCotacao.iCondPagtoPrazo = objCotacao.colCondPagtos.Item(2)
                    
                End If
                objPedidoCotacao.iStatus = STATUS_GERADO_NAO_ATUALIZADO
                objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
                objPedidoCotacao.dtData = gdtDataHoje
                objPedidoCotacao.dtDataEmissao = gdtDataHoje
                objPedidoCotacao.dtDataValidade = DATA_NULA
                                              
                'Procura no Grid de Fornecedores as linhas que possuem o mesmo Fornecedor e Filial do Pedido de Cotação que acaba de ser criado
                For iLinha2 = 1 To objGridFornecedores.iLinhasExistentes
                    
                    'Se encontrou e a linha está marcada
                    If GridFornecedores.TextMatrix(iLinha2, iGrid_Forn_Col) = objFornecedor.sNomeReduzido And Codigo_Extrai(GridFornecedores.TextMatrix(iLinha2, iGrid_FilialFornecedor_Col)) = objPedidoCotacao.iFilial And GridFornecedores.TextMatrix(iLinha2, iGrid_Escolhido_Col) = GRID_CHECKBOX_ATIVO Then
                
                        'Formata o Produto da linha
                        lErro = CF("Produto_Formata", GridFornecedores.TextMatrix(iLinha2, iGrid_Prod_Col), sProdutoFormatado, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 68340
                
                        'Lê o Fornecedor do Grid de Fornecedores
                        objFornecedor.sNomeReduzido = GridFornecedores.TextMatrix(iLinha2, iGrid_Forn_Col)
                        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                        If lErro <> SUCESSO And lErro <> 6681 Then gError 68339
                        
                        'Se não encontrou o Fornecedor, erro
                        If lErro = 6681 Then gError 70514
                
                        'Procura o Produto na coleção de Cotação Produto
                        For Each objCotacaoProduto In objCotacao.colCotacaoProduto
                                                        
                            'Se encontrou o mesmo Produto, Fornecedor, Filial (ou só Produto no caso de Fornecedor não EXCLUSIVO)
                            If sProdutoFormatado = objCotacaoProduto.sProduto Then
                                                                
                                'Cria Item de Pedido de Cotação
                                Set objItemPedCotacao = New ClassItemPedCotacao
                                
                                objItemPedCotacao.lCotacaoProduto = objCotacaoProduto.lNumIntDoc
                                objItemPedCotacao.sProduto = sProdutoFormatado
                            
                                'Adiciona o Item na coleção de Pedido de Cotação
                                objPedidoCotacao.colItens.Add objItemPedCotacao
                                
                                Exit For
                            
                            End If
                            
                        Next
                        
                    End If
                    
                Next
                
                'Adiciona na coleção de Pedidos de cotação
                colPedidoCotacao.Add objPedidoCotacao
            End If
        End If
    Next
                                    
    Move_GridFornecedores_Memoria = SUCESSO

    Exit Function

Erro_Move_GridFornecedores_Memoria:

    Move_GridFornecedores_Memoria = gErr

    Select Case gErr

        Case 63626, 68339, 68340
                
        Case 70513, 70514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161314)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objCotacao As ClassCotacao, colPedidoCotacao As Collection) As Long
'Recolhe os dados da tela

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    Set colPedidoCotacao = New Collection

    'Move os dados que não pertencem aos grids da tela para objCotacao
    lErro = Move_Cotacao_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 63144

    'Recolhe os dados do GridProdutos
    lErro = Move_GridProdutos_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 63145

    'Recolhe os dados do GridFornecedores
    lErro = Move_GridFornecedores_Memoria(objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63146

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 63144 To 63146
            'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161315)

    End Select

    Exit Function

End Function

Function Move_GridProdutos_Memoria(objCotacao As ClassCotacao) As Long
'Recolhe os dados do GridProdutos e guarda-os em colCotacaoProduto

Dim lErro As Long
Dim objCotProduto As ClassCotacaoProduto
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objFornecedor As New ClassFornecedor
Dim lNumIntProvisorio As Long

On Error GoTo Erro_Move_GridProdutos_Memoria

    lNumIntProvisorio = 1
    
    'Para cada linha do Grid de Produtos
    For iIndice = 1 To objGridProdutos.iLinhasExistentes

        Set objCotProduto = New ClassCotacaoProduto
        
        'Preenche objCotProduto com os dados do GridProdutos
        objCotProduto.dQuantidade = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        'Coloca o produto no formato do BD
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 63627

        objCotProduto.sProduto = sProdutoFormatado
        objCotProduto.sUM = GridProdutos.TextMatrix(iIndice, iGrid_UM_Col)

        'Coloca número provisório em NumIntDoc para linkar com ItensPedCOtacao de PedidosCotacao
        objCotProduto.lNumIntDoc = lNumIntProvisorio
                    
        'Adiciona em colCotacaoProduto
        objCotacao.colCotacaoProduto.Add objCotProduto
        
        'Incrementa o número provisório
        lNumIntProvisorio = lNumIntProvisorio + 1

    Next

    Move_GridProdutos_Memoria = SUCESSO

    Exit Function

Erro_Move_GridProdutos_Memoria:

    Move_GridProdutos_Memoria = gErr

    Select Case gErr

        Case 63627
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161316)

    End Select

    Exit Function

End Function
    
Private Function GridFornecedores_Devolve(colItensGridFornecedores As Collection) As Long
'Devolve os elementos ordenados para o GridFornecedores

Dim lErro As Long
Dim objItemGridFornecedores As New ClassItemGridFornecedores
Dim iLinha As Long

On Error GoTo Erro_GridFornecedores_Devolve

    Call Grid_Limpa(objGridFornecedores)

    For Each objItemGridFornecedores In colItensGridFornecedores

        iLinha = iLinha + 1

        'Preenche o GridFornecedores
        GridFornecedores.TextMatrix(iLinha, iGrid_CondicaoPagto_Col) = objItemGridFornecedores.sCondicaoPagto
        GridFornecedores.TextMatrix(iLinha, iGrid_DataUltimaCompra_Col) = objItemGridFornecedores.sDataUltimaCompra
        GridFornecedores.TextMatrix(iLinha, iGrid_DataUltimaCotacao_Col) = objItemGridFornecedores.sDataUltimaCotacao
        GridFornecedores.TextMatrix(iLinha, iGrid_DescProduto_Col) = objItemGridFornecedores.sDescProduto
        GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornecedor_Col) = objItemGridFornecedores.sFilialForn
        GridFornecedores.TextMatrix(iLinha, iGrid_Forn_Col) = objItemGridFornecedores.sFornecedor
        GridFornecedores.TextMatrix(iLinha, iGrid_Observacao_Col) = objItemGridFornecedores.sObservacao
        GridFornecedores.TextMatrix(iLinha, iGrid_PrazoEntrega_Col) = objItemGridFornecedores.sPrazoEntrega
        GridFornecedores.TextMatrix(iLinha, iGrid_Prod_Col) = objItemGridFornecedores.sProduto
        GridFornecedores.TextMatrix(iLinha, iGrid_QuantPedida_Col) = objItemGridFornecedores.sQuantPedida
        GridFornecedores.TextMatrix(iLinha, iGrid_QuantRecebida_Col) = objItemGridFornecedores.sQuantRecebida
        GridFornecedores.TextMatrix(iLinha, iGrid_SaldoTitulos_Col) = objItemGridFornecedores.sSaldoTitulos
        GridFornecedores.TextMatrix(iLinha, iGrid_TipoFrete_Col) = objItemGridFornecedores.sTipoFrete
        GridFornecedores.TextMatrix(iLinha, iGrid_UltimaCotacao_Col) = objItemGridFornecedores.sUltimaCotacao
        GridFornecedores.TextMatrix(iLinha, iGrid_UMCompra_Col) = objItemGridFornecedores.sUMCompra

        GridFornecedores.TextMatrix(iLinha, iGrid_Escolhido_Col) = objItemGridFornecedores.iSelecionado
        
        objGridFornecedores.iLinhasExistentes = iLinha

    Next

    Call Grid_Refresh_Checkbox(objGridFornecedores)
    
    GridFornecedores_Devolve = SUCESSO

    Exit Function

Erro_GridFornecedores_Devolve:

    GridFornecedores_Devolve = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161317)

    End Select

    Exit Function

End Function


Private Sub Ordenacao_Click()

Dim lErro As Long
Dim colItensGridFornecedores As Collection
Dim colItensGridFornecedoresSaida As New Collection
Dim colCampos As New Collection

On Error GoTo Erro_Ordenacao_Click

    If gsOrdenacao = "" Then Exit Sub

    'Verifica se Ordenacao da tela é diferente de gsOrdenacao
    If Ordenacao.Text <> gsOrdenacao Then

        'Recolhe os itens do GridFornecedores
        lErro = GridFornecedores_Recolhe(colItensGridFornecedores)
        If lErro <> SUCESSO Then gError 63225

        Call Monta_Colecao_Campos_Fornecedor(colCampos, Ordenacao.ListIndex)

        lErro = Ordena_Colecao(colItensGridFornecedores, colItensGridFornecedoresSaida, colCampos)
        If lErro <> SUCESSO Then gError 63248

        'Devolve os elementos ordenados para o  GridFornecedores
        lErro = GridFornecedores_Devolve(colItensGridFornecedoresSaida)
        If lErro <> SUCESSO Then gError 63239

        gsOrdenacao = Ordenacao.Text

    End If

    Exit Sub

Erro_Ordenacao_Click:

    Select Case gErr

        Case 63225, 63239, 63248

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161318)

    End Select

    Exit Sub

End Sub

Sub Monta_Colecao_Campos_Fornecedor(colCampos As Collection, iOrdenacao As Integer)

    Select Case iOrdenacao

        Case 0
            colCampos.Add "sProduto"
            colCampos.Add "sFornecedor"
            colCampos.Add "sFilialForn"

        Case 1
            colCampos.Add "sFornecedor"
            colCampos.Add "sFilialForn"
            colCampos.Add "sProduto"

    End Select

End Sub

Private Function GridFornecedores_Recolhe(colItensGridFornecedores As Collection) As Long
'Recolhe os itens do GridFornecedores e adiciona em colItensGridFornecedores

Dim objItemGridFornecedores As New ClassItemGridFornecedores
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_GridFornecedores_Recolhe

    Set colItensGridFornecedores = New Collection

    'Percorre todas as linhas do GridFornecedores
    For iIndice = 1 To objGridFornecedores.iLinhasExistentes

        Set objItemGridFornecedores = New ClassItemGridFornecedores

        objItemGridFornecedores.iSelecionado = GridFornecedores.TextMatrix(iIndice, iGrid_Escolhido_Col)
        
        objItemGridFornecedores.sProduto = GridFornecedores.TextMatrix(iIndice, iGrid_Prod_Col)
        objItemGridFornecedores.sDescProduto = GridFornecedores.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemGridFornecedores.sFornecedor = GridFornecedores.TextMatrix(iIndice, iGrid_Forn_Col)
        objItemGridFornecedores.sFilialForn = GridFornecedores.TextMatrix(iIndice, iGrid_FilialFornecedor_Col)
        objItemGridFornecedores.sTipoFrete = GridFornecedores.TextMatrix(iIndice, iGrid_TipoFrete_Col)
        objItemGridFornecedores.sUltimaCotacao = GridFornecedores.TextMatrix(iIndice, iGrid_UltimaCotacao_Col)
        objItemGridFornecedores.sDataUltimaCompra = GridFornecedores.TextMatrix(iIndice, iGrid_DataUltimaCompra_Col)
        objItemGridFornecedores.sDataUltimaCotacao = GridFornecedores.TextMatrix(iIndice, iGrid_DataUltimaCotacao_Col)
        objItemGridFornecedores.sPrazoEntrega = GridFornecedores.TextMatrix(iIndice, iGrid_PrazoEntrega_Col)
        objItemGridFornecedores.sUMCompra = GridFornecedores.TextMatrix(iIndice, iGrid_UMCompra_Col)
        objItemGridFornecedores.sQuantPedida = GridFornecedores.TextMatrix(iIndice, iGrid_QuantPedida_Col)
        objItemGridFornecedores.sQuantRecebida = GridFornecedores.TextMatrix(iIndice, iGrid_QuantRecebida_Col)
        objItemGridFornecedores.sCondicaoPagto = GridFornecedores.TextMatrix(iIndice, iGrid_CondicaoPagto_Col)
        objItemGridFornecedores.sSaldoTitulos = GridFornecedores.TextMatrix(iIndice, iGrid_SaldoTitulos_Col)
        objItemGridFornecedores.sObservacao = GridFornecedores.TextMatrix(iIndice, iGrid_Observacao_Col)

        'Adiciona em colItensGridFornecedores
        colItensGridFornecedores.Add objItemGridFornecedores

    Next

    GridFornecedores_Recolhe = SUCESSO

    Exit Function

Erro_GridFornecedores_Recolhe:

    GridFornecedores_Recolhe = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161319)

    End Select

    Exit Function

End Function

Private Sub Ordenacao_GotFocus()

    'Guarda em gsOrdenacao a Ordenacao atual da tela
    gsOrdenacao = Ordenacao.Text

End Sub


Private Sub Prod_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Prod_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub Prod_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub Prod_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = Prod
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = DescProduto
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

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

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(giFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub


'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Pedidos de Cotação Avulsos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoPedCotacaoAvulsa"

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

Private Sub TipoDestino_Click(Index As Integer)

    'Torna invisivel o FrameDestino com índice igual a iFrameDestinoAtual
    FrameDestino(giFrameDestinoAtual).Visible = False

    'Torna visível o FrameDestino com índice igual a Index
    FrameDestino(Index).Visible = True

    'Armazena novo valor de giFrameDestinoAtual
    giFrameDestinoAtual = Index

End Sub

Private Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Forn Then
            Call BotaoFornecedores_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        End If
    End If

End Sub

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

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub


Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub FilialLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialLabel, Source, X, Y)
End Sub

Private Sub FilialLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialLabel, Button, Shift, X, Y)
End Sub
Private Sub FilialEmpresaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaLabel, Source, X, Y)
End Sub

Private Sub FilialEmpresaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaLabel, Button, Shift, X, Y)
End Sub

Private Sub Cotacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cotacao, Source, X, Y)
End Sub

Private Sub Cotacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cotacao, Button, Shift, X, Y)
End Sub

Private Sub Label54_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label54, Source, X, Y)
End Sub

Private Sub Label54_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label54, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub


Private Sub Label45_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label45, Source, X, Y)
End Sub

Private Sub Label45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label45, Button, Shift, X, Y)
End Sub

Private Sub Label55_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label55, Source, X, Y)
End Sub

Private Sub Label55_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label55, Button, Shift, X, Y)
End Sub

'######################################################
'Inserido por Wagner
Private Sub objEventoProdutos_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim bProdutoPresente As Boolean
Dim iIndice As Integer
Dim iLinha As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim sProd As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se o Produto está preenchido
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 132516

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Verifica se o Produto existe e pode ser Comprado
            lErro = CF("Produto_Critica_Compra", Produto.Text, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25605 Then gError 132517

            'Se o produto não existir ==> erro
            If lErro = 25605 Then gError 132518

            sProd = Produto.Text

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 132519

            'Preenche o Produto com o ProdutoEnxuto
            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            'Verifica se já existe o Produto em outra linha do Grid
            For iIndice = 1 To objGridProdutos.iLinhasExistentes
                If iIndice <> GridProdutos.Row Then
                    If GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 132520
                End If
            Next

            'Verifica se a linha atual do Grid é maior que o número de linhas existentes
            If GridProdutos.Row > objGridProdutos.iLinhasExistentes Then

                'Adiciona o Fornecedor lido em colFornecedor
                lErro = Fornecedores_Adiciona(objProduto)
                If lErro <> SUCESSO Then gError 132521
                
                'Incrementa o número de linhas existentes do GridProdutos
                objGridProdutos.iLinhasExistentes = objGridProdutos.iLinhasExistentes + 1

            End If

            If Not (Me.ActiveControl Is Produto) Then

                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col) = Produto.Text

                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 132522

            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 132516, 132517, 132519, 32521, 132522

        Case 132518
            'Pergunta se deseja criar produto
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de Produtos
                Call Chama_Tela("Produto", objProduto)
            End If

        Case 132520
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_GRID_PRODUTOS", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161320)

    End Select

    Exit Sub
    
End Sub
'######################################################

