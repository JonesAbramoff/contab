VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl FornFilialProdutoOcx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9255
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4800
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   855
      Width           =   8970
      Begin VB.CommandButton BotaoProdutoFornecedor 
         Caption         =   "Produto x Fornecedor"
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
         Left            =   2715
         TabIndex        =   79
         Top             =   4320
         Width           =   2220
      End
      Begin VB.Frame Frame8 
         Caption         =   "Fornecedores do Produto"
         Height          =   2670
         Left            =   300
         TabIndex        =   8
         Top             =   1395
         Width           =   7605
         Begin VB.OptionButton PadraoGrid 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3060
            TabIndex        =   12
            Top             =   405
            Width           =   825
         End
         Begin MSMask.MaskEdBox FilialFornGrid 
            Height          =   225
            Left            =   3870
            TabIndex        =   11
            Top             =   375
            Visible         =   0   'False
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornGrid 
            Height          =   225
            Left            =   390
            TabIndex        =   10
            Top             =   390
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridFiliaisFornecedores 
            Height          =   2205
            Left            =   210
            TabIndex        =   9
            Top             =   300
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Produto"
         Height          =   1245
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   45
         Width           =   7590
         Begin VB.CheckBox Fixar 
            Caption         =   "Fixar"
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
            Left            =   3375
            TabIndex        =   5
            Top             =   315
            Width           =   900
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   1470
            TabIndex        =   4
            Top             =   285
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Descricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1470
            TabIndex        =   7
            Top             =   780
            Width           =   3840
         End
         Begin VB.Label Label2 
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
            Index           =   0
            Left            =   405
            TabIndex        =   6
            Top             =   810
            Width           =   930
         End
         Begin VB.Label LabelProduto 
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
            Left            =   660
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   3
            Top             =   315
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4800
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   855
      Visible         =   0   'False
      Width           =   8970
      Begin VB.Frame Frame4 
         Caption         =   "Fornecedor"
         Height          =   2940
         Left            =   285
         TabIndex        =   19
         Top             =   1575
         Width           =   7515
         Begin VB.TextBox DescFornProd 
            Height          =   315
            Left            =   1740
            MaxLength       =   150
            TabIndex        =   25
            Top             =   1950
            Width           =   5295
         End
         Begin VB.ComboBox Fornecedor 
            Height          =   315
            ItemData        =   "FornFilialProdutoOcx.ctx":0000
            Left            =   1740
            List            =   "FornFilialProdutoOcx.ctx":0002
            TabIndex        =   21
            Top             =   375
            Width           =   2835
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   1725
            TabIndex        =   22
            Top             =   900
            Width           =   1860
         End
         Begin VB.TextBox ProdutoFornecedor 
            Height          =   315
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1410
            Width           =   1455
         End
         Begin VB.CheckBox Padrao 
            Caption         =   "Padrão"
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
            Left            =   4260
            TabIndex        =   23
            Top             =   975
            Width           =   945
         End
         Begin MSMask.MaskEdBox LoteMinimo 
            Height          =   315
            Left            =   1710
            TabIndex        =   26
            Top             =   2415
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Nota 
            Height          =   315
            Left            =   4245
            TabIndex        =   27
            Top             =   2415
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDescFornProd 
            AutoSize        =   -1  'True
            Caption         =   "Descrição Forn:"
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
            TabIndex        =   80
            Top             =   1980
            Width           =   1365
         End
         Begin VB.Label Label15 
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
            Left            =   1155
            TabIndex        =   28
            Top             =   960
            Width           =   465
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nota:"
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
            Left            =   3675
            TabIndex        =   30
            Top             =   2475
            Width           =   480
         End
         Begin VB.Label LabelForn 
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
            Left            =   585
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   435
            Width           =   1035
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Código no Forn.:"
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
            Left            =   210
            TabIndex        =   29
            Top             =   1470
            Width           =   1425
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Lote Mínimo:"
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
            Left            =   495
            TabIndex        =   31
            Top             =   2475
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Produto"
         Height          =   1245
         Index           =   1
         Left            =   270
         TabIndex        =   14
         Top             =   240
         Width           =   7500
         Begin VB.Label Prod 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   1395
            TabIndex        =   16
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   15
            Top             =   315
            Width           =   660
         End
         Begin VB.Label Label2 
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
            Index           =   1
            Left            =   330
            TabIndex        =   17
            Top             =   825
            Width           =   930
         End
         Begin VB.Label DescProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1395
            TabIndex        =   18
            Top             =   780
            Width           =   3840
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7020
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "FornFilialProdutoOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FornFilialProdutoOcx.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FornFilialProdutoOcx.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FornFilialProdutoOcx.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4800
      Index           =   3
      Left            =   120
      TabIndex        =   36
      Top             =   855
      Visible         =   0   'False
      Width           =   8970
      Begin VB.Frame Frame2 
         Caption         =   "Última Compra"
         Height          =   660
         Left            =   300
         TabIndex        =   65
         Top             =   2880
         Width           =   8580
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total do Produto:"
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
            Left            =   4305
            TabIndex        =   69
            Top             =   315
            Width           =   1995
         End
         Begin VB.Label PrecoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6435
            TabIndex        =   68
            Top             =   255
            Width           =   1485
         End
         Begin VB.Label DataUltimaCompra 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1890
            TabIndex        =   67
            Top             =   255
            Width           =   1455
         End
         Begin VB.Label Label22 
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
            Index           =   1
            Left            =   1260
            TabIndex        =   66
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pedidos de Compra"
         Height          =   1080
         Left            =   300
         TabIndex        =   58
         Top             =   480
         Width           =   8580
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Prazo médio de entrega:"
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
            Left            =   1485
            TabIndex        =   64
            Top             =   720
            Width           =   2085
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade em Pedidos Abertos:"
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
            Left            =   780
            TabIndex        =   63
            Top             =   315
            Width           =   2790
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "dias"
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
            Left            =   4380
            TabIndex        =   62
            Top             =   720
            Width           =   360
         End
         Begin VB.Label QuantPedAbertos 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3720
            TabIndex        =   61
            Top             =   255
            Width           =   1455
         End
         Begin VB.Label TempoRessup 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3720
            TabIndex        =   60
            Top             =   660
            Width           =   555
         End
         Begin VB.Label UMCompra 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   5250
            TabIndex        =   59
            Top             =   255
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Último Pedido de Compra Recebido"
         Height          =   1170
         Left            =   300
         TabIndex        =   37
         Top             =   1635
         Width           =   8580
         Begin VB.Label DataReceb 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6975
            TabIndex        =   47
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label QuantPedida 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2355
            TabIndex        =   46
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label QuantRecebida 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2355
            TabIndex        =   45
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label DataPedido 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6975
            TabIndex        =   44
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label UMCompra 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   3930
            TabIndex        =   43
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Data do Recebimento:"
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
            Left            =   4980
            TabIndex        =   42
            Top             =   780
            Width           =   1920
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Data do Pedido:"
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
            Left            =   5505
            TabIndex        =   41
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade Recebida:"
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
            Left            =   360
            TabIndex        =   40
            Top             =   780
            Width           =   1920
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade Pedida:"
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
            Left            =   600
            TabIndex        =   39
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label UMCompra 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   3930
            TabIndex        =   38
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Última Cotação"
         Height          =   1110
         Left            =   300
         TabIndex        =   48
         Top             =   3630
         Width           =   8580
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Frete:"
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
            Left            =   6570
            TabIndex        =   72
            Top             =   300
            Width           =   510
         End
         Begin VB.Label TipoFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7185
            TabIndex        =   71
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label17 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1305
            TabIndex        =   57
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Preço Unitário:"
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
            Left            =   5025
            TabIndex        =   56
            Top             =   750
            Width           =   1290
         End
         Begin VB.Label UMCompra 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   3465
            TabIndex        =   55
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label20 
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
            Left            =   660
            TabIndex        =   54
            Top             =   750
            Width           =   1050
         End
         Begin VB.Label DataUltimaCotacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1890
            TabIndex        =   53
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label UltimaCotacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6435
            TabIndex        =   52
            Top             =   705
            Width           =   1470
         End
         Begin VB.Label QuantUltimaCotacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1860
            TabIndex        =   51
            Top             =   690
            Width           =   1470
         End
         Begin VB.Label CondPagtoUltimaCotacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4965
            TabIndex        =   50
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label19 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3765
            TabIndex        =   49
            Top             =   300
            Width           =   1065
         End
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
         Height          =   195
         Index           =   4
         Left            =   6375
         TabIndex        =   78
         Top             =   150
         Width           =   465
      End
      Begin VB.Label LabelFornFilial 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6975
         TabIndex        =   77
         Top             =   105
         Width           =   1890
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   2685
         TabIndex        =   76
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label LabelFornecedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3780
         TabIndex        =   75
         Top             =   105
         Width           =   2160
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   345
         TabIndex        =   74
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Prod 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1125
         TabIndex        =   73
         Top             =   105
         Width           =   1200
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5235
      Left            =   90
      TabIndex        =   0
      Top             =   480
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   9234
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produto X Fornecedor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fornecedor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Estatísticas"
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
Attribute VB_Name = "FornFilialProdutoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iProdutoAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iFrameAtual As Integer

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorProduto As AdmEvento
Attribute objEventoFornecedorProduto.VB_VarHelpID = -1

'GridFiliaisFornecedores
Dim objGridFiliaisFornecedores As AdmGrid
Dim iGrid_Padrao_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialFornecedor_Col As Integer

Private Sub BotaoProdutoFornecedor_Click()

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_BotaoProdutoFornecedor_Click

    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 89259
        
        objFornecedorProdutoFF.sProduto = sProdutoFormatado

    End If

    'Chama a tela EstoqueInicialLista
    Call Chama_Tela("FornFilialProdutoLista", colSelecao, objFornecedorProdutoFF, objEventoFornecedorProduto)

    Exit Sub
    
Erro_BotaoProdutoFornecedor_Click:

    Select Case gErr
    
        Case 89259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160680)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lPosicao As Long
Dim lErro As Long

On Error GoTo Erro_Filial_Click

    'Se não tiver filial selecionada sai da rotina
    If Filial.ListIndex = -1 Then Exit Sub

    'Se filial não estiver preenchida, sai da rotina
    If Len(Trim(Filial.Text)) = 0 Then
        LabelFornFilial.Caption = ""
        Exit Sub
    End If

    'Critica o formato do Produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 54169

    'Preenche objFornecedorProdutoFF
    objFornecedorProdutoFF.sProduto = sProdutoFormatado

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Verifica se Fornecedor existe
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 54211
    If lErro = 6681 Then Error 54210

    objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo

    objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa
    objFornecedorProdutoFF.iFilialForn = Codigo_Extrai(Filial.Text)

    iAlterado = REGISTRO_ALTERADO
    
    'Preenche LabelFornFilial
    lPosicao = Len(CStr(Codigo_Extrai(Filial.Text)) & SEPARADOR) + 1
    LabelFornFilial.Caption = Mid(Filial.Text, lPosicao)
    
    'Traz os dados do FornecedorProduto na tela
    lErro = Traz_FornecedorProdutoFF_Tela(objFornecedorProdutoFF)
    If lErro <> SUCESSO Then Error 54168

    Exit Sub

Erro_Filial_Click:

    Select Case Err

        Case 54168, 54169, 54211

        Case 54210
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160681)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim sNomeRed As String
Dim vbMsgRes As VbMsgBoxResult
Dim lPosicao As Long

On Error GoTo Erro_Filial_Validate

    'Se a filial não tiver sido preenchida, sai da rotina
    If Len(Trim(Filial.Text)) = 0 Then
        LabelFornFilial.Caption = ""
        Call Limpa_Campos_FornecedorProdutoFF
        Exit Sub
    End If

    'Se a filial tiver sido selecionada, sai da rotina
    If Filial.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 54284

    'Se não encontrou o ítem com o código informado
    If lErro = 6730 Then

        'Verifica se o Fornecedor foi preenchido
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 54286

        sNomeRed = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe a Filial do Fornecedor
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 54283

        'Se não encontrou a Filial do Fornecedor --> erro
        If lErro = 18272 Then Error 54285

        'Coloca a Filial do Fornecedor na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        'Preenche LabelFornFilial
        lPosicao = Len(CStr(Codigo_Extrai(Filial.Text)) & SEPARADOR) + 1
        LabelFornFilial.Caption = Mid(Filial.Text, lPosicao)

        'Traz os dados da FornecedorProdutoFF para a Tela
        lErro = Traz_FornecedorProdutoFF_Tela(objFornecedorProdutoFF)
        If lErro <> SUCESSO Then Error 54291

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 54287

    Exit Sub

Erro_Filial_Validate:

    Cancel = True
    
    Select Case Err

        Case 54283, 54284

        Case 54285
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela FiliaisFornecedores
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 54286
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 54287
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_INEXISTENTE", Err)

        Case 54291

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160682)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objFornecedor As New ClassFornecedor
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFixado As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Produto foi fixado
    If Fixar.Value = vbChecked Then
        sProdutoFixado = Produto.Text
    End If
    
    'Verifica preenchimento de produto
    If Len(Trim(Produto.Text)) = 0 Then Error 54147

    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 54148

    'Verifica preenchimento de Filial
    If Len(Trim(Filial.Text)) = 0 Then Error 54149

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê Nome Reduzido do Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 54150

    If lErro = 6681 Then Error 54151

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then Error 54152

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 54153

    If lErro = 25041 Then Error 54154

    'Preenche objFornecedorProdutoFF
    objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo
    objFornecedorProdutoFF.sProduto = objProduto.sCodigo
    objFornecedorProdutoFF.iFilialForn = Codigo_Extrai(Filial.Text)
    objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa

    'Lê FornecedorProduto
    lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
    If lErro <> SUCESSO And lErro <> 54217 Then Error 54155

    If lErro = 54217 Then Error 54156

    'Pede confirmação para exclusão Fornecedor Produto
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FORNECEDOR_PRODUTO", objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto)

    If vbMsgRes = vbYes Then

        lErro = CF("FornecedorProdutoFF_Exclui", objFornecedorProdutoFF)
        If lErro <> SUCESSO Then Error 54157
    
        'Excluir o fornecedor da combo
        Call Combo_Item_Igual_Remove(Fornecedor)
    
        'Limpa a Filial da combo
        Filial.Clear
    
        'Limpa dados do fornecedor produto
        Call Limpa_Tela_FornecedorProdutoFF

        'Preenche Produto
        If Len(Trim(sProdutoFixado)) > 0 Then
            Produto.Text = sProdutoFixado
            Produto_Validate (bCancel)
        End If
        
        iAlterado = 0
        
        Call ComandoSeta_Fechar(Me.Name)
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 54148
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 54147, 54153
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 54149
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 54150, 54152, 54155, 54157

        Case 54154
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 54156
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORFILIALPRODUTO_NAO_CADASTRADA", Err, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.lFornecedor)

        Case 54151
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160683)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim sForn As String
Dim sFilial As String
Dim sProduto As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se o Produto foi fixado
    If Fixar.Value = vbChecked Then
        sProduto = Produto.Text
    End If
    
    'Chama rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 54158

    'Excluir e incluir o fornecedor na combo
    sForn = Fornecedor.Text
    Call Combo_Item_Igual_Remove(Fornecedor)
    Fornecedor.AddItem sForn

    Fornecedor.ListIndex = -1

    'Excluir e incluir a filial na combo
    sFilial = Filial.Text
    Call Combo_Item_Igual_Remove(Filial)
    Filial.AddItem sFilial

    Filial.ListIndex = -1

    'Limpa a tela
    Call Limpa_Tela_FornecedorProdutoFF

    'Colocar o Produto fixado na tela
    If Len(Trim(sProduto)) > 0 Then
        Produto.Text = sProduto
        Produto_Validate (bCancel)
    End If
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 54158 'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160684)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 54159

    Fixar.Value = vbUnchecked
    
    'Limpa a tela
    Call Limpa_Tela_FornecedorProdutoFF

    iAlterado = 0

    'Fecha o comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 54159

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160685)

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
Dim iIndice As Integer
Dim colFornecedor As New Collection
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objEventoProduto = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoFornecedorProduto = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 54162

    Set objGridFiliaisFornecedores = New AdmGrid
    
    lErro = Inicializa_Grid_FiliaisFornecedores(objGridFiliaisFornecedores)
    If lErro <> SUCESSO Then gError 74829
    
    iAlterado = 0
    iProdutoAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 54162, 74829

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160686)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long

Dim lErro As Long
Dim sCodigo As String
Dim bCancel As Boolean

On Error GoTo Erro_Trata_Parametros

    'Se há um FornecedorProduto selecionado, exibir seus dados
    If Not (objFornecedorProdutoFF Is Nothing) Then
        
'        If objFornecedorProdutoFF.lFornecedor > 0 Then
'
'            'Verifica se o FornecedorProduto existe
'            lErro = CF("FornecedorProdutoFF_Le",objFornecedorProdutoFF)
'            If lErro <> SUCESSO And lErro <> 54217 Then Error 54166
'
'        End If
'
        'Não encontrou Fornecedor do Produto
        sCodigo = objFornecedorProdutoFF.sProduto
        lErro = CF("Traz_Produto_MaskEd", sCodigo, Produto, Descricao)
        If lErro <> SUCESSO Then Error 54165
        
        'Dispara o Validate de Produto
        Call Produto_Validate(bCancel)

        If objFornecedorProdutoFF.lFornecedor > 0 And objFornecedorProdutoFF.iFilialForn > 0 And lErro = SUCESSO Then

            'Mostra Fornecedor na Tela
            Fornecedor.Text = objFornecedorProdutoFF.lFornecedor

            'Dispara o Validate de fornecedor
            Fornecedor_Validate (bCancel)

            'Mostra a Filial na tela
            Filial.Text = objFornecedorProdutoFF.iFilialForn

            'Dispara o Validate de Filial
            Filial_Validate (bCancel)

        End If

    End If

    iAlterado = 0
    iProdutoAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 54165, 54166

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160687)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoProduto = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoFornecedorProduto = Nothing
    Set objGridFiliaisFornecedores = Nothing
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Inicializa_Grid_FiliaisFornecedores(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid de Filiais Fornecedores

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_FiliaisFornecedores

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Padrão")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")

    'campos de edição do grid
    objGridInt.colCampo.Add (PadraoGrid.Name)
    objGridInt.colCampo.Add (FornGrid.Name)
    objGridInt.colCampo.Add (FilialFornGrid.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Padrao_Col = 1
    iGrid_Fornecedor_Col = 2
    iGrid_FilialFornecedor_Col = 3

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridFiliaisFornecedores

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_FORNFILIALFF + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    GridFiliaisFornecedores.ColWidth(0) = 450

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FiliaisFornecedores = SUCESSO

    Exit Function

Erro_Inicializa_Grid_FiliaisFornecedores:

    Inicializa_Grid_FiliaisFornecedores = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 160688)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lFornecedor As Long

On Error GoTo Erro_Fornecedor_Click

    'Se nenhum fornecedor foi selecionado, sai da rotina
    If Fornecedor.ListIndex = -1 Then Exit Sub

    'Limpa a combo de Filiais
    Filial.Clear
    
    'Limpa LabelFornecedor
    LabelFornecedor.Caption = ""

    'Limpa os Campos da tela
    Call Limpa_Campos_FornecedorProdutoFF

    'preencher objFornecedorProdutoFF a partir dos campos de Produto e Fornecedor
    'Verifica se Fornecedor existe
    objFornecedor.sNomeReduzido = Fornecedor.Text
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 54170

    If lErro = 6681 Then Error 54174

    'Preenche LabelFornecedor
    LabelFornecedor.Caption = objFornecedor.sNomeReduzido

    'Carrega as filiais do fornecedor em questão
    lErro = Carrega_FiliaisFornecedores(objFornecedor.lCodigo)
    If lErro <> SUCESSO Then Error 54173

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Fornecedor_Click:

    Select Case Err

        Case 54170, 54173

        Case 54174
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160689)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FiliaisFornecedores(lFornecedor As Long) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FiliaisFornecedores

    objFornecedor.lCodigo = lFornecedor

    'Le as filiais do fornecedor em questão
    lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
    If lErro <> SUCESSO Then Error 54175

    'Preenche Combobox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)
    
    'Se foram carregadas filiais
    If Filial.ListCount > 0 Then
        
        'Seleciona a primeira  filial
        Filial.ListIndex = 0
        
    End If
    
    Carrega_FiliaisFornecedores = SUCESSO

    Exit Function

Erro_Carrega_FiliaisFornecedores:

    Carrega_FiliaisFornecedores = Err

    Select Case Err

        Case 54175

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160690)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objFornecedor As New ClassFornecedor
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lCodigo As Long
Dim lFornecedor As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Error_Fornecedor_Validate

    'Se Fornecedor não foi alterado sai da rotina
    If iFornecedorAlterado = 0 Then Exit Sub
    
    'Se o Fornecedor tiver sido selecionado, sai da rotina
    If Fornecedor.ListIndex <> -1 Then Exit Sub

    'Limpa LabelFornecedor
    LabelFornecedor.Caption = ""
    
    'Limpa os Campos da tela
    Call Limpa_Campos_FornecedorProdutoFF
    
    'Verifica se foi preenchida a ComboBox Fornecedor
    If Len(Trim(Fornecedor.Text)) = 0 Then
        Filial.Clear
        Exit Sub
    End If

    'Limpa a combo Filial
    Filial.Clear

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = LCombo_Seleciona(Fornecedor, lCodigo)
    If lErro <> SUCESSO And lErro <> 25639 And lErro <> 25640 Then Error 54176

    'Não encontrou o código do Fornecedor
    If lErro = 25639 Then

        objFornecedor.lCodigo = lCodigo

        'Lê FornecedorProdutoFF
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 54177
        If lErro <> SUCESSO Then Error 54172

        'Coloca na tela o Nome Reduzido do fornecedor
        Fornecedor.Text = objFornecedor.sNomeReduzido

    'Não encontrou o NomeReduzido do Fornecedor ou encontrou
    Else
        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Verifica se Fornecedor existe
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 54213
        If lErro <> SUCESSO Then Error 54212

    End If

    'Preenche LabelFornecedor
    LabelFornecedor.Caption = objFornecedor.sNomeReduzido
    
    'Carrega as filiais do Fornecedor em questão
    lErro = Carrega_FiliaisFornecedores(objFornecedor.lCodigo)
    If lErro <> SUCESSO Then Error 54178
    
    iFornecedorAlterado = 0

    Exit Sub

Error_Fornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 54172
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", objFornecedor.lCodigo)
            If vbMsgRes = vbYes Then
                'Chama a tela de Fornecedores
                Call Chama_Tela("Fornecedores", objFornecedor)
            End If

        Case 54176

        Case 54177, 54178, 54213

        Case 54212
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160691)

    End Select

    Exit Sub

End Sub

Private Function Mostra_FornecedorProdutoFF_Tela(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Mostra os dados de objFornecedorProdutoFF, quando o produto e o fornecedor já estão selecionados

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_Mostra_FornecedorProdutoFF_Tela
    
    objProduto.sCodigo = objFornecedorProdutoFF.sProduto
    
    'Le no Produto a Unidade de Compras
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 89355
    
    If lErro = 28030 Then gError 89356
        
    'Mostra os dados do Produto do Frame Fornecedor na tela
    If objFornecedorProdutoFF.sProdutoFornecedor <> "" Then
        ProdutoFornecedor.Text = objFornecedorProdutoFF.sProdutoFornecedor
    Else
        ProdutoFornecedor.Text = ""
    End If
    
    If objFornecedorProdutoFF.iNota <> 0 Then
        Nota.Text = CStr(objFornecedorProdutoFF.iNota)
    Else
        Nota.Text = ""
    End If
    
    If objFornecedorProdutoFF.dLoteMinimo <> 0 Then
        LoteMinimo.Text = Formata_Estoque(objFornecedorProdutoFF.dLoteMinimo)
    Else
        LoteMinimo.Text = ""
    End If

    DescFornProd.Text = objFornecedorProdutoFF.sDescricao

    'Mostra dos dados do Frame Estatísticas na tela
    'Pedido de Compra
    If objFornecedorProdutoFF.dQuantPedAbertos <> 0 Then
        QuantPedAbertos.Caption = CStr(objFornecedorProdutoFF.dQuantPedAbertos)
    Else
        QuantPedAbertos.Caption = ""
    End If
    
    If objFornecedorProdutoFF.dTempoRessup <> 0 Then
        TempoRessup.Caption = CStr(objFornecedorProdutoFF.dTempoRessup)
    Else
        TempoRessup.Caption = ""
    End If
    
    'Último Pedido de Compra Recebido
    
    'Se a U.M. da Quant Pedida foi preenchida
    If Len(Trim(objFornecedorProdutoFF.sUMQuantPedida)) > 0 Then
    
        'Converte as unidades de medida
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objFornecedorProdutoFF.sUMQuantPedida, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 89357
    
    End If
    
    'Quantidade Pedida
    If objFornecedorProdutoFF.dQuantPedida <> 0 Then
        QuantPedida.Caption = CStr(objFornecedorProdutoFF.dQuantPedida * dFator)
    Else
        QuantPedida.Caption = ""
    End If
    
    'Se a U.M. da Quant Recebida foi preenchida
    If Len(Trim(objFornecedorProdutoFF.sUMQuantRecebida)) > 0 Then
    
        'Converte as unidades de medida
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objFornecedorProdutoFF.sUMQuantRecebida, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 89358
    
    End If
    
    If objFornecedorProdutoFF.dQuantRecebida <> 0 Then
        QuantRecebida.Caption = CStr(objFornecedorProdutoFF.dQuantRecebida * dFator)
    Else
        QuantRecebida.Caption = ""
    End If
    
    If objFornecedorProdutoFF.dtDataPedido <> DATA_NULA Then
        DataPedido.Caption = Format(objFornecedorProdutoFF.dtDataPedido, "dd/mm/yy")
    Else
        DataPedido.Caption = ""
    End If
    
    If objFornecedorProdutoFF.dtDataReceb <> DATA_NULA Then
        DataReceb.Caption = Format(objFornecedorProdutoFF.dtDataReceb, "dd/mm/yy")
    Else
        DataReceb.Caption = ""
    End If

    'Última Compra
    If objFornecedorProdutoFF.dPrecoTotal <> 0 Then
        PrecoTotal.Caption = Format(objFornecedorProdutoFF.dPrecoTotal, "Standard")
    Else
        PrecoTotal.Caption = ""
    End If
    
    If objFornecedorProdutoFF.dtDataUltimaCompra <> DATA_NULA Then
        DataUltimaCompra.Caption = Format(objFornecedorProdutoFF.dtDataUltimaCompra, "dd/mm/yy")
    Else
        DataUltimaCompra.Caption = ""
    End If
    
    'Última Cotação
    
    'Se a U.M. da Quant da última cotação foi preenchida
    If Len(Trim(objFornecedorProdutoFF.sUMQuantUltimaCotacao)) > 0 Then
    
        'Converte as unidades de medida
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objFornecedorProdutoFF.sUMQuantUltimaCotacao, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 89359
    
    End If
    
    'Quantidade da última cotação
    If objFornecedorProdutoFF.dQuantUltimaCotacao <> 0 Then
        QuantUltimaCotacao.Caption = CStr(objFornecedorProdutoFF.dQuantUltimaCotacao * dFator)
    Else
        QuantUltimaCotacao.Caption = ""
    End If
    
    If objFornecedorProdutoFF.dUltimaCotacao <> 0 Then
        UltimaCotacao.Caption = Format(objFornecedorProdutoFF.dUltimaCotacao, "Fixed")
    Else
        UltimaCotacao.Caption = ""
    End If
    
    If objFornecedorProdutoFF.dtDataUltimaCotacao <> DATA_NULA Then
        DataUltimaCotacao.Caption = Format(objFornecedorProdutoFF.dtDataUltimaCotacao, "dd/mm/yy")
    Else
        DataUltimaCotacao.Caption = ""
    End If
    
    If objFornecedorProdutoFF.iCondPagto <> 0 Then
        CondPagtoUltimaCotacao.Caption = objFornecedorProdutoFF.sCondPagto
    Else
        CondPagtoUltimaCotacao.Caption = ""
    End If
    
    If objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_CIF Then
        TipoFrete.Caption = "CIF"
    ElseIf objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_FOB Then
        TipoFrete.Caption = "FOB"
    End If
    
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    objProdutoFilial.sProduto = objFornecedorProdutoFF.sProduto

    'Lê o ProdutoFilial
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 54180
    
    If lErro <> 28261 Then
    
        If objProdutoFilial.lFornecedor <> objFornecedorProdutoFF.lFornecedor Or objFornecedorProdutoFF.iFilialForn <> objProdutoFilial.iFilialForn Then
            Padrao.Value = FORN_PROD_NAO_PADRAO
        Else
            Padrao.Value = 1
        End If
        
    End If

    iAlterado = 0
    
    Mostra_FornecedorProdutoFF_Tela = SUCESSO

    Exit Function

Erro_Mostra_FornecedorProdutoFF_Tela:

    Mostra_FornecedorProdutoFF_Tela = gErr

    Select Case gErr

        Case 54180, 89355, 89357, 89358, 89359

        Case 89356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160692)

    End Select

    Exit Function

End Function

Private Sub LabelForn_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Preenche objFornecedor com NomeReduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82659

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 82659

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160693)

    End Select

    Exit Sub

End Sub

Private Sub Nota_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Nota, iAlterado)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedor As ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objFornProdFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1
    
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 82664
    
    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Coloca nome no LabelFornecedor
    LabelFornecedor.Caption = objFornecedor.sNomeReduzido

    'Limpa os Campos da tela
    Call Limpa_Campos_FornecedorProdutoFF

    'Carrega as filiais do fornecedor em questão
    lErro = Carrega_FiliaisFornecedores(objFornecedor.lCodigo)
    If lErro <> SUCESSO Then gError 82665
    
    'Critica o formato do Produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 82666

    For iIndice = 0 To Filial.ListCount - 1
        objFornProdFF.sProduto = sProdutoFormatado
        objFornProdFF.lFornecedor = objFornecedor.lCodigo
        objFornProdFF.iFilialEmpresa = giFilialEmpresa
        objFornProdFF.iFilialForn = Filial.ItemData(iIndice)
        
        lErro = CF("FornecedorProdutoFF_Le", objFornProdFF)
        If lErro <> SUCESSO And lErro <> 54217 Then gError 82667
        If lErro = SUCESSO Then
            Filial.ListIndex = iIndice
            Exit For
        End If
                
    Next

    Me.Show
    
    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case gErr

        Case 82664
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
            Fornecedor.SetFocus

        Case 82665, 82666, 82667

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160694)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82660

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82661
    
    'Verifica se Produto é gerencial
    If objProduto.iGerencial = GERENCIAL Then Exit Sub
    If objProduto.iCompras <> PRODUTO_COMPRAVEL Then gError 82662

    'Coloca a unidade de medida na tela
    UMCompra(0).Caption = objProduto.sSiglaUMCompra
    UMCompra(1).Caption = objProduto.sSiglaUMCompra
    UMCompra(2).Caption = objProduto.sSiglaUMCompra
    UMCompra(3).Caption = objProduto.sSiglaUMCompra
    
    'Coloca o código na tela
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 82663

    'Dispara o Validate no código do produto
    Produto_Validate (bCancel)

    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 82660, 82663

        Case 82661
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 82662
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160695)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedorProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornProdFF As New ClassFornecedorProdutoFF
Dim bCancel As Boolean

On Error GoTo Erro_objEventoFornecedorProduto_evSelecao

    Set objFornProdFF = obj1

    'Coloca o código na tela
    lErro = CF("Traz_Produto_MaskEd", objFornProdFF.sProduto, Produto, Descricao)
    If lErro <> SUCESSO Then gError 82663

    'Dispara o Validate no código do produto
    Produto_Validate (bCancel)

    'Maristela
    DescFornProd.Text = objFornProdFF.sDescricao
    
    Me.Show
    
    Exit Sub

Erro_objEventoFornecedorProduto_evSelecao:

    Select Case gErr

        Case 89260

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160696)

    End Select

    Exit Sub

End Sub

Private Sub TempoRessup_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteMinimo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteMinimo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteMinimo_Validate

    'Se o lote mínimo não estiver preenchido, sai da rotina
    If Len(Trim(LoteMinimo.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(LoteMinimo.Text)
    If lErro <> SUCESSO Then Error 54181
    
    LoteMinimo.Text = Formata_Estoque(StrParaDbl(LoteMinimo.Text))

    Exit Sub

Erro_LoteMinimo_Validate:

    Cancel = True
    
    Select Case Err

        Case 54181 'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160697)

    End Select

    Exit Sub

End Sub

Private Sub Nota_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Padrao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Change()

    iProdutoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim colFornecedores As New AdmCollCodigoNome
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Produto_Validate

    If iProdutoAlterado = 0 Then Exit Sub

    'Limpa a combo Fornecedor
    Fornecedor.Clear
    'Limpa a combo Filial
    Filial.Clear

    'Limpa labels de Fornecedor e FornFilial
    LabelFornecedor.Caption = ""
    LabelFornFilial.Caption = ""

    'Limpa os Campos da tela
    Call Limpa_Campos_FornecedorProdutoFF

    'Se o produto estiver preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        sProduto = Produto.Text

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 54292

        'Se não encontro o produto -> Erro
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 54293

        If lErro = 25041 Then gError 54295

        'O Produto é gerencial
        If lErro = 25043 Then gError 54297
        
        'If objProduto.iCompras <> PRODUTO_COMPRAVEL Then Error 62703

        'Preenche Descricao com Descrição do Produto e UnidMed
        Descricao.Caption = objProduto.sDescricao
        UMCompra.Item(0).Caption = objProduto.sSiglaUMCompra
        UMCompra.Item(1).Caption = objProduto.sSiglaUMCompra
        UMCompra.Item(2).Caption = objProduto.sSiglaUMCompra
        UMCompra.Item(3).Caption = objProduto.sSiglaUMCompra
        
        'Carrega os Fornecedores associados a este Produto
        lErro = Carrega_FornecedorProduto(colFornecedores, objProduto)
        If lErro <> SUCESSO Then gError 54296

        If Fornecedor.ListCount <> 0 Then Fornecedor.ListIndex = 0

        'Preenche o GridFiliaisFornecedores com os Fornecedores da colecao
        lErro = GridFiliaisFornecedores_Preenche(objProduto)
        If lErro <> SUCESSO Then gError 74830
        
        'Preenche Prod e DescProd
        Prod(0).Caption = Produto.Text
        Prod(1).Caption = Produto.Text
        DescProd.Caption = objProduto.sDescricao
        
    'Se não estiver peenchido
    Else
        'Limpa os campos
        Descricao.Caption = ""
        UMCompra(0).Caption = ""
        UMCompra(1).Caption = ""
        UMCompra(2).Caption = ""
        UMCompra(3).Caption = ""
        
        Prod(0).Caption = ""
        Prod(1).Caption = ""
        DescProd.Caption = ""
        
        Call Grid_Limpa(objGridFiliaisFornecedores)
    End If

    iProdutoAlterado = 0

    Exit Sub

Erro_Produto_Validate:

    Cancel = True
    
    Select Case gErr

        Case 54292, 74830

        Case 54293, 54295 'Não encontrou Produto no BD

            'Pergunta se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                Descricao.Caption = ""
                UMCompra(0).Caption = ""
                UMCompra(1).Caption = ""
                UMCompra(2).Caption = ""
                UMCompra(3).Caption = ""
                
            End If

        Case 54296

        Case 54297
            Call Rotina_Erro(vbOKOnly, "AVISO_PRODUTO_TEM_FILHOS", gErr)

        Case 62703
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", Err, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160698)

    End Select

    Exit Sub

End Sub
Private Function GridFiliaisFornecedores_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim colFornecedorProdutoFF As New Collection
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim iLinha As Integer
Dim objFornecedor As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_GridFiliaisFornecedores_Preenche

    'Limpa o GridFiliaisFornecedores
    Call Grid_Limpa(objGridFiliaisFornecedores)
    
    'Lê os Fornecedores/Filiais relacionados ao Produto informado
    lErro = CF("FornecedoresProdutoFF_Le", colFornecedorProdutoFF, objProduto)
    If lErro <> SUCESSO And lErro <> 63156 Then gError 74831
    
    If lErro <> 63156 Then
    
        'Se o número de Filiais/Fornecedores relacionados for maior que o número de linhas do Grid
        If colFornecedorProdutoFF.Count >= GridFiliaisFornecedores.Rows Then
        
            GridFiliaisFornecedores.Rows = colFornecedorProdutoFF.Count + 1
            
            Call Grid_Inicializa(objGridFiliaisFornecedores)
                
        End If
            
        For Each objFornecedorProdutoFF In colFornecedorProdutoFF
        
            iLinha = iLinha + 1
            
            objProdutoFilial.sProduto = objFornecedorProdutoFF.sProduto
            objProdutoFilial.iFilialEmpresa = objFornecedorProdutoFF.iFilialEmpresa
            
            lErro = CF("ProdutoFilial_Le", objProdutoFilial)
            If lErro <> SUCESSO And lErro <> 28261 Then gError 74835
            
            'Verifica se o par Fornecedor/Filial é padrão para o Produto
            If objFornecedorProdutoFF.lFornecedor = objProdutoFilial.lFornecedor And objFornecedorProdutoFF.iFilialForn = objProdutoFilial.iFilialForn Then
                GridFiliaisFornecedores.TextMatrix(iLinha, iGrid_Padrao_Col) = vbChecked
            End If
            
            objFornecedor.lCodigo = objFornecedorProdutoFF.lFornecedor
            
            'Lê o Fornecedor cujo código foi informado
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 74832
            
            GridFiliaisFornecedores.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
            
            objFilialForn.iCodFilial = objFornecedorProdutoFF.iFilialForn
            objFilialForn.lCodFornecedor = objFornecedorProdutoFF.lFornecedor
            
            'Lê a Filial Fornecedor cujo código foi informado
            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 74833
            
            GridFiliaisFornecedores.TextMatrix(iLinha, iGrid_FilialFornecedor_Col) = objFilialForn.iCodFilial & SEPARADOR & objFilialForn.sNome
            
            objGridFiliaisFornecedores.iLinhasExistentes = iLinha
            
        Next
            
        Call Grid_Refresh_Checkbox(objGridFiliaisFornecedores)
    
    End If
    
    GridFiliaisFornecedores_Preenche = SUCESSO
    
    Exit Function
    
Erro_GridFiliaisFornecedores_Preenche:
    
    GridFiliaisFornecedores_Preenche = gErr
    
    Select Case gErr
    
        Case 74831, 74832, 74833, 74835
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160699)
        
    End Select
    
    Exit Function
    
End Function
Private Sub ProdutoFornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Carrega_FornecedorProduto(colFornecedores As AdmCollCodigoNome, objProduto As ClassProduto) As Long
'Carrega a ComboBox Fornecedor

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Carrega_FornecedorProduto

    'Preenche colecao de Fornecedores associados a sProduto existentes no BD
    lErro = CF("FornecedorProdutoFF_Le_Fornecedores", colFornecedores, objProduto)
    If lErro <> SUCESSO Then Error 54294

    For iIndice = 1 To colFornecedores.Count
        
        'Preenche a ComboBox Fornecedor com os objetos da colecao colFornecedores
        Fornecedor.AddItem colFornecedores(iIndice).sNome
        Fornecedor.ItemData(Fornecedor.NewIndex) = colFornecedores.Item(iIndice).lCodigo
    
    Next

    Carrega_FornecedorProduto = SUCESSO

    Exit Function

Erro_Carrega_FornecedorProduto:

    Carrega_FornecedorProduto = Err

    Select Case Err

        Case 54294

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160700)

    End Select

    Exit Function

End Function
Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

    Exit Sub

End Sub

Function Limpa_Tela_FornecedorProdutoFF() As Long
'Limpa a Tela FornecedorProduto

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_FornecedorProdutoFF

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then Error 54190

    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)

    'Limpa os campos da tela que não foram limpos pela função acima
    Prod(0).Caption = ""
    Prod(1).Caption = ""
    DescProd.Caption = ""
    LabelFornFilial.Caption = ""
    LabelFornecedor.Caption = ""
    Descricao.Caption = ""
    Fornecedor.Clear
    Filial.Clear
    ProdutoFornecedor.Text = ""
    Padrao.Value = 0

    'Frame Estatísticas
    UMCompra(0).Caption = ""
    UMCompra(1).Caption = ""
    UMCompra(2).Caption = ""
    UMCompra(3).Caption = ""
    
    'Pedidos de Compra
    QuantPedAbertos.Caption = ""
    TempoRessup.Caption = ""

    'Último Pedido de Compra Recebido
    QuantPedida.Caption = ""
    QuantRecebida.Caption = ""
    DataPedido.Caption = ""
    DataReceb.Caption = ""
    
    'Última Compra
    PrecoTotal.Caption = ""
    DataUltimaCompra.Caption = ""
    
    'Última Cotação
    QuantUltimaCotacao.Caption = ""
    UltimaCotacao.Caption = ""
    DataUltimaCotacao.Caption = ""
    CondPagtoUltimaCotacao.Caption = ""
    TipoFrete.Caption = ""
    
    Call Grid_Limpa(objGridFiliaisFornecedores)
    
    iAlterado = 0

    Limpa_Tela_FornecedorProdutoFF = SUCESSO

    Exit Function

Erro_Limpa_Tela_FornecedorProdutoFF:

    Select Case Err

        Case 54190

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160701)

    End Select

    Exit Function

End Function
Private Sub GridFiliaisFornecedores_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFiliaisFornecedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFiliaisFornecedores, iAlterado)
    End If

End Sub

Private Sub GridFiliaisFornecedores_GotFocus()
    Call Grid_Recebe_Foco(objGridFiliaisFornecedores)
End Sub

Private Sub GridFiliaisFornecedores_EnterCell()
    Call Grid_Entrada_Celula(objGridFiliaisFornecedores, iAlterado)
End Sub

Private Sub GridFiliaisFornecedores_LeaveCell()
    Call Saida_Celula(objGridFiliaisFornecedores)
End Sub

Private Sub GridFiliaisFornecedores_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridFiliaisFornecedores)
End Sub
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 74834

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 74834
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160702)

    End Select

    Exit Function

End Function

Private Sub GridFiliaisFornecedores_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFiliaisFornecedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFiliaisFornecedores, iAlterado)
    End If

End Sub

Private Sub GridFiliaisFornecedores_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridFiliaisFornecedores)
End Sub

Private Sub GridFiliaisFornecedores_RowColChange()
    Call Grid_RowColChange(objGridFiliaisFornecedores)
End Sub

Private Sub GridFiliaisFornecedores_Scroll()
    Call Grid_Scroll(objGridFiliaisFornecedores)
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim iPadrao As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 54191

    'Verifica preenchimento de produto
    If Len(Trim(Produto.ClipText)) = 0 Then Error 54192

    'Verifica preenchimento da filial
    If Len(Trim(Filial.Text)) = 0 Then Error 54194

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objFornecedorProdutoFF)
    If lErro <> SUCESSO Then Error 54288

    'Fornecedor padrão
    iPadrao = Padrao.Value
    
    lErro = Trata_Alteracao(objFornecedorProdutoFF, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn)
    If lErro <> SUCESSO Then Error 32290

    'Chama FornecedorProduto_Grava
    lErro = CF("FornecedorProdutoFF_Grava", objFornecedorProdutoFF, iPadrao)
    If lErro <> SUCESSO Then Error 54193

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err

        Case 32290

        Case 54191
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 54192
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 54193, 54288

        Case 54194
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160703)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Function Move_Tela_Memoria(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Move os dados da Tela para ObjFornecedorProdutoFF

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sNomeRed As String
Dim iCodigo As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Se a combo fornecedor está preenchida
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Se o Fornecedor for selecionado
        If Fornecedor.ListIndex > -1 Then
            objFornecedorProdutoFF.lFornecedor = Fornecedor.ItemData(Fornecedor.ListIndex)
        Else
            
            'Verifica se Fornecedor existe
            objFornecedor.sNomeReduzido = Fornecedor.Text
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 54195
    
            If lErro = 6681 Then Error 54197
    
            objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo
        End If
        
    End If

    'Se a combo filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objFornecedorProdutoFF.iFilialForn = Codigo_Extrai(Filial.Text)
    End If

    'Se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 54196

        objProduto.sCodigo = sProdutoFormatado

        objFornecedorProdutoFF.sProduto = objProduto.sCodigo

    End If

    'Preenche objFornecedorProdutoFF
    objFornecedorProdutoFF.sProdutoFornecedor = ProdutoFornecedor.Text

    'Verifica se os campos estão preenchidos

    'Frame Fornecedor
    objFornecedorProdutoFF.dLoteMinimo = StrParaDbl(LoteMinimo.Text)
    objFornecedorProdutoFF.iNota = StrParaInt(Nota.Text)

    'Se Padrão está marcado
    objFornecedorProdutoFF.iPadrao = Padrao.Value
   
    objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa

    'Maristela
    objFornecedorProdutoFF.sDescricao = DescFornProd.Text
    
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err
        
        Case 54195, 54196
        
        Case 54197
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160704)

    End Select

    Exit Function

End Function

Private Function Traz_FornecedorProdutoFF_Tela(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Traz os dados da FornecedorProdutoFF para a Tela

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Traz_FornecedorProdutoFF_Tela

    'Lê FornecedorProduto
    lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
    If lErro <> SUCESSO And lErro <> 54217 Then gError 28293

    objFornecedorProdutoFF.dtDataUltimaCotacao = DATA_NULA
    objFornecedorProdutoFF.dtDataPedido = DATA_NULA
    objFornecedorProdutoFF.dtDataReceb = DATA_NULA

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
    
    'Se não encontrou o FornecedorProduto
    If lErro = 54217 Then
        'Limpa os campos da tela
        Call Limpa_Campos_FornecedorProdutoFF
    Else
        'Mostra os dados de FornecedorProduto na Tela
        lErro = Mostra_FornecedorProdutoFF_Tela(objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 54200
    End If

    Traz_FornecedorProdutoFF_Tela = SUCESSO

    Exit Function

Erro_Traz_FornecedorProdutoFF_Tela:

    Traz_FornecedorProdutoFF_Tela = gErr

    Select Case gErr

        Case 28293, 54200, 89346, 89349, 89354, 89368, 89370

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160705)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FornecedorProdutoFF"

    'Lê os dados da Tela FornecedorProduto
    lErro = Move_Tela_Memoria(objFornecedorProdutoFF)
    If lErro <> SUCESSO Then Error 54201

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objFornecedorProdutoFF.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Fornecedor", objFornecedorProdutoFF.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "FilialForn", objFornecedorProdutoFF.iFilialForn, 0, "FilialForn"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 54201

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160706)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim bCancel As Boolean

On Error GoTo Erro_Tela_Preenche

    objFornecedorProdutoFF.sProduto = colCampoValor.Item("Produto").vValor
    objFornecedorProdutoFF.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    objFornecedorProdutoFF.iFilialForn = colCampoValor.Item("FilialForn").vValor
    objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa

    If (objFornecedorProdutoFF.lFornecedor <> 0) And (objFornecedorProdutoFF.sProduto <> "") Then

        'Mostra o produto na tela
        Produto.PromptInclude = False
        Produto.Text = objFornecedorProdutoFF.sProduto
        Produto.PromptInclude = True

        'Dispara o Validate de Produto
        Call Produto_Validate(bCancel)

        'Mostra Fornecedor na Tela
        Fornecedor.Text = objFornecedorProdutoFF.lFornecedor

        'Dispara o Validate de fornecedor
        Call Fornecedor_Validate(bCancel)

        'Mostra a Filial na tela
        Filial.Text = objFornecedorProdutoFF.iFilialForn

        'Dispara o Validate de filial
        Call Filial_Validate(bCancel)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160707)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Campos_FornecedorProdutoFF()

    Padrao.Value = 0

    'Limpa campos da tela

    'Frame Fornecedor
    ProdutoFornecedor.Text = ""
    Nota.Text = ""
    LoteMinimo.Text = ""

    'Frame Estatísticas
    
    'Pedidos de Compra
    QuantPedAbertos.Caption = ""
    TempoRessup.Caption = ""

    'Última Compra
    QuantPedida.Caption = ""
    QuantRecebida.Caption = ""
    PrecoTotal.Caption = ""
    DataUltimaCompra.Caption = ""
    DataPedido.Caption = ""
    DataReceb.Caption = ""

    'Última Cotação
    QuantUltimaCotacao.Caption = ""
    UltimaCotacao.Caption = ""
    DataUltimaCotacao.Caption = ""
    CondPagtoUltimaCotacao.Caption = ""
    TipoFrete.Caption = ""
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Produto X Fornecedor"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FornFilialProduto"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call LabelForn_Click
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

Private Sub UMCompra_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(UMCompra(Index), Source, X, Y)
End Sub

Private Sub UMCompra_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UMCompra(Index), Button, Shift, X, Y)
End Sub


Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub



Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub PrecoTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PrecoTotal, Source, X, Y)
End Sub

Private Sub PrecoTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PrecoTotal, Button, Shift, X, Y)
End Sub

Private Sub DataUltimaCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltimaCompra, Source, X, Y)
End Sub

Private Sub DataUltimaCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltimaCompra, Button, Shift, X, Y)
End Sub


Private Sub QuantPedAbertos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantPedAbertos, Source, X, Y)
End Sub

Private Sub QuantPedAbertos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantPedAbertos, Button, Shift, X, Y)
End Sub

Private Sub TempoRessup_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TempoRessup, Source, X, Y)
End Sub

Private Sub TempoRessup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TempoRessup, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub DataUltimaCotacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltimaCotacao, Source, X, Y)
End Sub

Private Sub DataUltimaCotacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltimaCotacao, Button, Shift, X, Y)
End Sub

Private Sub UltimaCotacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UltimaCotacao, Source, X, Y)
End Sub

Private Sub UltimaCotacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UltimaCotacao, Button, Shift, X, Y)
End Sub

Private Sub QuantUltimaCotacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantUltimaCotacao, Source, X, Y)
End Sub

Private Sub QuantUltimaCotacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantUltimaCotacao, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoUltimaCotacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoUltimaCotacao, Source, X, Y)
End Sub

Private Sub CondPagtoUltimaCotacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoUltimaCotacao, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub DataReceb_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataReceb, Source, X, Y)
End Sub

Private Sub DataReceb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataReceb, Button, Shift, X, Y)
End Sub

Private Sub QuantPedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantPedida, Source, X, Y)
End Sub

Private Sub QuantPedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantPedida, Button, Shift, X, Y)
End Sub

Private Sub QuantRecebida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantRecebida, Source, X, Y)
End Sub

Private Sub QuantRecebida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantRecebida, Button, Shift, X, Y)
End Sub

Private Sub DataPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPedido, Source, X, Y)
End Sub

Private Sub DataPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPedido, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub


Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label22(Index), Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22(Index), Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub TipoFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoFrete, Source, X, Y)
End Sub

Private Sub TipoFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoFrete, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Prod_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Prod(Index), Source, X, Y)
End Sub

Private Sub Prod_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Prod(Index), Button, Shift, X, Y)
End Sub

Private Sub DescProd_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProd, Source, X, Y)
End Sub

Private Sub DescProd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProd, Button, Shift, X, Y)
End Sub

Private Sub LabelFornFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornFilial, Source, X, Y)
End Sub

Private Sub LabelFornFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornFilial, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedor, Source, X, Y)
End Sub

Private Sub LabelFornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedor, Button, Shift, X, Y)
End Sub

