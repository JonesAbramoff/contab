VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ReqComprasOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8160
      Index           =   1
      Left            =   150
      TabIndex        =   26
      Top             =   810
      Width           =   16650
      Begin VB.Frame Frame6 
         Caption         =   "Local de Entrega"
         Height          =   1065
         Left            =   90
         TabIndex        =   58
         Top             =   5475
         Width           =   11145
         Begin VB.Frame Frame2 
            Caption         =   "Tipo"
            Height          =   600
            Left            =   225
            TabIndex        =   64
            Top             =   240
            Width           =   3330
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Filial Empresa"
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
               Left            =   120
               TabIndex        =   11
               Top             =   270
               Value           =   -1  'True
               Width           =   1515
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
               Left            =   1770
               TabIndex        =   12
               Top             =   270
               Width           =   1335
            End
         End
         Begin VB.Frame FrameTipoDestino 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   750
            Index           =   0
            Left            =   4305
            TabIndex        =   62
            Top             =   210
            Width           =   3645
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               Left            =   1125
               TabIndex        =   13
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
               Left            =   600
               TabIndex        =   63
               Top             =   255
               Width           =   465
            End
         End
         Begin VB.Frame FrameTipoDestino 
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   4335
            TabIndex        =   59
            Top             =   225
            Visible         =   0   'False
            Width           =   3645
            Begin VB.ComboBox FilialFornecedor 
               Height          =   315
               Left            =   1110
               TabIndex        =   41
               Top             =   360
               Width           =   2160
            End
            Begin MSMask.MaskEdBox Fornecedor 
               Height          =   300
               Left            =   1110
               TabIndex        =   40
               Top             =   0
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
               Left            =   585
               TabIndex        =   61
               Top             =   405
               Width           =   465
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
               Left            =   15
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   60
               Top             =   60
               Width           =   1035
            End
         End
      End
      Begin VB.CommandButton BotaoModelos 
         Caption         =   "Requisições Modelo"
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
         Left            =   90
         TabIndex        =   14
         Top             =   7485
         Width           =   1995
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cabeçalho"
         Height          =   4035
         Left            =   90
         TabIndex        =   27
         Top             =   30
         Width           =   11145
         Begin VB.TextBox Observacao 
            Height          =   1515
            Left            =   1320
            MaxLength       =   255
            TabIndex        =   8
            Top             =   2310
            Width           =   9330
         End
         Begin VB.ComboBox FilialCompra 
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   765
            Width           =   2295
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
            Left            =   8670
            TabIndex        =   6
            Top             =   1320
            Width           =   2100
         End
         Begin VB.ComboBox TipoTributacao 
            Height          =   315
            Left            =   6525
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   780
            Width           =   4125
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2115
            Picture         =   "ReqComprasOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Numeração Automática"
            Top             =   285
            Width           =   300
         End
         Begin VB.CheckBox Urgente 
            Caption         =   "Urgente"
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
            Left            =   6525
            TabIndex        =   5
            Top             =   1320
            Width           =   1005
         End
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   300
            Left            =   6525
            TabIndex        =   1
            Top             =   300
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoOP 
            Height          =   315
            Left            =   1320
            TabIndex        =   7
            Top             =   1770
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   270
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   1260
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label CodigoPV 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6525
            TabIndex        =   75
            Top             =   1770
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Código PV:"
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
            Left            =   5490
            TabIndex        =   74
            Top             =   1830
            Width           =   960
         End
         Begin VB.Label CodigoOPLabel 
            AutoSize        =   -1  'True
            Caption         =   "Código OP:"
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
            TabIndex        =   33
            Top             =   1830
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Left            =   105
            TabIndex        =   30
            Top             =   825
            Width           =   1155
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tributação:"
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
            Left            =   4755
            TabIndex        =   31
            Top             =   810
            Width           =   1695
         End
         Begin VB.Label CclLabel 
            AutoSize        =   -1  'True
            Caption         =   "C.Custo:"
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
            Left            =   525
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label ObservacaoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
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
            TabIndex        =   34
            Top             =   2370
            Width           =   1095
         End
         Begin VB.Label RequisitanteLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   5310
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   345
            Width           =   1140
         End
         Begin VB.Label CodigoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nº Req:"
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
            Left            =   570
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   330
            Width           =   690
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datas"
         Height          =   780
         Left            =   75
         TabIndex        =   35
         Top             =   4485
         Width           =   11145
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   2925
            TabIndex        =   37
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1815
            TabIndex        =   9
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEntrega 
            Height          =   300
            Left            =   6735
            TabIndex        =   39
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   300
            Left            =   5640
            TabIndex        =   10
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
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
            Left            =   1305
            TabIndex        =   36
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Limite:"
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
            Left            =   4560
            TabIndex        =   38
            Top             =   330
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   8235
      Index           =   3
      Left            =   150
      TabIndex        =   65
      Top             =   810
      Visible         =   0   'False
      Width           =   16650
      Begin MSMask.MaskEdBox QuantReqPV 
         Height          =   315
         Left            =   5940
         TabIndex        =   73
         Top             =   1800
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
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
      Begin VB.TextBox DescProdPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   2235
         MaxLength       =   50
         TabIndex        =   72
         Top             =   2310
         Width           =   4000
      End
      Begin MSMask.MaskEdBox CodPV 
         Height          =   315
         Left            =   4500
         TabIndex        =   71
         Top             =   1335
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ItemReqPV 
         Height          =   315
         Left            =   375
         TabIndex        =   70
         Top             =   2310
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.TextBox UMPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   5415
         MaxLength       =   50
         TabIndex        =   69
         Top             =   1785
         Width           =   435
      End
      Begin MSMask.MaskEdBox QuantPV 
         Height          =   315
         Left            =   6870
         TabIndex        =   67
         Top             =   1815
         Width           =   945
         _ExtentX        =   1667
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
      Begin MSMask.MaskEdBox ProdutoPV 
         Height          =   315
         Left            =   1080
         TabIndex        =   68
         Top             =   1800
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridPV 
         Height          =   8040
         Left            =   165
         TabIndex        =   66
         Top             =   120
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   14182
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8235
      Index           =   2
      Left            =   120
      TabIndex        =   42
      Top             =   780
      Visible         =   0   'False
      Width           =   16695
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
         Height          =   285
         Left            =   3150
         TabIndex        =   22
         Top             =   7830
         Width           =   1710
      End
      Begin VB.CommandButton BotaoCcl 
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
         Height          =   285
         Left            =   1335
         TabIndex        =   21
         Top             =   7830
         Width           =   1710
      End
      Begin VB.CommandButton BotaoFiliaisFornProd 
         Caption         =   "Fornecedores do Produto"
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
         Left            =   6510
         TabIndex        =   24
         Top             =   7830
         Width           =   2370
      End
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
         Height          =   285
         Left            =   4935
         TabIndex        =   23
         Top             =   7830
         Width           =   1485
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
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   7830
         Width           =   1005
      End
      Begin VB.Frame Frame4 
         Caption         =   "Itens"
         Height          =   7635
         Left            =   90
         TabIndex        =   43
         Top             =   120
         Width           =   16455
         Begin VB.ComboBox Exclusivo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "ReqComprasOcx.ctx":00EA
            Left            =   5715
            List            =   "ReqComprasOcx.ctx":00F4
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   3120
            Width           =   1305
         End
         Begin VB.ComboBox TipoTribItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   3045
            Width           =   2520
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2685
            MaxLength       =   50
            TabIndex        =   46
            Top             =   2475
            Width           =   4000
         End
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   315
            Width           =   885
         End
         Begin VB.TextBox ObservacaoGrid 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6210
            MaxLength       =   255
            TabIndex        =   56
            Top             =   3195
            Width           =   2355
         End
         Begin VB.ComboBox FilialFornecGrid 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "ReqComprasOcx.ctx":0111
            Left            =   3900
            List            =   "ReqComprasOcx.ctx":0113
            TabIndex        =   54
            Top             =   3090
            Width           =   1770
         End
         Begin MSMask.MaskEdBox CentroCusto 
            Height          =   225
            Left            =   6210
            TabIndex        =   50
            Top             =   435
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   5100
            TabIndex        =   49
            Top             =   420
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
            Left            =   4065
            TabIndex        =   48
            Top             =   375
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
            Left            =   225
            TabIndex        =   45
            Top             =   390
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecGrid 
            Height          =   225
            Left            =   1890
            TabIndex        =   53
            Top             =   3270
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
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   240
            Left            =   7170
            TabIndex        =   51
            Top             =   450
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   423
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   3375
            Left            =   210
            TabIndex        =   44
            Top             =   270
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   14250
      ScaleHeight     =   495
      ScaleWidth      =   2625
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   30
      Width           =   2685
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   622
         Picture         =   "ReqComprasOcx.ctx":0115
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1626
         Picture         =   "ReqComprasOcx.ctx":026F
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2130
         Picture         =   "ReqComprasOcx.ctx":07A1
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   120
         Picture         =   "ReqComprasOcx.ctx":091F
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1124
         Picture         =   "ReqComprasOcx.ctx":0A21
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8670
      Left            =   75
      TabIndex        =   25
      Top             =   450
      Width           =   16830
      _ExtentX        =   29686
      _ExtentY        =   15293
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos de Venda"
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
Attribute VB_Name = "ReqComprasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'EVENTOS DOS BROWSERS
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoRequisitante As AdmEvento
Attribute objEventoRequisitante.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoObservacao As AdmEvento
Attribute objEventoObservacao.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoBotaoCcl As AdmEvento
Attribute objEventoBotaoCcl.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxarifados As AdmEvento
Attribute objEventoAlmoxarifados.VB_VarHelpID = -1
Private WithEvents objEventoFiliaisFornProduto As AdmEvento
Attribute objEventoFiliaisFornProduto.VB_VarHelpID = -1
Private WithEvents objEventoReqModelo As AdmEvento
Attribute objEventoReqModelo.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoOPCodigo As AdmEvento
Attribute objEventoOPCodigo.VB_VarHelpID = -1
    

'Variáveis Globais
Dim gColItemReqCompras As Collection
Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim iFrameAtual As Integer
Dim giTipoTributacao As Integer

'GridItens
Dim objGridItens As AdmGrid
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_CCL_Col As Integer
Dim iGrid_ContaContabil_Col As Integer
Dim iGrid_TipoTributacao_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialFornecedor_Col As Integer
Dim iGrid_Exclusivo_Col As Integer
Dim iGrid_Observacao_Col As Integer

'GridPV
Dim objGridPV As AdmGrid
Dim iGrid_ItemReqPV_Col As Integer
Dim iGrid_ProdutoPV_Col As Integer
Dim iGrid_DescProdPV_Col As Integer
Dim iGrid_CodPV_Col As Integer
Dim iGrid_UMPV_Col As Integer
Dim iGrid_QuantReqPV_Col As Integer
Dim iGrid_QuantPV_Col As Integer

'workflow
Private Const CodigoRequisicao As String = "CodigoRequisicao"
Private Const FILIAL_EMPRESA As String = "FilialEmpresa"

Function Trata_Parametros(Optional objRequisicaoCompras As ClassRequisicaoCompras) As Long
    
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma Requisição foi passada por parâmetro
    If Not (objRequisicaoCompras Is Nothing) Then

        'Se o número interno estiver preenchido
        If objRequisicaoCompras.lNumIntDoc > 0 Then

            'Lê a Requisição Compras a partir de seu número interno
            lErro = CF("RequisicaoCompras_Le", objRequisicaoCompras)
            If lErro = SUCESSO Then

                'Traz os dados da Requisição para a tela
                lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompras)
                If lErro <> SUCESSO Then Error 61835
            
            End If
            
        'Se o NumIntDoc não estiver preenchido
        ElseIf objRequisicaoCompras.lCodigo <> 0 Then
                
            objRequisicaoCompras.iFilialEmpresa = giFilialEmpresa
            
            'Lê a Requisição Compras a partir do código
            lErro = CF("RequisicaoCompras_Le_Numero", objRequisicaoCompras)
            If lErro <> SUCESSO And lErro <> 54439 Then Error 61836
            
            'Se não encontrou -> Erro
            If lErro = 54439 Then
                Codigo.PromptInclude = False
                Codigo.Text = objRequisicaoCompras.lCodigo
                Codigo.PromptInclude = True
            Else
                'Traz os dados da Requisição para a tela
                lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompras)
                If lErro <> SUCESSO Then Error 61838
            End If
            
        End If

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 61835, 61836, 61838
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173861)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objReqCompras As New ClassRequisicaoCompras
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 76015
    If Len(Trim(Requisitante.Text)) = 0 Then gError 76016
    If Len(Trim(Data.ClipText)) = 0 Then gError 76017
    
    objReqCompras.lCodigo = StrParaLong(Codigo.Text)
    objReqCompras.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se a Requisicao de Compra informada existe
    lErro = CF("RequisicaoCompras_Le_Numero", objReqCompras)
    If lErro <> SUCESSO And lErro <> 54439 Then gError 76018
    
    'Se a Requisicao não existe ==> erro
    If lErro = 54439 Then gError 76019
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Requisições de Compra", "REQCOMTO.NumIntDoc = @NREQCOM", 1, "REQCOM", "NREQCOM", objReqCompras.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76020
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 76015
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 76016
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_PREENCHIDO", gErr)
            
        Case 76017
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            
        Case 76018, 76020
        
        Case 76019
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objReqCompras.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173862)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 65557
    
    'Limpa a Tela
    Call Limpa_Tela_RequisicaoCompra
   
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
   
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 65557
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173863)

    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_RequisicaoCompra()

Dim iIndice As Integer
Dim bCancel As Boolean
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim lErro As Long, iFilialEmpresa As Integer

On Error GoTo Erro_Limpa_Tela_RequisicaoCompra

    'Função genérica que limpa a tela
    Call Limpa_Tela(Me)
    
    Set gColItemReqCompras = New Collection
    
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
        
    DataLimite.PromptInclude = False
    DataLimite.Text = ""
    DataLimite.PromptInclude = True
    
    'Limpa o GridItens
    Call Grid_Limpa(objGridItens)
    
    'Limpa o GridPV
    Call Grid_Limpa(objGridPV)
    
    Urgente.Value = vbUnchecked
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True
    
    'Limpa as combos
    FilialFornecedor.Clear
    
    'Coloca Filiais default para FilialCompra e FilialEmpresa
    iFilialEmpresa = giFilialEmpresa
    'Coloca FiliaEmpresa Default na Tela
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126951
    
    FilialEmpresa.Text = iFilialEmpresa

    Call FilialEmpresa_Validate(bCancel)
    
    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then gError (65564)

    'Coloca FilialCompra Default na tela
    If objConfiguraCOM.iFilialCompra > 0 Then
        FilialCompra.Text = objConfiguraCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If
    Call FilialCompra_Validate(bCancel)
        
    'Coloca Tipo de tributação Default
    For iIndice = 0 To TipoTributacao.ListCount - 1
        If Codigo_Extrai(TipoTributacao.List(iIndice)) = giTipoTributacao Then
            TipoTributacao.ListIndex = iIndice
            Exit For
        End If
    Next
        
    EnviaCompras.Value = vbUnchecked
    
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    CodigoPV.Caption = ""
    
    Exit Sub
    
Erro_Limpa_Tela_RequisicaoCompra:

    Select Case gErr
        
        Case 65564
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173864)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera o próximo código para requisição.
    lErro = CF("RequisicaoCompras_Codigo_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 61839

    'Coloca o código gerado na tela.
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 61839
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173865)
    
    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim sMascaraCclPadrao As String
Dim bCancel As Boolean
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    'Inicializa os ObjEventos
    Set objEventoCodigo = New AdmEvento
    Set objEventoRequisitante = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoObservacao = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoBotaoCcl = New AdmEvento
    Set objEventoAlmoxarifados = New AdmEvento
    Set objEventoFiliaisFornProduto = New AdmEvento
    Set objEventoReqModelo = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoOPCodigo = New AdmEvento
    
    'Inicializa coleção de Itens de Requisição
    Set gColItemReqCompras = New Collection
    
    'Data Atual
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'Atualiza a variável global para controle de frames e seta um tipo Padrao
    iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True

    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 61841

    'Preenche a Combo FilialCompra com as filiais Empresas
    lErro = Carrega_ComboFiliais(colCodigoDescricao)
    If lErro <> SUCESSO Then gError 61842

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 61843

    'Inicializa Máscara de CentroCusto e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 61844

    Ccl.Mask = sMascaraCclPadrao
    CentroCusto.Mask = sMascaraCclPadrao

    'Inicializa mascara de ContaContabil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 61845

    'Inicializa o GridItens
    Set objGridItens = New AdmGrid

    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 61846

    Set objGridPV = New AdmGrid

    lErro = Inicializa_GridPV(objGridPV)
    If lErro <> SUCESSO Then gError 178595

    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then gError 61847

    'Coloca FilialCompra Default na tela
    If objConfiguraCOM.iFilialCompra > 0 Then
        FilialCompra.Text = objConfiguraCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If
    Call FilialCompra_Validate(bCancel)
    
    iFilialEmpresa = giFilialEmpresa
    'Coloca FiliaEmpresa Default na Tela
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126951
    
    FilialEmpresa.Text = iFilialEmpresa

    Call FilialEmpresa_Validate(bCancel)
    'FilialEmpresa.ListIndex = 0
    
    'Carrega Tipos de Tributação
    lErro = Carrega_TipoTributacao()
    If lErro <> SUCESSO Then gError 66122
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_LIGHT Then
        
        FilialCompra.left = POSICAO_FORA_TELA
        FilialCompra.TabStop = False
        Label3.left = POSICAO_FORA_TELA
        Label3.Visible = False
        FilialEmpresa.left = POSICAO_FORA_TELA
        FilialEmpresa.TabStop = False
        Label6.left = POSICAO_FORA_TELA
        Label6.Visible = False
        
    End If

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 61841, 61842, 61843, 61844, 61845, 61846, 61847, 66122, 126951, 178595

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173866)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Carrega_TipoTributacao() As Long
'Carrega Tipos de Tributação

Dim lErro As Long
Dim colTributacao As New AdmColCodigoNome
Dim iIndice As Integer
Dim iTipoTrib As Integer

On Error GoTo Erro_Carrega_TipoTributacao

    'Lê os Tipos de Tributação associadas a Compras
    lErro = CF("TiposTributacaoCompras_Le", colTributacao)
    If lErro <> SUCESSO Then gError 66123
        
    'Lê o Tipo de Tributação Padrão
    lErro = CF("TipoTributacaoPadrao_Le", iTipoTrib)
    If lErro <> SUCESSO And lErro <> 66597 Then gError 66124
    
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
    
    giTipoTributacao = iTipoTrib
    
    Carrega_TipoTributacao = SUCESSO
    
    Exit Function
    
Erro_Carrega_TipoTributacao:

    Carrega_TipoTributacao = gErr
    
    Select Case gErr
        
        Case 66123, 66124
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173867)
        
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Finaliza os objEventos
    Set objEventoCodigo = Nothing
    Set objEventoRequisitante = Nothing
    Set objEventoCcl = Nothing
    Set objEventoObservacao = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoProduto = Nothing
    Set objEventoBotaoCcl = Nothing
    Set objEventoAlmoxarifados = Nothing
    Set objEventoFiliaisFornProduto = Nothing
    Set objEventoReqModelo = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoOPCodigo = Nothing
    
    'Libera variáveis globais
    Set gColItemReqCompras = Nothing
    Set objGridItens = Nothing
    Set objGridPV = Nothing
    
    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    'Se nenhuma linha do Grid foi selecionada, erro
    If GridItens.Row = 0 Then gError 66589

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then gError 66590

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 66591

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 66589
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 66590
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 66591

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173868)

    End Select

    Exit Sub

End Sub

Private Sub CodigoOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemProducao As New ClassOrdemDeProducao

On Error GoTo Erro_CodigoOP_Validate
        
    'Se o Código OP não foi preenchido, sai da rotina
    If Len(Trim(CodigoOP.Text)) = 0 Then Exit Sub
    
    objOrdemProducao.sCodigo = CodigoOP.Text
    objOrdemProducao.iFilialEmpresa = giFilialEmpresa
    
    'Lê a Ordem de Produção não baixada a partir do código passado
    lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemProducao)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 67300
    
    'Se não encontrou, erro
    If lErro = 21960 Then gError 67301
    
    Exit Sub
    
Erro_CodigoOP_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 67300
        
        Case 67301
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", gErr, objOrdemProducao.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173869)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub CodigoOPLabel_Click()

Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(CodigoOP.Text)) <> 0 Then objOrdemDeProducao.sCodigo = CodigoOP.Text
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemDeProducao, objEventoOPCodigo)

End Sub

Private Sub DescProduto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoOPCodigo_evSelecao(obj1 As Object)

Dim objOrdemDeProducao As ClassOrdemDeProducao

On Error GoTo Erro_objEventoOPCodigo_evSelecao

    Set objOrdemDeProducao = obj1

    CodigoOP.Text = objOrdemDeProducao.sCodigo
    
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoOPCodigo_evSelecao:

    Select Case gErr

        Case 67299

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173870)
    
    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If Len(Trim(objPlanoConta.sConta)) > 0 Then

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 66588
        
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
        
        GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 66588
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173871)

    End Select

    Exit Sub

End Sub

Function Carrega_ComboFiliais(colCodigoDescricao As AdmColCodigoNome) As Long
'Carrega as Combos (FilialEmpresa e FilialCompra com as Filiais Empresa passada na colecao

Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Carrega_ComboFiliais

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

    Carrega_ComboFiliais = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173872)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil")
    objGridInt.colColuna.Add ("Tipo Tributação")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusividade")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)
    objGridInt.colCampo.Add (TipoTribItem.Name)
    objGridInt.colCampo.Add (FornecGrid.Name)
    objGridInt.colCampo.Add (FilialFornecGrid.Name)
    objGridInt.colCampo.Add (Exclusivo.Name)
    objGridInt.colCampo.Add (ObservacaoGrid.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UM_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_Almoxarifado_Col = 5
    iGrid_CCL_Col = 6
    iGrid_ContaContabil_Col = 7
    iGrid_TipoTributacao_Col = 8
    iGrid_Fornecedor_Col = 9
    iGrid_FilialFornecedor_Col = 10
    iGrid_Exclusivo_Col = 11
    iGrid_Observacao_Col = 12

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 20

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)
'Rotina que habilita a entrada na celula

Dim lErro As Long
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim objProduto As New ClassProduto
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim iIndice As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor
Dim iCodigo As Integer
Dim sTipoTrib As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 61848

    'Passa o produto controle
    Select Case objControl.Name
    
        Case Produto.Name

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
               objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case UM.Name

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
                objControl.Enabled = True
    
                objProduto.sCodigo = sProdutoFormatado
    
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 61849
    
                If lErro = 28030 Then gError 61851
    
                objClasseUM.iClasse = objProduto.iClasseUM
    
                'Preenche a List da Combo UnidadeMed com as UM's do Produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 61850
    
                'Guardo o valor da Unidade de Medida da Linha
                sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col)
    
                'Limpar as Unidades utilizadas anteriormente
                UM.Clear
    
                For Each objUnidadeDeMedida In colSiglas
                    UM.AddItem objUnidadeDeMedida.sSigla
                Next
    
                'Tento selecionar na Combo a Unidade anterior
                If UM.ListCount <> 0 Then
    
                    For iIndice = 0 To UM.ListCount - 1
                        If UM.List(iIndice) = sUnidadeMed Then
                            UM.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
            Else
                objControl.Enabled = False
            End If
            
        Case Almoxarifado.Name, CentroCusto.Name

            'Verifica se o detino é a empresa
            If iFrameTipoDestinoAtual <> TIPO_DESTINO_EMPRESA Or Len(Trim(FilialEmpresa.Text)) = 0 Then
                
                Almoxarifado.Text = ""
                CentroCusto.PromptInclude = False
                CentroCusto.Text = ""
                CentroCusto.PromptInclude = True
                
                objControl.Enabled = False
            Else

                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If

            End If

        Case Quantidade.Name, ContaContabil.Name, FornecGrid.Name, ObservacaoGrid.Name ', DescProduto.Name
    
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        Case DescProduto.Name
    
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And gobjCOM.iReqCompraDescProdBloq <> MARCADO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case Exclusivo.Name
    
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case TipoTribItem.Name
        
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
                objControl.Enabled = True
                
                'Guardo o valor do Tipo de tributação
                sTipoTrib = GridItens.TextMatrix(GridItens.Row, iGrid_TipoTributacao_Col)
    
                'Limpar os Tipos de tributação
                TipoTribItem.Clear
    
                For iIndice = 0 To TipoTributacao.ListCount - 1
                    TipoTribItem.AddItem TipoTributacao.List(iIndice)
                Next
    
                'Tento selecionar na Combo o Tipo anterior
                If TipoTribItem.ListCount <> 0 Then
    
                    For iIndice = 0 To TipoTribItem.ListCount - 1
                        If TipoTribItem.List(iIndice) = sTipoTrib Then
                            TipoTribItem.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
            
            Else
                objControl.Enabled = False
            End If
    
        Case FilialFornecGrid.Name
    
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                
                objControl.Enabled = True
                
                'Se o Fornecedor não está preenchido
                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col))) = 0 Then
                    
                    'Desabilita combo de Filiais
                    objControl.Enabled = False
                
                Else
                    
                    objFornecedor.sNomeReduzido = GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col)
                    
                    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                    If lErro <> SUCESSO And lErro <> 6681 Then gError (65590)
                    If lErro = 6681 Then gError (65591)
                    
                    lErro = CF("FornecedorProdutoFF_Le_FilialForn", sProdutoFormatado, objFornecedor.lCodigo, Codigo_Extrai(FilialCompra.Text), colCodigoNome)
                    If lErro <> SUCESSO Then gError (66592)
                    
                    If colCodigoNome.Count = 0 Then gError (65593)
                
                    If Len(Trim(FilialFornecGrid.Text)) = 0 Then
                        iCodigo = colCodigoNome.Item(1).iCodigo
                    Else
                        iCodigo = Codigo_Extrai(FilialFornecGrid.Text)
                    End If
    
                    FilialFornecGrid.Clear
                    
                    Call CF("Filial_Preenche", FilialFornecGrid, colCodigoNome)
                    Call CF("Filial_Seleciona", FilialFornecGrid, iCodigo)
                
                End If
            End If
    
        Case ItemReqPV.Name
            If Len(Trim(GridPV.TextMatrix(GridPV.Row, iGrid_ItemReqPV_Col))) = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        Case CodPV.Name, QuantPV.Name
            If Len(Trim(GridPV.TextMatrix(GridPV.Row, iGrid_ItemReqPV_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 61848, 61849, 61850, 65590, 66592

        Case 61851
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 65591
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case 65593
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR", gErr, objFornecedor.sNomeReduzido, sProdutoFormatado)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173873)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        If objGridInt Is objGridItens Then

            'Verifica qual a coluna atual do Grid
            Select Case objGridInt.objGrid.Col
    
                'Produto
                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 61852
                
                Case iGrid_Descricao_Col
                    lErro = Saida_Celula_Descricao(objGridInt)
                    If lErro <> SUCESSO Then gError 86176
                
                'Unidade de Medida
                Case iGrid_UM_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 61853
                
                'Quantidade
                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 61854
                
                'Almoxarifado
                Case iGrid_Almoxarifado_Col
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 61855
                
                'Ccl
                Case iGrid_CCL_Col
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 61856
                        
                'ContaContabil
                Case iGrid_ContaContabil_Col
                    lErro = Saida_Celula_ContaContabil(objGridInt)
                    If lErro <> SUCESSO Then gError 61857
    
                'Fornecedor
                Case iGrid_Fornecedor_Col
                    lErro = Saida_Celula_Fornecedor(objGridInt)
                    If lErro <> SUCESSO Then gError 61858
                    
                'Filial Fornecedor
                Case iGrid_FilialFornecedor_Col
                    lErro = Saida_Celula_FilialForn(objGridInt)
                    If lErro <> SUCESSO Then gError 61859
                    
                'Exclusivo
                Case iGrid_Exclusivo_Col
                    lErro = Saida_Celula_Exclusivo(objGridInt)
                    If lErro <> SUCESSO Then gError 61860
                
                'Observação
                Case iGrid_Observacao_Col
                    lErro = Saida_Celula_Observacao(objGridInt)
                    If lErro <> SUCESSO Then gError 61861
            
                'Tipo de Tributação
                Case iGrid_TipoTributacao_Col
                    lErro = Saida_Celula_TipoTributacao(objGridInt)
                    If lErro <> SUCESSO Then gError 66125
            
            End Select

        ElseIf objGridInt Is objGridPV Then
        
            'Verifica qual a coluna atual do Grid
            Select Case objGridInt.objGrid.Col
    
                'Produto
                Case iGrid_ItemReqPV_Col
                    lErro = Saida_Celula_ItemReqPV(objGridInt)
                    If lErro <> SUCESSO Then gError 178596
                
                'Pedido de Venda
                Case iGrid_CodPV_Col
                    lErro = Saida_Celula_CodPV(objGridInt)
                    If lErro <> SUCESSO Then gError 178597
                
                'Quantidade
                Case iGrid_QuantPV_Col
                    lErro = Saida_Celula_QuantPV(objGridInt)
                    If lErro <> SUCESSO Then gError 178598
                
            End Select
        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 61862

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 61852 To 61862, 66125, 86176, 178596 To 178598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173874)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AO SISTEMA DE BROWSE "'
'""""""""""""""""""""""""""""""""""""""""""""""
'SISTEMA DE SETAS
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objRequisicaoCompra As New ClassRequisicaoCompras
Dim sNomeRed As String

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ReqCompras_Requisitante"
    
    'Move todos os dados Presentes da Tela para objRequisicaoCompra
    lErro = Move_Tela_Memoria(objRequisicaoCompra)
    If lErro <> SUCESSO Then Error 61863

    sNomeRed = Requisitante.Text
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objRequisicaoCompra.lCodigo, 0, "Codigo"
    colCampoValor.Add "Data", objRequisicaoCompra.dtData, 0, "Data"
    colCampoValor.Add "DataEnvio", objRequisicaoCompra.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataLimite", objRequisicaoCompra.dtDataLimite, 0, "DataLimite"
    colCampoValor.Add "Ccl", objRequisicaoCompra.sCcl, STRING_BUFFER_MAX_TEXTO, "Ccl"
    colCampoValor.Add "Requisitante", objRequisicaoCompra.lRequisitante, 0, "Requisitante"
    colCampoValor.Add "FilialCompra", objRequisicaoCompra.iFilialCompra, 0, "FilialCompra"
    colCampoValor.Add "Observacao", objRequisicaoCompra.lObservacao, 0, "Observacao"
    colCampoValor.Add "TipoDestino", objRequisicaoCompra.iTipoDestino, 0, "TipoDestino"
    colCampoValor.Add "FornCliDestino", objRequisicaoCompra.lFornCliDestino, 0, "FornCliDestino"
    colCampoValor.Add "FilialDestino", objRequisicaoCompra.iFilialDestino, 0, "FilialDestino"
    colCampoValor.Add "Urgente", objRequisicaoCompra.lUrgente, 0, "Urgente"
    colCampoValor.Add "TipoTributacao", objRequisicaoCompra.iTipoTributacao, 0, "TipoTributacao"
    colCampoValor.Add "OPCodigo", objRequisicaoCompra.sOPCodigo, STRING_BUFFER_MAX_TEXTO, "OPCodigo"
''    colCampoValor.Add "NomeReduzido", sNomeRed, STRING_BUFFER_MAX_TEXTO, "NomeReduzido"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "DataEnvio", OP_IGUAL, DATA_NULA
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 61863

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173875)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objRequisicaoCompra As New ClassRequisicaoCompras

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objRequisicaoCompra.lCodigo = colCampoValor.Item("Codigo").vValor
    objRequisicaoCompra.dtData = colCampoValor.Item("Data").vValor
    objRequisicaoCompra.dtDataEnvio = colCampoValor.Item("DataEnvio").vValor
    objRequisicaoCompra.dtDataLimite = colCampoValor.Item("DataLimite").vValor
    objRequisicaoCompra.sCcl = colCampoValor.Item("Ccl").vValor
    objRequisicaoCompra.lRequisitante = colCampoValor.Item("Requisitante").vValor
    objRequisicaoCompra.iFilialCompra = colCampoValor.Item("FilialCompra").vValor
    objRequisicaoCompra.lObservacao = colCampoValor.Item("Observacao").vValor
    objRequisicaoCompra.iTipoDestino = colCampoValor.Item("TipoDestino").vValor
    objRequisicaoCompra.lFornCliDestino = colCampoValor.Item("FornCliDestino").vValor
    objRequisicaoCompra.iFilialDestino = colCampoValor.Item("FilialDestino").vValor
    objRequisicaoCompra.lUrgente = colCampoValor.Item("Urgente").vValor
    objRequisicaoCompra.iTipoTributacao = colCampoValor.Item("TipoTributacao").vValor
    objRequisicaoCompra.sOPCodigo = colCampoValor.Item("OPCodigo").vValor
    objRequisicaoCompra.iFilialEmpresa = giFilialEmpresa
    
    'Traz os dados da Requisição de Compras para tela
    lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompra)
    If lErro <> SUCESSO Then Error 61864

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 61864

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173876)

    End Select

    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim colSelecao As New Collection
Dim objRequisicaoCompra As New ClassRequisicaoCompras

    'Verifica se o código da Requisição foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then objRequisicaoCompra.lCodigo = CLng(Codigo.Text)
    
    Call Chama_Tela("ReqComprasNaoEnvLista", colSelecao, objRequisicaoCompra, objEventoCodigo)

End Sub

Private Sub FilialFornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisicaoCompra As New ClassRequisicaoCompras

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRequisicaoCompra = obj1

    'Traz dados da Requisição de Compras para a tela
    lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompra)
    If lErro <> SUCESSO Then Error 61865

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 61865

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173877)

    End Select

    Exit Sub

End Sub

Private Sub RequisitanteLabel_Click()

Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

    'Se o Requisitante estiver preenchido
    If Len(Trim(Requisitante.Text)) > 0 Then objRequisitante.sNomeReduzido = Requisitante.Text

    'Chama o Browser que Lista os Requisitantes
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoRequisitante)

End Sub

Private Sub objEventoRequisitante_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim bCancel As Boolean

    Set objRequisitante = obj1

    'Colocao Nome Reduzido do Requisitante na tela
    Requisitante.Text = objRequisitante.sNomeReduzido

    'Dispara o Validate de Requisitante
    Call Requisitante_Validate(bCancel)

    Me.Show

End Sub

Private Sub CclLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_LabelCcl_Click

    'Critica o formato do centro de custo
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 61866

    objCcl.sCcl = sCclFormatada

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub

Erro_LabelCcl_Click:

    Select Case Err

        Case 61866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173878)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    'Coloca a conta no formato conta enxuta
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 61867

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case Err

        Case 61867

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173879)

        End Select

    Exit Sub

End Sub

Private Sub ObservacaoLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objObservacao As New ClassObservacao

    objObservacao.sObservacao = Observacao.Text

    Call Chama_Tela("ObservacaoLista", colSelecao, objObservacao, objEventoObservacao)

End Sub

Private Sub objEventoObservacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objObservacao As New ClassObservacao

    Set objObservacao = obj1

    Observacao.Text = objObservacao.sObservacao

    Me.Show

End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Public Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo
    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Public Sub BotaoModelos_Click()

Dim objRequisicaoModelo As New ClassRequisicaoModelo
Dim colSelecao As Collection

    'Chama Tela RequisiçãoModeloLista
    Call Chama_Tela("RequisicaoModeloLista", colSelecao, objRequisicaoModelo, objEventoReqModelo)

End Sub

Private Sub objEventoReqModelo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisicaoModelo As New ClassRequisicaoModelo

On Error GoTo Erro_objEventoReqModelo_evSelecao

    Set objRequisicaoModelo = obj1

    'Traz dados da Requisição Modelo para a tela
    lErro = Traz_RequisicaoModelo_Tela(objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61868

    Me.Show

    Exit Sub

Erro_objEventoReqModelo_evSelecao:

    Select Case Err

        Case 61868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173880)

    End Select

    Exit Sub

End Sub

Function Traz_RequisicaoCompra_Tela(objRequisicaoCompra As ClassRequisicaoCompras) As Long

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim objObservacao As New ClassObservacao
Dim bCancel As Boolean
Dim objRequisicaoModelo As New ClassRequisicaoModelo
Dim iIndice As Integer

On Error GoTo Erro_Traz_RequisicaoCompra_Tela

    'Limpa tela Requisicao Compras
    Call Limpa_Tela_RequisicaoCompra

    'Coloca os dados na tela
    Codigo.PromptInclude = False
    Codigo.Text = CLng(objRequisicaoCompra.lCodigo)
    Codigo.PromptInclude = True
    
    If objRequisicaoCompra.dtData <> DATA_NULA Then
        Data.Text = Format(objRequisicaoCompra.dtData, "dd/mm/yy")
    End If
            
    CodigoOP.Text = objRequisicaoCompra.sOPCodigo
        
    If objRequisicaoCompra.dtDataLimite <> DATA_NULA Then
        DataLimite.Text = Format(objRequisicaoCompra.dtDataLimite, "dd/mm/yy")
    End If
        
    Urgente.Value = objRequisicaoCompra.lUrgente

    'Verifica se Observacao esta preenchido
    If objRequisicaoCompra.lObservacao <> 0 Then

        objObservacao.lNumInt = objRequisicaoCompra.lObservacao
        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError (65570)
        If lErro = 53827 Then gError (65572)
        Observacao.Text = objObservacao.sObservacao

    End If

    If objRequisicaoCompra.lRequisitante <> 0 Then
        
        objRequisitante.lCodigo = objRequisicaoCompra.lRequisitante

        'Le o requisitante para colocar o NomeReduzido na tela
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError (65566)
        If lErro = 49084 Then gError (65571)
    
        Requisitante.Text = objRequisitante.sNomeReduzido
    
    End If
    
    If Len(Trim(objRequisicaoCompra.sCcl)) > 0 Then
    
        'Preenche a CCL
        sCclMascarado = String(STRING_CCL, 0)
    
        lErro = Mascara_RetornaCclEnxuta(objRequisicaoCompra.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError (65567)
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
    
    End If
    
    'Tipo de Tributação
    For iIndice = 0 To TipoTributacao.ListCount - 1
        If objRequisicaoCompra.iTipoTributacao = Codigo_Extrai(TipoTributacao.List(iIndice)) Then
            TipoTributacao.ListIndex = iIndice
            Exit For
        End If
    Next
    
    FilialCompra.Text = objRequisicaoCompra.iFilialCompra
    Call FilialCompra_Validate(bCancel)

    objRequisicaoModelo.iTipoDestido = objRequisicaoCompra.iTipoDestino
    objRequisicaoModelo.iFilialDestino = objRequisicaoCompra.iFilialDestino
    objRequisicaoModelo.lFornCliDestino = objRequisicaoCompra.lFornCliDestino
    objRequisicaoModelo.iTipoTributacao = objRequisicaoCompra.iTipoTributacao
    
    'Preenche TipoDestino e suas Caracteristicas
    lErro = Preenche_TipoDestino(objRequisicaoModelo)
    If lErro <> SUCESSO Then gError (65568)

    'Lê os itens da Requisicao de Compras
    lErro = CF("ItensReqCompras_Le_ReqCompras", objRequisicaoCompra)
    If lErro <> SUCESSO And lErro <> 65603 Then gError (61529)

    'Se não encontrou itens, erro
    If objRequisicaoCompra.colItens.Count = 0 Then gError (65565)

    Call Move_ReqCompra_ReqModelo(objRequisicaoCompra, objRequisicaoModelo)
        
    'Preenche o grid com os Itens da requição modelo
    lErro = Preenche_GridItens(objRequisicaoModelo)
    If lErro <> SUCESSO Then gError 65569
    
    lErro = Preenche_GridPV(objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 178644

    iAlterado = 0

    Traz_RequisicaoCompra_Tela = SUCESSO

    Exit Function

Erro_Traz_RequisicaoCompra_Tela:

    Traz_RequisicaoCompra_Tela = gErr

    Select Case gErr

        Case 61529, 65566, 65567, 65568, 65569, 65570, 178644

        Case 65565
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_COMPRAS_AUSENCIA_ITENS", gErr, objRequisicaoCompra.lCodigo)

        Case 65571
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)

        Case 65572
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173881)

    End Select

    Exit Function

End Function

Sub Move_ReqCompra_ReqModelo(objRequisicaoCompra As ClassRequisicaoCompras, objRequisicaoModelo As ClassRequisicaoModelo)

Dim objItemReqCompra As ClassItemReqCompras
Dim objItemReqModelo As ClassItemReqModelo

    For Each objItemReqCompra In objRequisicaoCompra.colItens
    
        Set objItemReqModelo = New ClassItemReqModelo
        
        objItemReqModelo.sProduto = objItemReqCompra.sProduto
        objItemReqModelo.sDescProduto = objItemReqCompra.sDescProduto
        objItemReqModelo.sUM = objItemReqCompra.sUM
        objItemReqModelo.dQuantidade = objItemReqCompra.dQuantidade
        objItemReqModelo.iAlmoxarifado = objItemReqCompra.iAlmoxarifado
        objItemReqModelo.sCcl = objItemReqCompra.sCcl
        objItemReqModelo.sContaContabil = objItemReqCompra.sContaContabil
        objItemReqModelo.iTipoTributacao = objItemReqCompra.iTipoTributacao
        objItemReqModelo.iExclusivo = objItemReqCompra.iExclusivo
        objItemReqModelo.lFornecedor = objItemReqCompra.lFornecedor
        objItemReqModelo.iFilial = objItemReqCompra.iFilial
        objItemReqModelo.lObservacao = objItemReqCompra.lObservacao
        
        objRequisicaoModelo.colItensReqModelo.Add objItemReqModelo
    
    Next

End Sub

Function Traz_RequisicaoModelo_Tela(objRequisicaoModelo As ClassRequisicaoModelo) As Long

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim objObservacao As New ClassObservacao
Dim bCancel As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Traz_RequisicaoModelo_Tela

    'Limpa tela Requisicao Compras
    Call Limpa_Tela_RequisicaoCompra

    'Coloca os dados na tela
    Urgente.Value = objRequisicaoModelo.iUrgente

    'Verifica se Observacao esta preenchido
    If objRequisicaoModelo.lObservacao <> 0 Then

        objObservacao.lNumInt = objRequisicaoModelo.lObservacao
        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError 65570
        If lErro = 53827 Then gError 65572
        Observacao.Text = objObservacao.sObservacao

    End If

    If objRequisicaoModelo.lRequisitante <> 0 Then
        
        objRequisitante.lCodigo = objRequisicaoModelo.lRequisitante

        'Le o requisitante para colocar o NomeReduzido na tela
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 65566
        If lErro = 49084 Then gError 65571
    
        Requisitante.Text = objRequisitante.sNomeReduzido
        
    End If
    
    'Tipo de Tributação
    For iIndice = 0 To TipoTributacao.ListCount - 1
        If Codigo_Extrai(TipoTributacao.List(iIndice)) = objRequisicaoModelo.iTipoTributacao Then
            TipoTributacao.ListIndex = iIndice
            Exit For
        End If
    Next
    
    If Len(Trim(objRequisicaoModelo.sCcl)) > 0 Then
    
        'Preenche a CCL
        sCclMascarado = String(STRING_CCL, 0)
    
        lErro = Mascara_MascararCcl(objRequisicaoModelo.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 65567
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
    
    End If
    
    FilialCompra.Text = objRequisicaoModelo.iFilialCompra
    Call FilialCompra_Validate(bCancel)

    'Preenche TipoDestino e suas Caracteristicas
    lErro = Preenche_TipoDestino(objRequisicaoModelo)
    If lErro <> SUCESSO Then gError (65568)

    'Lê os itens da Requisicao Modelo
    lErro = CF("ItensReqModelo_Le", objRequisicaoModelo)
    If lErro <> SUCESSO And lErro <> 61533 Then gError (61529)

    'Se não encontrou itens, erro
    If objRequisicaoModelo.colItensReqModelo.Count = 0 Then gError (65565)

    'Preenche o grid com os Itens da requição modelo
    lErro = Preenche_GridItens(objRequisicaoModelo)
    If lErro <> SUCESSO Then gError (65569)

    iAlterado = 0

    Traz_RequisicaoModelo_Tela = SUCESSO

    Exit Function

Erro_Traz_RequisicaoModelo_Tela:

    Traz_RequisicaoModelo_Tela = gErr

    Select Case gErr

        Case 61529, 65566, 65567, 65568, 65569, 65570

        Case 65565
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_MODELO_AUSENCIA_ITENS", gErr, objRequisicaoModelo.lCodigo)

        Case 65571
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)

        Case 65572
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173882)

    End Select

    Exit Function

End Function

Function Preenche_GridItens(objRequicaoModelo As ClassRequisicaoModelo) As Long

Dim lErro As Long
Dim objItemReqModelo As ClassItemReqModelo
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim objObservacao As New ClassObservacao
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iIndice As Integer
Dim iItem As Integer
Dim sProdutoMascarado As String
Dim sCclMascarado As String
Dim sContaEnxuta As String
Dim iCont As Integer

On Error GoTo Erro_Preenche_GridItens
    
    Set gColItemReqCompras = New Collection
    
    'Preenche GridItens
    For Each objItemReqModelo In objRequicaoModelo.colItensReqModelo

        iIndice = iIndice + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Colococa o Produto Mascarado no Grid
        lErro = Mascara_RetornaProdutoEnxuto(objItemReqModelo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError (65573)

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        
        GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemReqModelo.sDescProduto
        GridItens.TextMatrix(iIndice, iGrid_UM_Col) = objItemReqModelo.sUM
        If objItemReqModelo.dQuantidade > 0 Then GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Format(objItemReqModelo.dQuantidade, "Standard")
        
        If objRequicaoModelo.iTipoDestido = TIPO_DESTINO_EMPRESA Then
        
            If objItemReqModelo.iAlmoxarifado <> 0 Then
                
                'Lê o Almoxarifado e coloca seu nome Reduzido no Grid
                objAlmoxarifado.iCodigo = objItemReqModelo.iAlmoxarifado
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError (65576)
                If lErro = 25056 Then gError (65581)
            
                GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            
            End If
                        
            sCclMascarado = ""
            
            'Coloca o Ccl mascarado no Grid
            If objItemReqModelo.sCcl <> "" Then
            
                lErro = Mascara_RetornaCclEnxuta(objItemReqModelo.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError (65577)
            
                Ccl.PromptInclude = False
                Ccl.Text = sCclMascarado
                Ccl.PromptInclude = True
                
                GridItens.TextMatrix(iIndice, iGrid_CCL_Col) = Ccl.Text

            End If

                    
        End If
        
        'Coloca Conta Contábil no Grid
        If objItemReqModelo.sContaContabil <> "" Then

            lErro = Mascara_RetornaContaEnxuta(objItemReqModelo.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError (65574)

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True
            
            GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col) = ContaContabil.Text

        End If
        
        'Tipo de Tributação
        For iCont = 0 To TipoTributacao.ListCount - 1
            If Codigo_Extrai(TipoTributacao.List(iCont)) = objItemReqModelo.iTipoTributacao Then
                GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col) = TipoTributacao.List(iCont)
                Exit For
            End If
        Next
        
        'Coloca Nome Reduzido do Fornecedor no Grid
        If objItemReqModelo.lFornecedor > 0 Then

            objFornecedor.lCodigo = objItemReqModelo.lFornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError (65575)
            If lErro = 12729 Then gError (65582)

            GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido

        End If

        'Coloca Filial do Fornecedor no Grid
        If objItemReqModelo.iFilial > 0 Then
            
            objFilialFornecedor.iCodFilial = objItemReqModelo.iFilial
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError (65579)
            If lErro = 12929 Then gError (65583)
            
            GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        End If
        
        If objItemReqModelo.lFornecedor > 0 Then
        
            'Preenche Exclusivo
            For iItem = 0 To Exclusivo.ListCount - 1
                If objItemReqModelo.iExclusivo = Exclusivo.ItemData(iItem) Then
                    GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) = Exclusivo.List(iItem)
                    Exit For
                End If
            Next
                
        End If
        
        'Se possui observação
        If objItemReqModelo.lObservacao <> 0 Then

            objObservacao.lNumInt = objItemReqModelo.lObservacao

            'Lê a observação a partir do número interno
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError (65580)
            If lErro <> SUCESSO Then gError (65584)

            GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao

        End If
    
        gColItemReqCompras.Add objItemReqModelo.lNumIntDoc
        
    Next
    
    lErro = Grid_Refresh_Checkbox(objGridItens)
    If lErro <> SUCESSO Then gError (65578)

    objGridItens.iLinhasExistentes = gColItemReqCompras.Count

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 65573, 65574, 65575, 65576, 65577, 65578, 65579, 65580
        
        Case 65581
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
                    
        Case 65582
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
        
        Case 65583
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
            
        Case 65584
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173883)

    End Select

    Exit Function

End Function

Function Preenche_TipoDestino(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Preenche o Tipo destino

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente
Dim bCancel As Boolean

On Error GoTo Erro_Preenche_TipoDestino

    TipoDestino.Item(objRequisicaoModelo.iTipoDestido).Value = True

    Select Case objRequisicaoModelo.iTipoDestido

        Case TIPO_DESTINO_EMPRESA

            FilialEmpresa.Text = objRequisicaoModelo.iFilialDestino
            Call FilialEmpresa_Validate(bCancel)

        Case TIPO_DESTINO_FORNECEDOR
            
            objFornecedor.lCodigo = objRequisicaoModelo.lFornCliDestino

            'Lê o fornecedor, seu nome reduzido
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError (65585)
            
            'Se não encontrou o Forncedor, Erro
            If lErro = 12729 Then gError (65588)
            
            Fornecedor.Text = objFornecedor.sNomeReduzido

            FilialFornecedor.Text = objRequisicaoModelo.iFilialDestino
            Call FilialFornecedor_Validate(bCancel)
            
        Case Else
            gError (65587)

    End Select

    Preenche_TipoDestino = SUCESSO

    Exit Function

Erro_Preenche_TipoDestino:

    Preenche_TipoDestino = gErr

    Select Case gErr

        Case 65585, 65587

        Case 65588
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173884)

    End Select

    Exit Function

End Function

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 61869
   
    '###############################################
    'Inserido por Wagner 05/05/06
    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 177416
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
    '###############################################

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoLista_Consulta
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 61869
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 177416

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173885)

    End Select

    Exit Sub

End Sub

Public Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoEnxuto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 61872

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoEnxuto = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then Error 61874

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 61873

            'Se não encontrou o Produto --> Erro
            If lErro = 28030 Then Error 61875

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is Produto) Then

                GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then Error 61871

            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 61871, 61872, 61873

        Case 61874
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 61875
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173886)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sContaEnxuta As String
Dim objProdutoFilial As New ClassProdutoFilial
Dim objFornecedor As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_ProdutoLinha_Preenche

    'Preenche no Grid a Descrição do Produto e a Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = objProduto.sSiglaUMCompra
    GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao

    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True Then
        If Len(Trim(FilialEmpresa.Text)) > 0 Then
            lErro = CF("AlmoxarifadoPadrao_Le", Codigo_Extrai(FilialEmpresa.Text), objProduto.sCodigo, iAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 61876
        
            If lErro = SUCESSO And iAlmoxarifado <> 0 Then
        
                objAlmoxarifado.iCodigo = iAlmoxarifado
        
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 61877
        
                If lErro = 25056 Then gError 61878
                
                'Coloca o Nome Reduzido na Coluna Almoxarifado
                GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
                                    
                'Coloca Conta Contábil no GridItens
                objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
                objEstoqueProduto.sProduto = objProduto.sCodigo
                lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                If lErro <> SUCESSO And lErro <> 49991 Then gError 66593
    
                If lErro <> 49991 Then
    
                    lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
                    If lErro <> SUCESSO Then gError 66594
            
                    ContaContabil.PromptInclude = False
                    ContaContabil.Text = sContaEnxuta
                    ContaContabil.PromptInclude = True
            
                    GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
            
                Else
            
                    'Preenche em branco a conta de estoque no grid
                    GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ""
            
                End If
            End If
        End If
        
        'Preenche Ccl
        If Len(Trim(Ccl.ClipText)) > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col) = Ccl.Text
        End If

    End If
        
    'Preenche Tipo de Tributação
    If Len(Trim(TipoTributacao.Text)) > 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_TipoTributacao_Col) = TipoTributacao.Text
    End If
    
    objProdutoFilial.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
    objProdutoFilial.sProduto = objProduto.sCodigo

    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 62666
    If lErro = SUCESSO Then
        If (objProdutoFilial.lFornecedor > 0) And (objProdutoFilial.iFilialForn > 0) Then
            
            objFornecedor.lCodigo = objProdutoFilial.lFornecedor
            
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 62667
            If lErro <> SUCESSO Then gError 62668
            
            objFilialForn.lCodFornecedor = objFornecedor.lCodigo
            objFilialForn.iCodFilial = objProdutoFilial.iFilialForn
            
            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 62669
            If lErro <> SUCESSO Then gError 62670
                
            GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
            GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = objFilialForn.iCodFilial & SEPARADOR & objFilialForn.sNome
            GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = "Preferencial"
                
        End If
    End If
    
    'Se necessário cria uma nova linha no Grid
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        gColItemReqCompras.Add (0)
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 61876, 61877, 66593, 66594, 62666, 62667, 62669

        Case 61878
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case 62668
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 62670
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialForn.iCodFilial, objFilialForn.lCodFornecedor)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173887)

    End Select

    Exit Function

End Function

Private Sub BotaoCcl_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_BotaoCcls_Click

    'Se nenhuma linha foi selecionada do Grid, Erro
    If GridItens.Row = 0 Then Error 61879
    
    'Se o campo está desabilitado, sai da rotina
    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = False Then Exit Sub
    
    'Verifica se o Produto foi preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then Error 61880

    'Verifica se o Ccl Foi preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col))) > 0 Then

        sCclFormatada = String(STRING_CCL, 0)

        lErro = CF("Ccl_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col), sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then Error 61881

        objCcl.sCcl = sCclFormatada

    End If

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoBotaoCcl)

    Exit Sub

Erro_BotaoCcls_Click:

    Select Case Err

        Case 61879
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 61880
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 61881

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173888)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se nenhuma foi selecionada, erro
    If GridItens.Row = 0 Then Error 61882

    sContaEnxuta = String(STRING_CCL, 0)

    'Coloca a conta no formato conta enxuta
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 61883

    GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col) = sCclMascarado
    CentroCusto.PromptInclude = False
    CentroCusto.Text = sCclMascarado
    CentroCusto.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case Err

        Case 61882
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 61883

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173889)

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
Dim iFilialEntrega As Integer

On Error GoTo Erro_BotaoALmoxarifados_Click

    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = False Then Exit Sub

    If GridItens.Row = 0 Then Error 61884

    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    sProdutoFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 61886

    'Carrega a Variável com os dados do frame visível
    iFilialEntrega = Codigo_Extrai(FilialEmpresa.Text)
    If Len(Trim(iFilialEntrega)) = 0 Then Error 31151
    
    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        
        colSelecao.Add iFilialEntrega
        
        Call Chama_Tela("AlmoxarifadoFilialLista", colSelecao, objEstoqueProduto, objEventoAlmoxarifados)
    Else
        Error 61885
    End If

    Exit Sub

Erro_BotaoALmoxarifados_Click:

    Select Case Err

        Case 31151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_ENTREGA_NAO_PREENCHIDA", Err)
        
        Case 61884
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 61885
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 61886

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173890)

    End Select

    Exit Sub

End Sub

Public Sub objEventoAlmoxarifados_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto
Dim bCancel As Boolean

    Set objEstoqueProduto = obj1

    'Preenche campo Almoxarifado
    GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    Me.Show

    Exit Sub

End Sub

Private Sub BotaoFiliaisFornProd_Click()

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoFiliaisFornProd_Click

    'Se nenhuma linha foi selecionada, erro
    If GridItens.Row = 0 Then Error 61888
    
    'Se a  FilialCompra não estiver prrenchida, erro
    If Len(Trim(FilialCompra.Text)) = 0 Then Error 61745
    
    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 61887

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        colSelecao.Add Codigo_Extrai(FilialCompra.Text)
        
        Call Chama_Tela("FiliaisFornProdutoLista", colSelecao, objFornecedorProdutoFF, objEventoFiliaisFornProduto)
    Else
        Error 61889
    End If

    Exit Sub

Erro_BotaoFiliaisFornProd_Click:

    Select Case Err

        Case 61745
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", Err)

        Case 61887

        Case 61888
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 61889
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173891)

    End Select

    Exit Sub

End Sub

Public Sub objEventoFiliaisFornProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim objFornecedor As New ClassFornecedor
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_objEventoFiliaisFornProd_evSelecao

    Set objFornecedorProdutoFF = obj1

    'Lê o Nome Reduzido do Fornecedor
    objFornecedor.lCodigo = objFornecedorProdutoFF.lFornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO Then gError 61890

    'Preenche campo Fornecedor
    GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
    FornecGrid.Text = objFornecedor.sNomeReduzido
    
    'Lê coleção de códigos e nomes da Filial do Fornecedor
    lErro = CF("FornecedorProdutoFF_Le_FilialForn", objFornecedorProdutoFF.sProduto, objFornecedor.lCodigo, objFornecedorProdutoFF.iFilialEmpresa, colCodigoNome)
    If lErro <> SUCESSO Then gError 61891
    
    'Se não encontrou nenhuma Filial, erro
    If colCodigoNome.Count = 0 Then gError 66638
    
    'Se foi passada um Filial como parâmetro
    If objFornecedorProdutoFF.iFilialForn > 0 Then

        'Verifica se ela está presente na coleção de filiais
        For iIndice = 1 To colCodigoNome.Count
            If colCodigoNome.Item(iIndice).iCodigo = objFornecedorProdutoFF.iFilialForn Then
                Exit For
            End If
        Next
    
        'Se não encontrou, erro
        If iIndice > colCodigoNome.Count Then gError 66640
    
    'Se não foi passada uma filial como parâmetro
    ElseIf objFornecedorProdutoFF.iFilialForn = 0 Then
        'Coloca como default a primeira filial da coleção
        objFornecedorProdutoFF.iFilialForn = colCodigoNome.Item(1).iCodigo
    End If

    'Coloca no Grid a filial passada
    For iIndice = 1 To colCodigoNome.Count
        If colCodigoNome.Item(iIndice).iCodigo = objFornecedorProdutoFF.iFilialForn Then
            GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
            Exit For
        End If
    Next
            
    'Se não foi preenchida a exclusividade, coloca como default "Preferencial"
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col))) = 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = "Preferencial"
    End If

    Me.Show

    Exit Sub

Erro_objEventoFiliaisFornProd_evSelecao:

    Select Case gErr

        Case 61890, 61891, 66638, 66639

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173892)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Se a Data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 61892

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True
    
    Select Case Err

        Case 61892

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173893)

    End Select

    Exit Sub

End Sub

Private Sub TipoTributacao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then Error 61893

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case Err

        Case 61893

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173894)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then Error 61894

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case Err

        Case 61894

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173895)

    End Select

    Exit Sub

End Sub

Private Sub DataLimite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimite_Validate

    'Se a data limite está preenchida
    If Len(Trim(DataLimite.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataLimite.Text)
        If lErro <> SUCESSO Then Error 61898

    End If

    Exit Sub

Erro_DataLimite_Validate:

    Cancel = True
    
    Select Case Err

        Case 61898

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173896)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_DownClick

    'Se a data limite está preenchida
    If Len(Trim(DataLimite.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataLimite, DIMINUI_DATA)
        If lErro <> SUCESSO Then Error 61899

    End If

    Exit Sub

Erro_UpDownDataEntrega_DownClick:

    Select Case Err

        Case 61899

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173897)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_UpClick

    'Se a data limite está preenchida
    If Len(Trim(DataLimite.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataLimite, AUMENTA_DATA)
        If lErro <> SUCESSO Then Error 61900

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownDataEntrega_UpClick:

    Select Case Err

        Case 61900

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173898)

    End Select

    Exit Sub

End Sub

Private Sub Requisitante_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclEnxuta As String
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Requisitante_Validate

    'Se o Requisitante não foi preenchido, sai da rotina
    If Len(Trim(Requisitante.Text)) = 0 Then Exit Sub
            
    lErro = TP_Requisitante_Le(Requisitante, objRequisitante)
    If lErro <> SUCESSO Then Error 61903
    
    If Len(Trim(objRequisitante.sCcl)) > 0 Then
        
        lErro = Mascara_RetornaCclEnxuta(objRequisitante.sCcl, sCclEnxuta)
        If lErro <> SUCESSO Then gError 79988
        
        Ccl.PromptInclude = False
        Ccl.Text = sCclEnxuta
        Ccl.PromptInclude = True
    End If
    
    Exit Sub

Erro_Requisitante_Validate:

    Cancel = True
    
    Select Case Err

        Case 61903
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173899)

    End Select

    Exit Sub

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Ccl_Validate

    'Se o Ccl não está preenchido, sai da rotina
    If Len(Trim(Ccl.ClipText)) = 0 Then Exit Sub

    'Critica o Ccl
    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then Error 61901

    'Se o Ccl não está cadastrado, erro
    If lErro = 5703 Then Error 61902

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True
    
    Select Case Err

        Case 61901

        Case 61902

            'Pergunta se deseja cadastrar nova Ccl
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsg = vbYes Then
                objCcl.sCcl = sCclFormatada
                Call Chama_Tela("CclTela", objCcl)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173900)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61905

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.lCodEmpresa = glEmpresa
        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 61904

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then Error 61906

        'coloca na tela
        FilialCompra.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 61907

    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case Err

        Case 61904, 61905
            
        Case 61906
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, FilialCompra.Text)

        Case 61907
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", Err, FilialCompra.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173901)

    End Select

    Exit Sub

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

        'Limpa os almoxarifados e os Ccls do GridItens
        For iIndice = 1 To objGridItens.iLinhasExistentes
            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = ""
            GridItens.TextMatrix(iIndice, iGrid_CCL_Col) = ""
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173902)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        'Verifica preenchimento de Fornecedor
        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 61915

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then Error 61916

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornecedor, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornecedor, iCodFilial)

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            'Se Fornecedor não foi preenchido limpa a combo de Filiais
            FilialFornecedor.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 61915, 61916

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173903)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FilialFornecedor_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornecedor.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornecedor.Text = FilialFornecedor.List(FilialFornecedor.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFornecedor, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61918

    'Se não encontrar o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 61917

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 61919

        'Se não achou a Filial Fornecedor --> erro
        If lErro = 18272 Then Error 61920

        'coloca na tela
        FilialFornecedor.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 61921

    Exit Sub

Erro_FilialFornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 61917
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
    
        Case 61918, 61919
    
        Case 61920
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then

                objFornecedor.sNomeReduzido = Fornecedor.Text

                'Lê Fornecedor no BD
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)

                'Se achou o Fornecedor --> coloca o codigo em objFilialFornecedor
                If lErro = SUCESSO Then objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            
            End If

        Case 61921
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, FilialFornecedor.Text)
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173904)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub DataLimite_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataLimite, iAlterado)
    
End Sub

Private Sub Ccl_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Data_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimite_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Urgente_Click()
    
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
    If lErro <> SUCESSO Then gError (65562)
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub
    
Erro_FilialEmpresa_Click:
    
    Select Case gErr
        
        Case 65562 'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173905)
        
    End Select
    
    Exit Sub

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
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    
            objProduto.sCodigo = sProdutoFormatado
            
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError (65558)
    
            'Se não encontrou o Produto --> Erro
            If lErro = 28030 Then gError (65560)
            
            'Se o Produto possui Estoque
            If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                'Lê dados do seu almoxarifado Padrão
                lErro = CF("AlmoxarifadoPadrao_Le", iCodFilial, objProduto.sCodigo, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then gError (61710)
    
                'Se encontrou
                If lErro = SUCESSO And iAlmoxarifadoPadrao <> 0 Then
    
                    objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
    
                    'Lê os dados do Almoxarifado a partir do código passado
                    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25056 Then gError (65559)
            
                    'Se não encontrou, erro
                    If lErro = 25056 Then gError (65561)
            
                    'Coloca o Nome Reduzido na Coluna Almoxarifado
                    GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
                Else
                    GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = ""
                End If
                
            
            End If
        
        End If
        
    Next
    
    AlmoxarifadoPadrao_Preenche = SUCESSO
    
    Exit Function
    
Erro_AlmoxarifadoPadrao_Preenche:

    AlmoxarifadoPadrao_Preenche = gErr
    
    Select Case gErr
        
        Case 65558, 65559
        
        Case 65560
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 65561
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173906)
    
    End Select
    
    Exit Function
    
End Function

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim iCodFilial As Integer

On Error GoTo Erro_FilialEmpresa_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialEmpresa.Text)) = 0 Then Exit Sub

    'Verifica se é uma FilialEmpresa selecionada
    If FilialEmpresa.Text = FilialEmpresa.List(FilialEmpresa.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61922

    'Se não encontrou o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 61923

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then Error 61924

        'coloca na tela
        FilialEmpresa.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 61925
        
    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True
    
    Select Case Err

        Case 61922, 61923

        Case 61924
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, FilialEmpresa.Text)
            
        Case 61925
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", Err, FilialEmpresa.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173907)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'Critica o produto passado
        lErro = CF("Produto_Critica_Compra", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25605 Then Error 61926

        'Produto não cadastrado
        If lErro = 25605 Then Error 61930
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Preenche a UM, a Descrição e o Almoxarifado Padrão do Produto
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then Error 61927

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then Error 61929

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61928

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = Err

    Select Case Err
        
'''        Case 61681
'''            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", Err, iIndice)
'''            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 61926, 61927, 61928, 61929
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61930
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173908)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UM

    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UM_Col) = UM.Text

    For iLinha = 1 To objGridPV.iLinhasExistentes
    
        If GridPV.TextMatrix(iLinha, iGrid_ItemReqPV_Col) = GridItens.Row Then

            GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_UMPV_Col) = UM.Text
            
        End If

    Next

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61931

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = Err

    Select Case Err

        Case 61931
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173909)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then Error 61932

        dQuantidade = CDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)

        For iLinha = 1 To objGridPV.iLinhasExistentes
        
            If GridPV.TextMatrix(iLinha, iGrid_ItemReqPV_Col) = GridItens.Row Then

                GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantReqPV_Col) = Quantidade.Text
                
            End If

        Next

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61933

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = Err

    Select Case Err

        Case 61932, 61933
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173910)

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
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sContaEnxuta As String

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado foi preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        'Formata o Produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 61934

        'Lê o Almoxarifado
        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 61935
        If lErro = 25130 Then gError 61937
        If lErro = 25135 Then gError 61938
        If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 86099
    
        'Coloca Conta Contábil no GridItens
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 67865
        
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        'Se o campo Conta Contábil não está preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col))) = 0 Then
        
            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 49991 Then gError 67863
    
            If lErro <> 49991 Then
    
                lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
                If lErro <> SUCESSO Then gError 67864
        
                ContaContabil.PromptInclude = False
                ContaContabil.Text = sContaEnxuta
                ContaContabil.PromptInclude = True
        
                GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
        
            Else
        
                'Preenche em branco a conta de estoque no grid
                GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ""
        
            End If
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 61936

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 61934, 61935, 61936, 67863, 67864, 67865
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61937
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 61938
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 86099
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173911)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Ccl do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = CentroCusto

    'Verifica se Ccl foi preenchido
    If Len(Trim(CentroCusto.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", CentroCusto, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then Error 61939

        If lErro = 5703 Then Error 61940

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61941

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err

    Select Case Err

        Case 61939
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61940
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CentroCusto.Text)
            If vbMsgRes = vbYes Then

                objCcl.sCcl = sCclFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("CclTela", objCcl)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61941
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173912)

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

    Set objGridItens.objControle = ContaContabil

    'Se a Conta Contábil foi preenchida
    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabil.ClipText, objPlanoConta, MODULO_COMPRAS)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 61942

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaEnxuta = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 61945

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_COMPRAS)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 61943

            'conta não cadastrada
            If lErro = 5700 Then Error 61946

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61944

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = Err

    Select Case Err

        Case 61942, 61943, 61944
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61945
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61946
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173913)

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

    Set objGridInt.objControle = FornecGrid

    sFornecedor = FornecGrid.Text
    
    'Se o fornecedor foi preenchido
    If Len(Trim(FornecGrid.ClipText)) > 0 Then

        'Verifica se o fornecedor está cadastrado
        lErro = TP_Fornecedor_Grid(FornecGrid, objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 25611 And lErro <> 25613 And lErro <> 25616 And lErro <> 25619 Then Error 61947

        'Fornecedor não cadastrado
        'Nome Reduzido
        If lErro = 25611 Then Error 61950

        'Codigo
        If lErro = 25613 Then Error 61951

        'CGC/CPF
        If lErro = 25616 Or lErro = 25619 Then Error 61952

        If sFornecedor <> objFornecedor.sNomeReduzido Then
        
            'Formata o Produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 61949
            
            iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
            
            'Lê coleção de códigos e nomes da Filial do Fornecedor
            lErro = CF("FornecedorProdutoFF_Le_FilialForn", sProdutoFormatado, objFornecedor.lCodigo, iFilialEmpresa, colCodigoNome)
            If lErro <> SUCESSO Then Error 61635
    
            'Se não encontrou nenhuma Filial, erro
            If colCodigoNome.Count = 0 Then Error 61953
    
            If iCodFilial > 0 Then
    
                For iIndice = 1 To colCodigoNome.Count
                    If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                        Exit For
                    End If
                Next
    
                If iIndice = colCodigoNome.Count Then Error 61954
    
            ElseIf iCodFilial = 0 Then
                iCodFilial = colCodigoNome.Item(1).iCodigo
            End If
    
            For iIndice = 1 To colCodigoNome.Count
                If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                    GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
                    Exit For
                End If
            Next
        
        End If
    
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col))) = 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = "Preferencial"
        End If
        
    Else
        
        'Limpa a Filial e Exclusividade Correspondente
        GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61948

    Saida_Celula_Fornecedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Fornecedor:

    Saida_Celula_Fornecedor = Err

    Select Case Err

        Case 61947, 61948, 61949
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61950 'Fornecedor com Nome Reduzido %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", FornecGrid.Text)
            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61951 'Fornecedor com código %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", FornecGrid.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61952 'Fornecedor com CGC/CPF %s não encontado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_3", FornecGrid.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61953
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FIL_PROD_FORN_FILIALCOMPRA", Err, objFornecedor.sNomeReduzido, sProdutoFormatado)
            FornecGrid.Text = ""
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61954
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", Err, iCodFilial, objFornecedor.sNomeReduzido, sProdutoFormatado)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173914)

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

    Set objGridInt.objControle = FilialFornecGrid
    
    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornecGrid.Text)) > 0 Then

        'Verifica se não é uma filial selecionada
        If Not FilialFornecGrid.Text = FilialFornecGrid.List(FilialFornecGrid.ListIndex) Then

            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialFornecGrid, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError (61955)
    
            'Se nao encontra o ítem com o código informado
            If lErro = 6730 Then
    
                'Verifica se o Fornecedor foi preenchido
                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col))) = 0 Then gError (61960)
    
                lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError (61959)
                
                sFornecedor = GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col)
                objFornecedorProdutoFF.iFilialForn = iCodigo
                objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa
                objFornecedorProdutoFF.sProduto = sProdutoFormatado
    
                'Pesquisa se existe filial com o codigo extraido
                lErro = CF("FornecedorProdutoFF_Le_NomeRed", sFornecedor, sNomeFilial, objFornecedorProdutoFF)
                If lErro <> SUCESSO And lErro <> 61780 Then gError (61956)
    
                'Se não encontrou a Filial do Fornecedor
                If lErro = 61780 Then
    
                    'Lê FilialFornecedor do BD
                    objFilialFornecedor.iCodFilial = iCodigo
                    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
                    If lErro <> SUCESSO And lErro <> 18272 Then gError (61957)
    
                    'Se não encontrou, pergunta se deseja criar
                    If lErro = 18272 Then
                        gError (61961)
                    
                    'Se encontrou, erro
                    Else
                        gError (61963)
                    End If
                
                'Se encontrou a Filial do Fornecedor
                Else
    
                    'coloca na tela
                    FilialFornecGrid.Text = iCodigo & SEPARADOR & sNomeFilial
    
                End If
    
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError (61962)

        End If
    Else
        gError 86096
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError (61958)

    Saida_Celula_FilialForn = SUCESSO
    
    Exit Function

Erro_Saida_Celula_FilialForn:

    Saida_Celula_FilialForn = gErr
    
    Select Case gErr

        Case 61955, 61956, 61957, 61958, 61959
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61960
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_FORNECEDOR_NAO_PREENCHIDO", gErr, GridItens.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61961

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

        Case 61962
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_NAO_ENCONTRADA_ASSOCIADA", gErr, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 61963
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", gErr, objFilialFornecedor.iCodFilial, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 86096
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173915)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Exclusivo(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Exclusivo do grid que está deixando de ser a corrente
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Exclusivo

    Set objGridInt.objControle = Exclusivo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61964

    Saida_Celula_Exclusivo = SUCESSO

    Exit Function

Erro_Saida_Celula_Exclusivo:

    Saida_Celula_Exclusivo = Err

    Select Case Err

        Case 61964
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173916)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Observacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = ObservacaoGrid

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61965

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = Err

    Select Case Err

        Case 61965
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173917)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoTributacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo de Tributação que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TipoTributacao

    Set objGridInt.objControle = TipoTribItem

    GridItens.TextMatrix(GridItens.Row, iGrid_TipoTributacao_Col) = TipoTribItem.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 66598

    Saida_Celula_TipoTributacao = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoTributacao:

    Saida_Celula_TipoTributacao = gErr

    Select Case gErr

        Case 66598
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173918)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRequisicaoCompra As New ClassRequisicaoCompras
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o código não estiver preenchido, Erro
    If Len(Trim(Codigo.Text)) = 0 Then gError (65534)
     
    objRequisicaoCompra.iFilialEmpresa = giFilialEmpresa
    objRequisicaoCompra.lCodigo = StrParaLong(Codigo.Text)
    
    'Lê a Requisição de Compras a partir do código Passado em objRequisicaoCompra
    lErro = CF("RequisicaoCompras_Le_Numero", objRequisicaoCompra)
    If lErro <> SUCESSO And lErro <> 54439 Then gError (65535)
    If lErro = 54439 Then gError (65536)
    
    'Se a data de envio estiver preenchida -> Erro
    If objRequisicaoCompra.dtDataEnvio <> DATA_NULA Then gError (65563)
    
    'Envia aviso perguntando se realmente deseja excluir a Requisição de Compras
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_REQUISICAOCOMPRA", objRequisicaoCompra.lCodigo)

    'Se a resposta for positiva
    If vbMsgRes = vbYes Then
    
        'Exclui a Requisição de Compras
        lErro = CF("RequisicaoCompra_Exclui", objRequisicaoCompra)
        If lErro <> SUCESSO Then gError (65537)
        
        'Limpa a Tela
        Call Limpa_Tela_RequisicaoCompra
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        iAlterado = 0
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 65534
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 65535, 65537
            
        Case 65536
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objRequisicaoCompra.lCodigo)
            
        Case 65563
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_COMPRA_ENVIADA", gErr, objRequisicaoCompra.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173919)
    
    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 61966

    'Limpa a tela
    Call Limpa_Tela_RequisicaoCompra

    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 61966

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173920)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro()

Dim lErro As Long
Dim objRequisicaoCompra As New ClassRequisicaoCompras
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica o preenchimento dos campos obrigatórios
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 61967
    If Len(Trim(Data.ClipText)) = 0 Then gError 61968
    If Len(Trim(Requisitante.Text)) = 0 Then gError 61969
    
    'Verifica se a FilialCompra está preenchida
    If giTipoVersao = VERSAO_FULL Then
        If Len(Trim(FilialCompra.Text)) = 0 Then gError 61970
    End If
    
    'Verifica se a Data Limite é maior que a data de envio
    If (EnviaCompras.Value = vbChecked) And Len(Trim(DataLimite.ClipText)) > 0 Then
        If CDate(DataLimite.Text) < Date Then gError 54171
    End If
    
    'Verifica se a Data Limite é maior que a data de envio
    If Len(Trim(DataLimite.ClipText)) > 0 Then
        If StrParaDate(DataLimite.Text) < StrParaDate(Data.Text) Then gError 62710
    End If
    
    'Verifica se o Grid foi preenchido
    If objGridItens.iLinhasExistentes = 0 Then gError 61971

    'Se o tipo destino for empresa
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then
    
        'Se a FilialEmpresa não estiver preenchida, erro
        If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 61972
    
        'Critica o preenchimento do Almoxarifado
        For iIndice = 1 To objGridItens.iLinhasExistentes
                
            'Formata o Produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 67880
            
            objProduto.sCodigo = sProdutoFormatado
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 23080 Then gError 67881
            If lErro <> SUCESSO Then gError 67882
            
            'Se o produto possui Estoque e o almoxarifado não foi preenchido, erro
            If Not (objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 66587
                    
            'Verifica se existe no Grid outro produto com o mesmo leque Fornecedor, Filial, Exclusividade, Almoxarifado
            For iIndice2 = iIndice + 1 To objGridItens.iLinhasExistentes
                If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridItens.TextMatrix(iIndice2, iGrid_Produto_Col) And GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col) = GridItens.TextMatrix(iIndice2, iGrid_Fornecedor_Col) And GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col) = GridItens.TextMatrix(iIndice2, iGrid_FilialFornecedor_Col) And GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = GridItens.TextMatrix(iIndice2, iGrid_Almoxarifado_Col) And GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) = GridItens.TextMatrix(iIndice2, iGrid_Exclusivo_Col) Then gError 62683
            Next
        
        Next
    
    'Se o tipo destino for Fornecedor
    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then
    
        'Se o Fornecedor não estiver preenchido, erro
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 61973
        
        'Se a Filial do Fornecedor não estiver preenchida, erro
        If Len(Trim(FilialFornecedor.Text)) = 0 Then gError 61974
            
    End If
                
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 61977
    
    'Recolhe os dados do Grid
    lErro = Move_GridItens_Memoria(objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 61978
                
    'Recolhe os dados do Grid
    lErro = Move_GridPV_Memoria(objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 178637
                
    lErro = Trata_Alteracao(objRequisicaoCompra, objRequisicaoCompra.iFilialEmpresa, objRequisicaoCompra.lCodigo)
    If lErro <> SUCESSO Then gError 89156
                
    Set objRequisicaoCompra.objTela = Me
                
    'Grava a Requisição de Compras
    lErro = CF("RequisicaoCompra_Grava", objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 61979
    
    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 54171
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_MAIOR_DATAENVIO", gErr, DataLimite.Text, Date)
        
        Case 61967
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 61968
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            
        Case 61969
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_PREENCHIDO", gErr)
            
        Case 61970
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", gErr)
        
        Case 61971
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDITENS_VAZIO", gErr)
         
        Case 61972
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)
        
        Case 61973
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO", gErr)
        
        Case 61974
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA", gErr)
                        
        Case 61977, 61978, 61979, 67880, 67881, 89156, 178637
        
        Case 62683
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENS_MESMO_LEQUE", gErr, iIndice, iIndice2)
        
        Case 62710
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_INFERIOR_DATAREQ", gErr)
        
        Case 66587
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
            
        Case 67882
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173921)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function
    
End Function

Function Move_Tela_Memoria(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Move os dados da tela para o objRequisicaoCompra

Dim objRequisitante As New ClassRequisitante
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    objRequisicaoCompra.lCodigo = StrParaLong(Codigo.Text)
    objRequisicaoCompra.sDigitador = gsUsuario
    
    'Move as datas
    objRequisicaoCompra.dtData = StrParaDate(Data.Text)
    objRequisicaoCompra.dtDataLimite = StrParaDate(DataLimite.Text)
    If EnviaCompras.Value = vbChecked Then
        objRequisicaoCompra.dtDataEnvio = gdtDataHoje
    Else
        objRequisicaoCompra.dtDataEnvio = DATA_NULA
    End If
    objRequisicaoCompra.dtDataBaixa = DATA_NULA
    
    'Move Urgente
    objRequisicaoCompra.lUrgente = Urgente.Value

    If Len(Trim(Requisitante.Text)) > 0 Then
    
        'Move o requisitante
        objRequisitante.sNomeReduzido = Requisitante.Text
    
        'Lê os dados do Requisitante a partir do seu Nome Reduzido
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then Error 61980
    
        'Se não encontrou o Requisitante, erro
        If lErro = 51152 Then Error 61984
    
    End If
    
    objRequisicaoCompra.lRequisitante = objRequisitante.lCodigo
    objRequisicaoCompra.iTipoTributacao = Codigo_Extrai(TipoTributacao.Text)
    
    If Len(Trim(Ccl.ClipText)) > 0 Then
        'Move CCL
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then Error 61981
    
        objRequisicaoCompra.sCcl = sCclFormatada
    End If

    'Move Código de Ordem de Compras
    objRequisicaoCompra.sOPCodigo = CodigoOP.Text
    
    'Move a Filial Compra
    objRequisicaoCompra.iFilialCompra = Codigo_Extrai(FilialCompra.Text)

    'Move a Observacao
    objRequisicaoCompra.sObservacao = Observacao.Text

    'Move a FilialEmpresa
    objRequisicaoCompra.iFilialEmpresa = giFilialEmpresa

    'Move o Frame local de entrega
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then

        objRequisicaoCompra.iTipoDestino = TIPO_DESTINO_EMPRESA
        objRequisicaoCompra.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)

    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then

        objRequisicaoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 61982
        If lErro = 6681 Then Error 61985
                
        objRequisicaoCompra.lFornCliDestino = objFornecedor.lCodigo
        objRequisicaoCompra.iFilialDestino = Codigo_Extrai(FilialFornecedor.Text)

    End If

    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err

        Case 61980, 61981, 61982
        
        Case 61984
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO1", Err, objRequisitante.sNomeReduzido)
        
        Case 61985
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173922)

    End Select

    Exit Function

End Function

Function Move_GridItens_Memoria(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Move itens do Grid para objRequisicaoCompra

Dim lErro As Long
Dim iIndice As Integer, iCount As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCcl As String, sCclFormatada As String, iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemReqCompra As ClassItemReqCompras
Dim sContaContabil As String
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim objFornecedor As New ClassFornecedor
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemReqCompra = New ClassItemReqCompras
        
        sProduto = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError (61987)

        objItemReqCompra.sProduto = sProdutoFormatado
        objItemReqCompra.sDescProduto = GridItens.TextMatrix(iIndice, iGrid_Descricao_Col)
        objItemReqCompra.sUM = GridItens.TextMatrix(iIndice, iGrid_UM_Col)
        
        'Se a quantidade do item não foi preenchida -> Erro
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError (65599)
        
        objItemReqCompra.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) > 0 Then
        
            objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
    
            'Lê dados do almoxarifado a partir do Nome Reduzido
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25060 Then gError (61988)
    
            'Se não econtrou o almoxarifado, erro
            If lErro = 25060 Then gError (61992)
    
            objItemReqCompra.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        End If
        
        sCcl = GridItens.TextMatrix(iIndice, iGrid_CCL_Col)

        If Len(Trim(sCcl)) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError (61989)
        Else
            sCclFormatada = ""
        End If

        objItemReqCompra.sCcl = sCclFormatada

        sContaContabil = GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col)
        
        If Len(Trim(sContaContabil)) > 0 Then
            
            'Formata ContaContábil para BD
            lErro = CF("Conta_Formata", sContaContabil, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError (61990)
        
        Else
            sContaFormatada = ""
        End If
        
        objItemReqCompra.sContaContabil = sContaFormatada
        objItemReqCompra.iTipoTributacao = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col))
        
        'Move o Código do Fornecedor
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col))) > 0 Then
            
            objFornecedor.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col)
            
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError (61991)
            If lErro = 6681 Then gError (61993)
            objItemReqCompra.lFornecedor = objFornecedor.lCodigo
            
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col))) > 0 Then
                objItemReqCompra.iFilial = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col))
            End If
                        
            objFornecedorProdutoFF.lFornecedor = objItemReqCompra.lFornecedor
            objFornecedorProdutoFF.iFilialForn = objItemReqCompra.iFilial
            objFornecedorProdutoFF.sProduto = objItemReqCompra.sProduto
            objFornecedorProdutoFF.iFilialEmpresa = objRequisicaoCompra.iFilialCompra
            
            lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
            If lErro <> SUCESSO And lErro <> 54217 Then gError 86097
            If lErro <> SUCESSO Then gError 86098
            
            
        End If
        
        If GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) = "Exclusivo" Then
            objItemReqCompra.iExclusivo = 1
        Else
            objItemReqCompra.iExclusivo = 0
        End If
        
        objItemReqCompra.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)
        
        objItemReqCompra.lNumIntDoc = gColItemReqCompras.Item(iIndice)
        objRequisicaoCompra.colItens.Add objItemReqCompra
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 61987, 61988, 61989, 61990, 61991, 86097

        Case 61992
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 61993
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
                    
        Case 65599
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)
         
        Case 86098
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173923)

    End Select

    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAtual As Integer
Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iLinhaAtual = GridItens.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        gColItemReqCompras.Remove (iLinhaAtual)
        
        For iLinha = objGridPV.iLinhasExistentes To 1 Step -1
            
            If GridPV.TextMatrix(iLinha, iGrid_ItemReqPV_Col) > iLinhaAtual Then
                GridPV.TextMatrix(iLinha, iGrid_ItemReqPV_Col) = StrParaInt(GridPV.TextMatrix(iLinha, iGrid_ItemReqPV_Col)) - 1
            End If
            
            If GridPV.TextMatrix(iLinha, iGrid_ItemReqPV_Col) = iLinhaAtual Then
            
                'exclui a linha em questão
                Call Grid_Exclui_Linha(objGridPV, iLinha)
            
        
            End If
        
        Next
        

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173924)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CentroCusto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CentroCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CentroCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CentroCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CentroCusto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabil_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ContaContabil_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = ContaContabil
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoTribItem_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoTribItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub TipoTribItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub TipoTribItem_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = TipoTribItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FornecGrid_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub FornecGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub FornecGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub FornecGrid_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = FornecGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialFornecGrid_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub FilialFornecGrid_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornecGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub FilialFornecGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub FilialFornecGrid_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = FilialFornecGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Exclusivo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Exclusivo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Exclusivo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Exclusivo_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = Exclusivo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoGrid_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ObservacaoGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ObservacaoGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ObservacaoGrid_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = ObservacaoGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Requisição de Compra Não Enviada"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ReqCompras"

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

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub ObservacaoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ObservacaoLabel, Source, X, Y)
End Sub

Private Sub ObservacaoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ObservacaoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub RequisitanteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RequisitanteLabel, Source, X, Y)
End Sub

Private Sub RequisitanteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RequisitanteLabel, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
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
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is Requisitante Then
            Call RequisitanteLabel_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        ElseIf Me.ActiveControl Is Observacao Then
            Call ObservacaoLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is FornecGrid Then
            Call BotaoFiliaisFornProd_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoAlmoxarifados_Click
        ElseIf Me.ActiveControl Is CentroCusto Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is CodigoOP Then
            Call CodigoOPLabel_Click
        End If
    End If

End Sub

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

End Sub

Private Sub CodigoOPLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoOPLabel, Source, X, Y)
End Sub

Private Sub CodigoOPLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoOPLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Observacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = DescProduto

    If Len(Trim(DescProduto.Text)) = 0 Then gError 86175

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 61965

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 61965
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86175
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173925)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridPV(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Item Req.")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Pedido Venda")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant.Req.")
    objGridInt.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemReqPV.Name)
    objGridInt.colCampo.Add (ProdutoPV.Name)
    objGridInt.colCampo.Add (DescProdPV.Name)
    objGridInt.colCampo.Add (CodPV.Name)
    objGridInt.colCampo.Add (UMPV.Name)
    objGridInt.colCampo.Add (QuantReqPV.Name)
    objGridInt.colCampo.Add (QuantPV.Name)
    
    'Colunas do Grid
    iGrid_ItemReqPV_Col = 1
    iGrid_ProdutoPV_Col = 2
    iGrid_DescProdPV_Col = 3
    iGrid_CodPV_Col = 4
    iGrid_UMPV_Col = 5
    iGrid_QuantReqPV_Col = 6
    iGrid_QuantPV_Col = 7

    'Grid do GridInterno
    objGridInt.objGrid = GridPV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 23

    'Largura da primeira coluna
    GridPV.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridPV = SUCESSO

    Exit Function

End Function

Private Function Saida_Celula_ItemReqPV(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Item de Requisicao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemReqPV

    Set objGridInt.objControle = ItemReqPV

    If Len(Trim(ItemReqPV.Text)) > 0 Then

        If objGridItens.iLinhasExistentes < StrParaInt(ItemReqPV.Text) Or StrParaInt(ItemReqPV.Text) = 0 Then gError 178603
            
        If GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemReqPV_Col) <> ItemReqPV.Text Then
            
            GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoPV_Col) = GridItens.TextMatrix(StrParaInt(ItemReqPV.Text), iGrid_Produto_Col)
            GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_DescProdPV_Col) = GridItens.TextMatrix(StrParaInt(ItemReqPV.Text), iGrid_Descricao_Col)
            GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_UMPV_Col) = GridItens.TextMatrix(StrParaInt(ItemReqPV.Text), iGrid_UM_Col)
            GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantReqPV_Col) = GridItens.TextMatrix(StrParaInt(ItemReqPV.Text), iGrid_Quantidade_Col)
            GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_CodPV_Col) = ""
            GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantPV_Col) = ""

        End If

        'Se necessário cria uma nova linha no Grid
        If GridPV.Row - GridPV.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178601

    Saida_Celula_ItemReqPV = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemReqPV:

    Saida_Celula_ItemReqPV = gErr

    Select Case gErr

        Case 178601
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178603
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMREQPV_INEXISTENTE", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178602)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodPV(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Codigo de Pedido de Venda do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_Saida_Celula_CodPV

    Set objGridInt.objControle = CodPV

    'Se for o Pedido que está ativo
    If Len(Trim(CodPV.Text)) > 0 And Len(Trim(GridPV.TextMatrix(GridPV.Row, iGrid_ItemReqPV_Col))) > 0 Then
        
        lErro = CF("Produto_Formata", GridPV.TextMatrix(GridPV.Row, iGrid_ProdutoPV_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 178608
        
        objItemPedido.iFilialEmpresa = giFilialEmpresa
        objItemPedido.lCodPedido = CLng(CodPV.Text)
        objItemPedido.sProduto = sProduto
        
        'Lê o Pedido de Venda
        lErro = CF("ItemPedido_Le", objItemPedido)
        If lErro <> SUCESSO And lErro <> 23971 Then gError 178604

        'Não achou o Item do Pedido de Venda --> ERRO
        If lErro <> SUCESSO Then
        
            'Lê o Pedido de Venda Baixado
            lErro = CF("ItemPedidoBaixado_Le", objItemPedido)
            If lErro <> SUCESSO And lErro <> 178818 Then gError 178820

            If lErro <> SUCESSO Then gError 178605
        
        End If
        
        objProduto.sCodigo = sProduto
        
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 178609
    
        If lErro = 28030 Then gError 178610
        
        'Converte a quantidade para a UM de Venda
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPedido.sUnidadeMed, GridPV.TextMatrix(GridPV.Row, iGrid_UMPV_Col), dFator)
        If lErro <> SUCESSO Then gError 26640
        
        GridPV.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantPV_Col) = Formata_Estoque(objItemPedido.dQuantidade * dFator)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178606

    Saida_Celula_CodPV = SUCESSO

    Exit Function

Erro_Saida_Celula_CodPV:

    Saida_Celula_CodPV = gErr

    Select Case gErr

        Case 178604, 178606, 178608, 178609, 178820
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178605
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemPedido.lCodPedido, objItemPedido.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178610
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178607)

    End Select

    Exit Function

End Function

Private Sub GridPV_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPV, iAlterado)
    End If

End Sub

Private Sub GridPV_EnterCell()

    Call Grid_Entrada_Celula(objGridPV, iAlterado)

End Sub

Private Sub GridPV_GotFocus()

    Call Grid_Recebe_Foco(objGridPV)

End Sub

Private Sub GridPV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPV, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPV, iAlterado)
    End If

End Sub

Private Sub GridPV_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridPV)

End Sub

Private Sub GridPV_LeaveCell()

    Call Saida_Celula(objGridPV)

End Sub

Private Sub GridPV_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridPV)

End Sub

Private Sub GridPV_RowColChange()

    Call Grid_RowColChange(objGridPV)

End Sub

Private Sub GridPV_Scroll()

    Call Grid_Scroll(objGridPV)

End Sub

Public Sub ItemReqPV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ItemReqPV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPV)

End Sub

Public Sub ItemReqPV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)

End Sub

Public Sub ItemReqPV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ItemReqPV
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CodPV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodPV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPV)

End Sub

Public Sub CodPV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)

End Sub

Public Sub CodPV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CodPV
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub QuantPV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantPV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPV)

End Sub

Public Sub QuantPV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)

End Sub

Public Sub QuantPV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantPV
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_QuantPV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_QuantPV

    Set objGridInt.objControle = QuantPV

    'Se quantidade estiver preenchida
    If Len(Trim(QuantPV.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(QuantPV.Text)
        If lErro <> SUCESSO Then gError 178611

        dQuantidade = CDbl(QuantPV.Text)

        'Coloca o valor Formatado na tela
        QuantPV.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178612

    Saida_Celula_QuantPV = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantPV:

    Saida_Celula_QuantPV = gErr

    Select Case gErr

        Case 178611, 178612
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178613)

    End Select

    Exit Function

End Function

Function Move_GridPV_Memoria(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Move itens do Grid PV para objRequisicaoCompra

Dim lErro As Long
Dim objItemReqCompra As ClassItemReqCompras
Dim iIndice As Integer
Dim objItemPV As ClassItemPedido
Dim sProduto1 As String
Dim sProduto As String
Dim iPreenchido As Integer
Dim dFator As Double
Dim objProduto As New ClassProduto
Dim objItemReq As ClassItemReqCompras
Dim colItemReq As New Collection
Dim objItemReq1 As ClassItemReqCompras
Dim iAchou As Integer

On Error GoTo Erro_Move_GridPV_Memoria

    'Para cada linha do Grid
    For iIndice = 1 To objGridPV.iLinhasExistentes

        Set objItemReq = New ClassItemReqCompras

        objItemReq.iItem = StrParaInt(GridPV.TextMatrix(iIndice, iGrid_ItemReqPV_Col))
        
        objItemReq.dQuantPedida = StrParaDbl(GridPV.TextMatrix(iIndice, iGrid_QuantPV_Col))
        objItemReq.dQuantidade = StrParaDbl(GridPV.TextMatrix(iIndice, iGrid_QuantReqPV_Col))

        Set objItemReqCompra = objRequisicaoCompra.colItens(objItemReq.iItem)

        sProduto1 = GridPV.TextMatrix(iIndice, iGrid_ProdutoPV_Col)

        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 178638

        Set objItemPV = New ClassItemPedido

        objItemPV.iFilialEmpresa = giFilialEmpresa
        objItemPV.sProduto = sProduto
        objItemPV.lCodPedido = StrParaLong(GridPV.TextMatrix(iIndice, iGrid_CodPV_Col))
        
        lErro = CF("ItemPedido_Le", objItemPV)
        'Não achou o Item do Pedido de Venda --> ERRO
        If lErro <> SUCESSO Then
        
            'Lê o Pedido de Venda Baixado
            lErro = CF("ItemPedidoBaixado_Le", objItemPV)
            If lErro <> SUCESSO And lErro <> 178818 Then gError 178639

            If lErro <> SUCESSO Then gError 178605
        
        End If
        
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 178640
        
        'Se não encontrou, erro
        If lErro = 28030 Then gError 178641
                    
        'Converte a Unidade de Medida do ItemNF para a UM do Item selecionado no GridItens
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemReqCompra.sUM, objItemPV.sUnidadeMed, dFator)
        If lErro <> SUCESSO Then gError 178642
        
        objItemPV.dQuantSC = StrParaDbl(GridPV.TextMatrix(iIndice, iGrid_QuantPV_Col)) * dFator
        
        objItemReqCompra.colItemPV.Add objItemPV
                
        iAchou = 0
                
        For Each objItemReq1 In colItemReq
        
            If objItemReq1.iItem = objItemReq.iItem Then
            
                objItemReq1.dQuantPedida = objItemReq1.dQuantPedida + objItemReq.dQuantPedida
                iAchou = 1
                Exit For
        
            End If
        Next
        
        If iAchou = 0 Then colItemReq.Add objItemReq
        
    Next

    For Each objItemReq1 In colItemReq
    
        'se o total de um item em pedido de venda ultrapassar a quantidade requerida ==> erro
        If objItemReq1.dQuantPedida > objItemReq1.dQuantidade Then gError 178663
        
    Next

    Move_GridPV_Memoria = SUCESSO

    Exit Function

Erro_Move_GridPV_Memoria:

    Move_GridPV_Memoria = gErr

    Select Case gErr

        Case 178638 To 178640, 178642

        Case 178641
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 178663
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_ULTRAPASSA_REQ", gErr, objItemReq1.iItem)

        Case 178605
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemPV.lCodPedido, objItemPV.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178643)

    End Select

    Exit Function

End Function

Function Preenche_GridPV(objRequisicaoCompra As ClassRequisicaoCompras) As Long

Dim lErro As Long
Dim objItemReqCompra As ClassItemReqCompras
Dim colItensPV As Collection
Dim iIndice As Integer
Dim objItemPV As ClassItemPedido
Dim objItemPV1 As ClassItemPedido
Dim iItemReq As Integer
Dim objProduto As ClassProduto
Dim dFator As Double
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim lCodigoPV As Long
Dim iFilialPV As Integer

On Error GoTo Erro_Preenche_GridPV
    
    
    If Len(Trim(objRequisicaoCompra.sOPCodigo)) = 0 Then
    
        For Each objItemReqCompra In objRequisicaoCompra.colItens
    
            iItemReq = iItemReq + 1
    
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objItemReqCompra.sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 178660
            
            'Se não encontrou, erro
            If lErro = 28030 Then gError 178661
    
            Set colItensPV = New Collection
            
            lErro = CF("ItemPVItemRC_Le", objItemReqCompra.lNumIntDoc, colItensPV)
            If lErro <> SUCESSO Then gError 178650
            
            For Each objItemPV In colItensPV
            
                iIndice = iIndice + 1
    
                GridPV.TextMatrix(iIndice, iGrid_ItemReqPV_Col) = iItemReq
                GridPV.TextMatrix(iIndice, iGrid_ProdutoPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_Produto_Col)
                GridPV.TextMatrix(iIndice, iGrid_DescProdPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_Descricao_Col)
                
                Set objItemPV1 = New ClassItemPedido
                
                objItemPV1.iFilialEmpresa = giFilialEmpresa
                objItemPV1.lNumIntDoc = objItemPV.lNumIntDoc
                
                lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV1)
                If lErro <> SUCESSO And lErro <> 178567 Then gError 178671
                
                GridPV.TextMatrix(iIndice, iGrid_CodPV_Col) = objItemPV1.lCodPedido
                GridPV.TextMatrix(iIndice, iGrid_UMPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_UM_Col)
                GridPV.TextMatrix(iIndice, iGrid_QuantReqPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_Quantidade_Col)
                
                'Converte a Unidade de Medida do ItemNF para a UM do Item selecionado no GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPV1.sUnidadeMed, objItemReqCompra.sUM, dFator)
                If lErro <> SUCESSO Then gError 178662
                
                GridPV.TextMatrix(iIndice, iGrid_QuantPV_Col) = Formata_Estoque(objItemPV.dQuantSC * dFator)
                
            Next
    
        Next

    Else
    
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = objRequisicaoCompra.sOPCodigo
    
        lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30401 Then gError 178698

        If lErro <> SUCESSO Then
        
            lErro = CF("ItensOP_Baixada_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 178689 Then gError 178699
        
        End If
        
        If lErro = SUCESSO Then
        
            For Each objItemOP In objOrdemProducao.colItens
                
                If objItemOP.lCodPedido <> 0 Then
                    CodigoPV.Caption = objItemOP.lCodPedido
                    Exit For
                End If
                
                If objItemOP.lNumIntDocPai <> 0 Then
                
                    lErro = CF("ItensOP_Le_PV", objItemOP.lNumIntDocPai, lCodigoPV, iFilialPV)
                    If lErro <> SUCESSO And lErro <> 178696 And lErro <> 178697 Then gError 178700
            
                End If
            
                If lCodigoPV <> 0 Then
                    CodigoPV.Caption = lCodigoPV
                    Exit For
                End If
            
            Next
    
        End If
    
    End If
    
    objGridPV.iLinhasExistentes = iIndice

    Preenche_GridPV = SUCESSO

    Exit Function

Erro_Preenche_GridPV:

    Preenche_GridPV = gErr

    Select Case gErr

        Case 178650, 178660, 178662, 178671, 178698 To 178700

        Case 178661
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178704)

    End Select

    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long

Dim lErro As Long
Dim sRequisitante As String
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
                        
        Case CodigoRequisicao
            objMnemonicoValor.colValor.Add objContexto.lCodigo
            
        Case FILIAL_EMPRESA
            objMnemonicoValor.colValor.Add objContexto.iFilialEmpresa
              
        Case Else
            If Not (objContexto Is Nothing) Then
                lErro = objContexto.Calcula_Mnemonico(objMnemonicoValor, objContexto)
                If lErro <> SUCESSO Then gError 178683
            Else
                gError 178683
            End If

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 178683
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178684)

    End Select

    Exit Function

End Function

