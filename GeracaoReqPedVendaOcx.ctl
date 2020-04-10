VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoReqPedVendaOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8445
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   16665
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Pedidos"
         Height          =   5610
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   7110
         Begin VB.Frame FrameEntrega 
            Caption         =   "Data Entrega"
            Height          =   825
            Left            =   435
            TabIndex        =   21
            Top             =   4245
            Width           =   5520
            Begin MSMask.MaskEdBox DataEntregaDe 
               Height          =   300
               Left            =   705
               TabIndex        =   23
               Top             =   375
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEntregaDe 
               Height          =   300
               Left            =   1875
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   375
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEntregaAte 
               Height          =   300
               Left            =   3420
               TabIndex        =   26
               Top             =   375
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEntregaAte 
               Height          =   300
               Left            =   4590
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   375
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label8 
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
               Index           =   2
               Left            =   2985
               TabIndex        =   25
               Top             =   420
               Width           =   360
            End
            Begin VB.Label Label7 
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
               Height          =   195
               Index           =   0
               Left            =   300
               TabIndex        =   22
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.CheckBox ExibeTodos 
            Caption         =   "Exibe Todos os Pedidos"
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
            Left            =   1560
            TabIndex        =   3
            Top             =   390
            Width           =   2430
         End
         Begin VB.Frame FrameCliente 
            Caption         =   "Clientes"
            Height          =   825
            Left            =   435
            TabIndex        =   9
            Top             =   1950
            Width           =   5520
            Begin MSMask.MaskEdBox ClienteDe 
               Height          =   300
               Left            =   705
               TabIndex        =   11
               Top             =   375
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteAte 
               Height          =   300
               Left            =   3420
               TabIndex        =   13
               Top             =   375
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelClienteDe 
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
               Left            =   300
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   10
               Top             =   420
               Width           =   315
            End
            Begin VB.Label LabelClienteAte 
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
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   12
               Top             =   390
               Width           =   360
            End
         End
         Begin VB.Frame FramePedido 
            Caption         =   "Pedidos"
            Height          =   810
            Left            =   435
            TabIndex        =   4
            Top             =   840
            Width           =   5520
            Begin MSMask.MaskEdBox PedidoInicial 
               Height          =   300
               Left            =   720
               TabIndex        =   6
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PedidoFinal 
               Height          =   300
               Left            =   3420
               TabIndex        =   8
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelPedidoAte 
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
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   7
               Top             =   405
               Width           =   360
            End
            Begin VB.Label LabelPedidoDe 
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
               Left            =   300
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   5
               Top             =   405
               Width           =   315
            End
         End
         Begin VB.Frame FrameEmissao 
            Caption         =   "Data Emissão"
            Height          =   825
            Left            =   435
            TabIndex        =   14
            Top             =   3090
            Width           =   5520
            Begin MSMask.MaskEdBox DataEmissaoDe 
               Height          =   300
               Left            =   705
               TabIndex        =   16
               Top             =   345
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoDe 
               Height          =   300
               Left            =   1875
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissaoAte 
               Height          =   300
               Left            =   3420
               TabIndex        =   19
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoAte 
               Height          =   300
               Left            =   4590
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label23 
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
               Height          =   195
               Left            =   300
               TabIndex        =   15
               Top             =   420
               Width           =   315
            End
            Begin VB.Label Label24 
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
               Left            =   2985
               TabIndex        =   18
               Top             =   420
               Width           =   360
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8535
      Index           =   2
      Left            =   165
      TabIndex        =   28
      Top             =   510
      Visible         =   0   'False
      Width           =   16620
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   675
         Left            =   2010
         Picture         =   "GeracaoReqPedVendaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   7695
         Width           =   1425
      End
      Begin VB.CommandButton BotaoPedido 
         Caption         =   "Editar Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5640
         Picture         =   "GeracaoReqPedVendaOcx.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   7695
         Width           =   1425
      End
      Begin VB.TextBox NomeRed 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4290
         TabIndex        =   34
         Text            =   "Nome"
         Top             =   3480
         Width           =   5000
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   225
         Left            =   5550
         TabIndex        =   36
         Top             =   4845
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEntrega 
         Height          =   225
         Left            =   5955
         TabIndex        =   37
         Top             =   5565
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4935
         TabIndex        =   33
         Top             =   3960
         Width           =   1305
      End
      Begin VB.CheckBox Selecionado 
         DragMode        =   1  'Automatic
         Height          =   210
         Left            =   5685
         TabIndex        =   32
         Top             =   6120
         Width           =   1350
      End
      Begin VB.TextBox FilialCliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4785
         TabIndex        =   35
         Top             =   4275
         Width           =   2295
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "GeracaoReqPedVendaOcx.ctx":1E60
         Left            =   1560
         List            =   "GeracaoReqPedVendaOcx.ctx":1E62
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   263
         Width           =   4470
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   675
         Left            =   90
         Picture         =   "GeracaoReqPedVendaOcx.ctx":1E64
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   7695
         Width           =   1425
      End
      Begin MSFlexGridLib.MSFlexGrid GridPedido 
         Height          =   6750
         Left            =   150
         TabIndex        =   31
         Top             =   900
         Width           =   16170
         _ExtentX        =   28522
         _ExtentY        =   11906
         _Version        =   393216
         Rows            =   10
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label20 
         Caption         =   "Ordenados por:"
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
         Left            =   150
         TabIndex        =   29
         Top             =   300
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8550
      Index           =   5
      Left            =   180
      TabIndex        =   79
      Top             =   495
      Visible         =   0   'False
      Width           =   16545
      Begin VB.Frame Frame3 
         Caption         =   "Dados Principais"
         Height          =   1620
         Left            =   195
         TabIndex        =   80
         Top             =   360
         Width           =   8295
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
            TabIndex        =   85
            Top             =   465
            Width           =   2040
         End
         Begin VB.ComboBox FilialCompra 
            Height          =   315
            Left            =   1695
            TabIndex        =   84
            Top             =   930
            Width           =   2610
         End
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   315
            Left            =   1710
            TabIndex        =   83
            Top             =   390
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
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
            TabIndex        =   82
            Top             =   975
            Width           =   1155
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
            Left            =   435
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   81
            Top             =   420
            Width           =   1140
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Local de Entrega"
         Height          =   1365
         Left            =   195
         TabIndex        =   86
         Top             =   2145
         Width           =   8295
         Begin VB.Frame Frame2 
            Caption         =   "Tipo"
            Height          =   585
            Index           =   1
            Left            =   225
            TabIndex        =   87
            Top             =   390
            Width           =   3870
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
               TabIndex        =   89
               Top             =   225
               Width           =   1335
            End
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
               Left            =   405
               TabIndex        =   88
               Top             =   225
               Value           =   -1  'True
               Width           =   1635
            End
         End
         Begin VB.Frame FrameTipoDestino 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   675
            Index           =   0
            Left            =   4500
            TabIndex        =   90
            Top             =   345
            Width           =   3645
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               Left            =   1260
               TabIndex        =   92
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
               TabIndex        =   91
               Top             =   255
               Width           =   465
            End
         End
         Begin VB.Frame FrameTipoDestino 
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   4515
            TabIndex        =   93
            Top             =   345
            Visible         =   0   'False
            Width           =   3645
            Begin VB.ComboBox FilialFornec 
               Height          =   315
               Left            =   1230
               TabIndex        =   96
               Top             =   345
               Width           =   2160
            End
            Begin MSMask.MaskEdBox Fornec 
               Height          =   315
               Left            =   1230
               TabIndex        =   102
               Top             =   0
               Width           =   2160
               _ExtentX        =   3810
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
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
               TabIndex        =   94
               Top             =   60
               Width           =   1035
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
               TabIndex        =   95
               Top             =   405
               Width           =   465
            End
         End
      End
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
         Left            =   195
         TabIndex        =   97
         Top             =   3810
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8505
      Index           =   4
      Left            =   120
      TabIndex        =   53
      Top             =   495
      Visible         =   0   'False
      Width           =   16680
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
         Left            =   3075
         TabIndex        =   76
         Top             =   8190
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
         Left            =   1215
         TabIndex        =   75
         Top             =   8190
         Width           =   1710
      End
      Begin VB.CommandButton BotaoProdutoFiliaisForn 
         Caption         =   "Fornecedor x Produto"
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
         Left            =   6555
         TabIndex        =   78
         Top             =   8190
         Width           =   2130
      End
      Begin VB.CommandButton BotaoEstoqueProd 
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
         Left            =   4920
         TabIndex        =   77
         Top             =   8190
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
         Left            =   75
         TabIndex        =   74
         Top             =   8190
         Width           =   1005
      End
      Begin VB.Frame FrameProdutos 
         Caption         =   "Produtos"
         Height          =   7815
         Left            =   150
         TabIndex        =   54
         Top             =   300
         Width           =   16425
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   225
            Left            =   4275
            TabIndex        =   101
            Top             =   3120
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   180
            TabIndex        =   100
            Top             =   3150
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   "_"
         End
         Begin VB.TextBox FilialForn 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5415
            MaxLength       =   50
            TabIndex        =   71
            Top             =   3465
            Width           =   1320
         End
         Begin MSMask.MaskEdBox QuantReservada 
            Height          =   225
            Left            =   6105
            TabIndex        =   61
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
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
            Left            =   1515
            TabIndex        =   69
            Top             =   3585
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CentroCusto 
            Height          =   225
            Left            =   795
            TabIndex        =   68
            Top             =   3705
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2595
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   255
            Width           =   885
         End
         Begin MSMask.MaskEdBox QuantEmPedido 
            Height          =   225
            Left            =   135
            TabIndex        =   63
            Top             =   840
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
         Begin VB.ComboBox TipoTribItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   3540
            Width           =   2415
         End
         Begin VB.ComboBox Exclusivo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "GeracaoReqPedVendaOcx.ctx":2E7E
            Left            =   6600
            List            =   "GeracaoReqPedVendaOcx.ctx":2E88
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   3630
            Width           =   1305
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   7635
            MaxLength       =   255
            TabIndex        =   73
            Top             =   3645
            Width           =   2115
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   57
            Top             =   390
            Width           =   4000
         End
         Begin MSMask.MaskEdBox QuantEmReq 
            Height          =   225
            Left            =   1560
            TabIndex        =   64
            Top             =   855
            Width           =   1650
            _ExtentX        =   2910
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
         Begin MSMask.MaskEdBox SaldoAtual 
            Height          =   225
            Left            =   3060
            TabIndex        =   65
            Top             =   855
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
            Left            =   5130
            TabIndex        =   67
            Top             =   840
            Width           =   1605
            _ExtentX        =   2831
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
            Left            =   -165
            TabIndex        =   56
            Top             =   405
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
            Left            =   3870
            TabIndex        =   66
            Top             =   870
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
         Begin MSMask.MaskEdBox QuantComprar 
            Height          =   225
            Left            =   3705
            TabIndex        =   59
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSMask.MaskEdBox QuantEstoque 
            Height          =   225
            Left            =   4965
            TabIndex        =   60
            Top             =   210
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
         Begin MSMask.MaskEdBox QuantEmpenhada 
            Height          =   225
            Left            =   7080
            TabIndex        =   62
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
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
         Begin MSFlexGridLib.MSFlexGrid GridProdutos 
            Height          =   3555
            Left            =   180
            TabIndex        =   55
            Top             =   255
            Width           =   16065
            _ExtentX        =   28337
            _ExtentY        =   6271
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   8490
      Index           =   3
      Left            =   165
      TabIndex        =   39
      Top             =   510
      Visible         =   0   'False
      Width           =   16620
      Begin VB.CommandButton BotaoEntrega 
         Caption         =   "Datas de Entrega"
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
         TabIndex        =   106
         Top             =   8085
         Width           =   1845
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens de Pedidos de Venda"
         Height          =   7995
         Index           =   3
         Left            =   90
         TabIndex        =   40
         Top             =   105
         Width           =   16425
         Begin MSMask.MaskEdBox QuantSC 
            Height          =   240
            Left            =   3780
            TabIndex        =   103
            Top             =   450
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   423
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
         Begin VB.CheckBox SelItem 
            DragMode        =   1  'Automatic
            Height          =   210
            Left            =   45
            TabIndex        =   99
            Top             =   495
            Width           =   1050
         End
         Begin MSMask.MaskEdBox QuantFaturada 
            Height          =   240
            Left            =   6840
            TabIndex        =   52
            Top             =   150
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   423
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
         Begin MSMask.MaskEdBox QuantReservadaPV 
            Height          =   240
            Left            =   6210
            TabIndex        =   51
            Top             =   330
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   423
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
         Begin MSMask.MaskEdBox QuantCancelada 
            Height          =   255
            Left            =   5220
            TabIndex        =   50
            Top             =   135
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantidadeItem 
            Height          =   240
            Left            =   4590
            TabIndex        =   49
            Top             =   315
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   423
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
         Begin MSMask.MaskEdBox QuantComprarItem 
            Height          =   240
            Left            =   3630
            TabIndex        =   48
            Top             =   165
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "GeracaoReqPedVendaOcx.ctx":2EA5
            Left            =   2490
            List            =   "GeracaoReqPedVendaOcx.ctx":2EA7
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   150
            Width           =   720
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1665
            MaxLength       =   50
            TabIndex        =   45
            Top             =   315
            Width           =   4000
         End
         Begin VB.TextBox PedidoItem 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   135
            TabIndex        =   42
            Top             =   270
            Width           =   750
         End
         Begin VB.TextBox ItemPed 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   1125
            TabIndex        =   43
            Top             =   540
            Width           =   600
         End
         Begin MSMask.MaskEdBox DataEntregaItem 
            Height          =   225
            Left            =   3120
            TabIndex        =   47
            Top             =   300
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
         Begin MSMask.MaskEdBox ProdutoItem 
            Height          =   240
            Left            =   930
            TabIndex        =   44
            Top             =   255
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   7380
            Left            =   135
            TabIndex        =   41
            Top             =   525
            Width           =   16125
            _ExtentX        =   28443
            _ExtentY        =   13018
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
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
      Left            =   15660
      Picture         =   "GeracaoReqPedVendaOcx.ctx":2EA9
      Style           =   1  'Graphical
      TabIndex        =   98
      ToolTipText     =   "Fechar"
      Top             =   75
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8910
      Left            =   75
      TabIndex        =   0
      Top             =   180
      Width           =   16830
      _ExtentX        =   29686
      _ExtentY        =   15716
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos de Venda"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens de Pedido de Venda"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "GeracaoReqPedVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Dim iFrameSelecaoAlterado As Integer
Dim iFrameAtual As Integer
Dim giTipoTributacao As Integer
Dim iFrameAlterado As Integer
Dim giExibeTodos As Integer

Dim iFrameTipoDestinoAtual As Integer

Dim iFornecedorAlterado As Integer

Dim gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda

Const TAB_Selecao = 1

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1
Private WithEvents objEventoPedidoDe As AdmEvento
Attribute objEventoPedidoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedidoAte As AdmEvento
Attribute objEventoPedidoAte.VB_VarHelpID = -1
Private WithEvents objEventoBotaoFiliaisFornProduto As AdmEvento
Attribute objEventoBotaoFiliaisFornProduto.VB_VarHelpID = -1
Private WithEvents objEventoBotaoEstoqueProduto As AdmEvento
Attribute objEventoBotaoEstoqueProduto.VB_VarHelpID = -1
Private WithEvents objEventoRequisitante As AdmEvento
Attribute objEventoRequisitante.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoBotaoCcl As AdmEvento
Attribute objEventoBotaoCcl.VB_VarHelpID = -1

'GridPedidos
Dim objGridPedidos As AdmGrid
Dim iGrid_SelecionadoPed_Col As Integer
Dim iGrid_Pedido_Col As Integer
Dim iGrid_NomeRed_Col As Integer
Dim iGrid_FilialCliente_Col As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_DataEntrega_Col As Integer

'GridItens
Dim objGridItens As New AdmGrid
Dim iGrid_SelItem_Col As Integer
Dim iGrid_PedidoItem_Col As Integer
Dim iGrid_ItemPed_Col As Integer
Dim iGrid_ProdutoItem_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_DataEntregaItem_Col As Integer
Dim iGrid_QuantidadeItem_Col As Integer
Dim iGrid_QuantCancelada_Col As Integer
Dim iGrid_QuantReservadaPV_Col As Integer
Dim iGrid_QuantFaturada_Col As Integer
Dim iGrid_QuantComprarItem_Col As Integer
Dim iGrid_QuantSC_Col As Integer

'GridProdutos
Dim objGridProdutos As New AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_QuantComprar_Col As Integer
Dim iGrid_QuantEstoque_Col As Integer
Dim iGrid_QuantReservada_Col As Integer
Dim iGrid_QuantEmpenhada_Col As Integer
Dim iGrid_QuantEmPedido_Col As Integer
Dim iGrid_QuantEmReq_Col As Integer
Dim iGrid_SaldoAtual_Col As Integer
Dim iGrid_PontoPedido_Col As Integer
Dim iGrid_LoteEconomico_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_CentroCusto_Col As Integer
Dim iGrid_ContaContabil_Col As Integer
Dim iGrid_TipoTribItem_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialForn_Col As Integer
Dim iGrid_Exclusivo_Col As Integer
Dim iGrid_Observacao_Col As Integer

Dim iClienteAlterado As Integer
Dim gsOrdenacao As String
Dim asOrdenacao(2) As String
Dim asOrdenacaoString(2) As String

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Pedidos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Pedidos

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Selecionado")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial Cliente")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Entrega")

    'campos de edição do grid
    objGridInt.colCampo.Add (Selecionado.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (NomeRed.Name)
    objGridInt.colCampo.Add (FilialCliente.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_SelecionadoPed_Col = 1
    iGrid_Pedido_Col = 2
    iGrid_NomeRed_Col = 3
    iGrid_FilialCliente_Col = 4
    iGrid_DataEmissao_Col = 5
    iGrid_DataEntrega_Col = 6

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridPedido

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PEDIDOS_SEL + 1

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 25

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridPedido.Width = 7900

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Pedidos = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 68262
    
    'Pesquisa controle da coluna em questão
    Select Case objControl.Name
                                
        'Produto
        Case Produto.Name
            If Len(Trim(GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If

        'Nas demais
        Case QuantComprar.Name, Observacao.Name

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case ContaContabil.Name

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            ElseIf iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            End If

        Case FilialForn.Name, Exclusivo.Name
            If iProdutoPreenchido = PRODUTO_VAZIO Or Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case QuantComprarItem.Name
            If StrParaInt(GridItens.TextMatrix(iLinha, iGrid_SelItem_Col)) <> MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Almoxarifado.Name, CentroCusto.Name

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

        Case 68262
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161427)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Selecionado")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Entrega")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Quant Cancelada")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Requisitada")
    objGridInt.colColuna.Add ("Quant Faturada")
    objGridInt.colColuna.Add ("Quant Comprar")

    'campos de edição do grid
    objGridInt.colCampo.Add (SelItem.Name)
    objGridInt.colCampo.Add (PedidoItem.Name)
    objGridInt.colCampo.Add (ItemPed.Name)
    objGridInt.colCampo.Add (ProdutoItem.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (DataEntregaItem.Name)
    objGridInt.colCampo.Add (QuantidadeItem.Name)
    objGridInt.colCampo.Add (QuantCancelada.Name)
    objGridInt.colCampo.Add (QuantReservadaPV.Name)
    objGridInt.colCampo.Add (QuantSC.Name)
    objGridInt.colCampo.Add (QuantFaturada.Name)
    objGridInt.colCampo.Add (QuantComprarItem.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_SelItem_Col = 1
    iGrid_PedidoItem_Col = 2
    iGrid_ItemPed_Col = 3
    iGrid_ProdutoItem_Col = 4
    iGrid_DescricaoProduto_Col = 5
    iGrid_UM_Col = 6
    iGrid_DataEntregaItem_Col = 7
    iGrid_QuantidadeItem_Col = 8
    iGrid_QuantCancelada_Col = 9
    iGrid_QuantReservadaPV_Col = 10
    iGrid_QuantSC_Col = 11
    iGrid_QuantFaturada_Col = 12
    iGrid_QuantComprarItem_Col = 13

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENSPED_SEL + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridItens.Width = 7500

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO
    
    Exit Function

End Function

Private Function Inicializa_GridProdutos(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U. M.")
    objGridInt.colColuna.Add ("Quant Comprar")
    objGridInt.colColuna.Add ("Quant E que")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Empenhada")
    objGridInt.colColuna.Add ("Quant em Pedido")
    objGridInt.colColuna.Add ("Quant em Requisicao")
    objGridInt.colColuna.Add ("Saldo Atual")
    objGridInt.colColuna.Add ("Ponto Pedido")
    objGridInt.colColuna.Add ("Lote Econômico")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil")
    objGridInt.colColuna.Add ("Tipo de Tributação")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusividade")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (QuantComprar.Name)
    objGridInt.colCampo.Add (QuantEstoque.Name)
    objGridInt.colCampo.Add (QuantReservada.Name)
    objGridInt.colCampo.Add (QuantEmpenhada.Name)
    objGridInt.colCampo.Add (QuantEmPedido.Name)
    objGridInt.colCampo.Add (QuantEmReq.Name)
    objGridInt.colCampo.Add (SaldoAtual.Name)
    objGridInt.colCampo.Add (PontoPedido.Name)
    objGridInt.colCampo.Add (LoteEconomico.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)
    objGridInt.colCampo.Add (TipoTribItem.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (FilialForn.Name)
    objGridInt.colCampo.Add (Exclusivo.Name)
    objGridInt.colCampo.Add (Observacao.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_QuantComprar_Col = 4
    iGrid_QuantEstoque_Col = 5
    iGrid_QuantReservada_Col = 6
    iGrid_QuantEmpenhada_Col = 7
    iGrid_QuantEmPedido_Col = 8
    iGrid_QuantEmReq_Col = 9
    iGrid_SaldoAtual_Col = 10
    iGrid_PontoPedido_Col = 11
    iGrid_LoteEconomico_Col = 12
    iGrid_Almoxarifado_Col = 13
    iGrid_CentroCusto_Col = 14
    iGrid_ContaContabil_Col = 15
    iGrid_TipoTribItem_Col = 16
    iGrid_Fornecedor_Col = 17
    iGrid_FilialForn_Col = 18
    iGrid_Exclusivo_Col = 19
    iGrid_Observacao_Col = 20

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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 68196

            'se for o GridProdutos
            Case GridProdutos.Name

                lErro = Saida_Celula_GridProdutos(objGridInt)
                If lErro <> SUCESSO Then gError 68197

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 68198

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 68196 To 68198
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161428)

    End Select

    Exit Function

End Function
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

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Unidade de Medida
        Case iGrid_UM_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 68199

        'QuantComprar
        Case iGrid_QuantComprarItem_Col
            lErro = Saida_Celula_QuantComprarItem(objGridInt)
            If lErro <> SUCESSO Then gError 68200

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 68199, 68200

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161429)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridProdutos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProdutos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'QuantComprar
        Case iGrid_QuantComprar_Col
            lErro = Saida_Celula_QuantComprar(objGridInt)
            If lErro <> SUCESSO Then gError 68203

        'Almoxarifado
        Case iGrid_Almoxarifado_Col
            lErro = Saida_Celula_Almoxarifado(objGridInt)
            If lErro <> SUCESSO Then gError 68204

        'CentroCusto
        Case iGrid_CentroCusto_Col
            lErro = Saida_Celula_CentroCusto(objGridInt)
            If lErro <> SUCESSO Then gError 68205

        'ContaContabil
        Case iGrid_ContaContabil_Col
            lErro = Saida_Celula_ContaContabil(objGridInt)
            If lErro <> SUCESSO Then gError 68206

        'TipoTribItem
        Case iGrid_TipoTribItem_Col
            lErro = Saida_Celula_TipoTribItem(objGridInt)
            If lErro <> SUCESSO Then gError 68207

        'Fornecedor
        Case iGrid_Fornecedor_Col
            lErro = Saida_Celula_Fornecedor(objGridInt)
            If lErro <> SUCESSO Then gError 68208

        'FilialForn
        Case iGrid_FilialForn_Col
            lErro = Saida_Celula_FilialForn(objGridInt)
            If lErro <> SUCESSO Then gError 68209

        'Exclusivo
        Case iGrid_Exclusivo_Col
            lErro = Saida_Celula_Exclusivo(objGridInt)
            If lErro <> SUCESSO Then gError 68210

        'Observacao
        Case iGrid_Observacao_Col
            lErro = Saida_Celula_Observacao(objGridInt)
            If lErro <> SUCESSO Then gError 68211

        'UnidadeMed
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UnidadeMed(objGridInt)
            If lErro <> SUCESSO Then gError 68264

    End Select

    Saida_Celula_GridProdutos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridProdutos:

    Saida_Celula_GridProdutos = gErr

    Select Case gErr

        Case 68203 To 68211

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161430)

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
                
        'Verifica se o Fornecedor foi preenchido
        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col))) = 0 Then gError 68279
        
        iCodigo = Codigo_Extrai(FilialForn.Text)
        
        'Coloca o Produto no formato do BD
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 68280

        sFornecedor = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col)
        objFornecedorProdutoFF.iFilialForn = iCodigo
        objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa
        objFornecedorProdutoFF.sProduto = sProdutoFormatado

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FornecedorProdutoFF_Le_NomeRed", sFornecedor, sNomeFilial, objFornecedorProdutoFF)
        If lErro <> SUCESSO And lErro <> 61780 Then gError 68281

        'Se não encontrou a Filial do Fornecedor
        If lErro = 61780 Then

            'Lê FilialFornecedor do BD
            objFilialFornecedor.iCodFilial = iCodigo
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 68282

            'Se não encontrou, pergunta se deseja criar
            If lErro = 18272 Then
                gError 68283

            'Se encontrou, erro
            Else
                gError 68284
            End If

        'Se encontrou a Filial do Fornecedor
        Else

            'coloca na tela
            FilialForn.Text = CStr(iCodigo) & SEPARADOR & sNomeFilial

        End If

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 68285
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68286

    Saida_Celula_FilialForn = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialForn:

    Saida_Celula_FilialForn = gErr

    Select Case gErr

        Case 68280, 68281, 68282, 68286
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68279
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_FORNECEDOR_NAO_PREENCHIDO", gErr, GridItens.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68283

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

        Case 68285
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_NAO_ENCONTRADA_ASSOCIADA", gErr, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68284
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", gErr, objFilialFornecedor.iCodFilial, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161431)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoTribItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo de Tributação que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TipoTribItem

    Set objGridInt.objControle = TipoTribItem

    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TipoTribItem_Col) = TipoTribItem.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68244

    Saida_Celula_TipoTribItem = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoTribItem:

    Saida_Celula_TipoTribItem = gErr

    Select Case gErr

        Case 68244
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161432)

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

    sFornecedor = Fornecedor.Text

    'Se o fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Verifica se o fornecedor está cadastrado
        lErro = TP_Fornecedor_Grid(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 25611 And lErro <> 25613 And lErro <> 25616 And lErro <> 25619 Then gError 68245

        'Fornecedor não cadastrado
        'Nome Reduzido
        If lErro = 25611 Then gError 68246

        'Codigo
        If lErro = 25613 Then gError 68247

        'CGC/CPF
        If lErro = 25616 Or lErro = 25619 Then gError 68248

        If sFornecedor <> objFornecedor.sNomeReduzido Then

            'Formata o Produto
            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 68249

            iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)

            'Lê coleção de códigos e nomes da Filial do Fornecedor
            lErro = CF("FornecedorProdutoFF_Le_FilialForn", sProdutoFormatado, objFornecedor.lCodigo, iFilialEmpresa, colCodigoNome)
            If lErro <> SUCESSO Then gError 68250

            'Se não encontrou nenhuma Filial, erro
            If colCodigoNome.Count = 0 Then gError 68251

            If iCodFilial > 0 Then

                For iIndice = 1 To colCodigoNome.Count
                    If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                        Exit For
                    End If
                Next

                If iIndice = colCodigoNome.Count Then gError 68252

            ElseIf iCodFilial = 0 Then
                iCodFilial = colCodigoNome.Item(1).iCodigo
            End If

            For iIndice = 1 To colCodigoNome.Count
                If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_FilialForn_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
                    Exit For
                End If
            Next

        End If

        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col))) = 0 Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col) = "Preferencial"
        End If

    Else

        'Limpa a Filial e Exclusividade Correspondente
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_FilialForn_Col) = ""
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Exclusivo_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68253

    Saida_Celula_Fornecedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Fornecedor:

    Saida_Celula_Fornecedor = gErr

    Select Case gErr

        Case 68245, 68253, 68249, 68250
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68246 'Fornecedor com Nome Reduzido %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", Fornecedor.Text)
            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 68247 'Fornecedor com código %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 68248 'Fornecedor com CGC/CPF %s não encontado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_3", Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 68251
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FIL_PROD_FORN_FILIALCOMPRA", gErr, objFornecedor.sNomeReduzido, sProdutoFormatado)
            Fornecedor.Text = ""
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68252
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", gErr, iCodFilial, objFornecedor.sNomeReduzido, sProdutoFormatado)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161433)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Observacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68243

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 68243
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161434)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Exclusivo(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Exclusivo do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Exclusivo

    Set objGridInt.objControle = Exclusivo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68242

    Saida_Celula_Exclusivo = SUCESSO

    Exit Function

Erro_Saida_Celula_Exclusivo:

    Saida_Celula_Exclusivo = gErr

    Select Case gErr

        Case 68242
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161435)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Almoxarifado do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado está preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        'Formata o produto
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 84546 '68228
      
        'Valida o Almoxarifado
        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 84547 '68229
        
        'Se não for encontrado --> Erro
        If lErro = 25130 Then gError 84548 '68230
        If lErro = 25135 Then gError 84549 '68231

        If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 84550 '86099

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84551 '68232

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 84546, 84547, 84551
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84548
            'Pergunta de deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)
            'Se a resposta for sim
            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                'Chama a Tela Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 84549

            'Pergunta se deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(Almoxarifado.Text))
            'Se a resposta for positiva
            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = Codigo_Extrai(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

         Case 84550
            
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161436)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CentroCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula de CentroCusto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_CentroCusto

    Set objGridInt.objControle = CentroCusto

    'Verifica se CentroCusto foi preenchido
    If Len(Trim(CentroCusto.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", CentroCusto, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 68233

        'Se o Ccl nao existe
        If lErro = 5703 Then gError 68234

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68235

    Saida_Celula_CentroCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_CentroCusto:

    Saida_Celula_CentroCusto = gErr

    Select Case gErr

        Case 68233
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68234
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CentroCusto.Text)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68235
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161437)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long
'Faz a critica da celula de ContaContabil do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sContaSimples As String
Dim objPlanoConta As New ClassPlanoConta
Dim sContaEnxuta As String
Dim sContaFormatada As String
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_Saida_Celula_ContaContabil

    Set objGridInt.objControle = ContaContabil

    'Verifica se a conta contabil esta preenchida
    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'Critica a conta informada
        lErro = CF("ContaSimples_Critica", sContaSimples, ContaContabil.ClipText, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 44037 Then gError 68236

        'Se for uma conta simples
        If lErro = SUCESSO Then

            lErro = Mascara_RetornaContaEnxuta(sContaSimples, sContaEnxuta)
            If lErro <> SUCESSO Then gError 68237

            'Coloca conta enxuta na tela
            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True
            
        End If

        'Se a conta nao e uma conta simples
        If lErro = 44037 Then

            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_COMPRAS)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 68238

            'Nao encontrou a conta
            If lErro = 5700 Then gError 68239

        End If

        'Mascara a Conta Contabil
        lErro = Mascara_MascararConta(sContaFormatada, sContaMascarada)
        If lErro <> SUCESSO Then gError 68240

        'Coloca a descricao da conta na tela
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68241

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = gErr

    Select Case gErr

        Case 68236, 68237, 68238, 68240, 68241
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68239
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161438)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a critica da celula de UM do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UM
    
    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UM_Col) = UM.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68201

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 68201
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161439)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a critica da celula de UnidadeMed do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UM_Col) = UnidadeMed.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68265

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 68265
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161440)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantComprar(objGridInt As AdmGrid) As Long
'Faz a critica da celula de QuantComprar do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantComprar

    Set objGridInt.objControle = QuantComprar

    If Len(Trim(QuantComprar.Text)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(QuantComprar.Text)
        If lErro <> SUCESSO Then gError 68226
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68227

    Saida_Celula_QuantComprar = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprar:

    Saida_Celula_QuantComprar = gErr

    Select Case gErr

        Case 68226, 68227
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161441)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantComprarItem(objGridInt As AdmGrid) As Long
'Faz a critica da celula de QuantComprarItem do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantAnterior As Double
Dim bQuantidadeIgual As Boolean
Dim dQuantidade As Double
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Saida_Celula_QuantComprarItem

    Set objGridInt.objControle = QuantComprarItem

    If Len(Trim(QuantComprarItem.Text)) > 0 Then
        
        'Faz a crítica do valor informado
        lErro = Valor_NaoNegativo_Critica(QuantComprarItem.Text)
        If lErro <> SUCESSO Then gError 68266
        
        'Guarda a quantidade anterior
        dQuantAnterior = StrParaDbl(QuantComprarItem.Text)
        dQuantidade = StrParaDbl(QuantComprarItem.Text)

        'Coloca o valor Formatado na tela
        QuantComprarItem.Text = Formata_Estoque(dQuantidade)

    End If

    'Verifica se a quantidade foi alterada
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantComprarItem_Col)) = StrParaDbl(QuantComprarItem.Text) Then bQuantidadeIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 68271
    
    'Se quantidade foi alterada
    If bQuantidadeIgual = False Then

        'Formata o produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItem_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 68267

        'Guarda o Codigo do Produto em objProduto
        objProduto.sCodigo = sProduto

        'Le o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 68268
        'Se nao encontrou => erro
        If lErro = 28030 Then gError 68269

        'Reflete no GridProdutos
        lErro = GridProdutos_Atualiza(GridItens.Row, objProduto, dQuantidade)
        If lErro <> SUCESSO Then gError 68270

    End If

    Saida_Celula_QuantComprarItem = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprarItem:

    Saida_Celula_QuantComprarItem = gErr

    Select Case gErr

        Case 68202, 68266, 68267, 68268, 68270, 68271
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68269
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161442)

    End Select

    Exit Function

End Function
Function GridProdutos_Atualiza(iItem As Integer, objProduto As ClassProduto, dQuantidade As Double) As Long
'Atualiza a Quantidade a Comprar para o Produto do GridProdutos

Dim lErro As Long
Dim dFator As Double
Dim dQuantItem As Double
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dQuantProduto As Double
Dim bExisteOutro As Boolean

On Error GoTo Erro_GridProdutos_Atualiza

    'Converte a UM de GridItens para a UM de Compra
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iItem, iGrid_UM_Col), objProduto.sSiglaUMCompra, dFator)
    If lErro <> SUCESSO Then gError 68287

    dQuantidade = dQuantidade * dFator

    For iIndice = 1 To objGridProdutos.iLinhasExistentes

       'Verifica se existe algum registro no GridProdutos com o mesmo produto do item
        If (GridItens.TextMatrix(iItem, iGrid_ProdutoItem_Col) = GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col)) Then
            bExisteOutro = False
            For iIndice2 = 1 To objGridItens.iLinhasExistentes
                'Vericifa se exite outro item de pv para o mesmo produto
                If GridItens.TextMatrix(iItem, iGrid_ProdutoItem_Col) = GridItens.TextMatrix(iIndice2, iGrid_ProdutoItem_Col) And iIndice2 <> iItem Then
                    bExisteOutro = True
                    dQuantItem = StrParaDbl(GridItens.TextMatrix(iIndice2, iGrid_QuantComprarItem_Col))
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice2, iGrid_UM_Col), objProduto.sSiglaUMCompra, dFator)
                    If lErro <> SUCESSO Then gError 68288

                    dQuantItem = dQuantItem * dFator
                    dQuantProduto = dQuantItem + dQuantProduto
                End If
                
            Next
            
            If Not bExisteOutro And StrParaInt(GridItens.TextMatrix(iItem, iGrid_SelItem_Col)) = DESMARCADO Then
                Call Grid_Exclui_Linha(objGridProdutos, iIndice)
            Else
                GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col) = Formata_Estoque(dQuantidade + dQuantProduto)
            End If
        
        End If
        
    Next
    
    GridProdutos_Atualiza = SUCESSO

    Exit Function

Erro_GridProdutos_Atualiza:

    GridProdutos_Atualiza = gErr

    Select Case gErr

        Case 68287, 68288
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161443)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim sMascaraCcl As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    'objEventos
    Set objEventoFornecedor = New AdmEvento
    Set objEventoBotaoFiliaisFornProduto = New AdmEvento
    Set objEventoPedidoDe = New AdmEvento
    Set objEventoPedidoAte = New AdmEvento
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    Set objEventoBotaoEstoqueProduto = New AdmEvento
    Set objEventoRequisitante = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoBotaoCcl = New AdmEvento
    
    'Variavel global da tela
    Set gobjGeracaoReqPedVenda = New ClassGeracaoReqPedVenda
    
    'Carrega a Combo de Ordenacao
    Call Ordenacao_Carrega

    'Inicializa Mascara do Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 68103

    'Inicializa Mascara do ProdutoItem
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoItem)
    If lErro <> SUCESSO Then gError 68104

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 68105

    CentroCusto.Mask = sMascaraCcl

    'Inicializa Mascara da Conta Contábil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 68106

    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 75421

    'Preenche a Combo FilialCompra com as filiais Empresas
    lErro = Carrega_ComboFiliais(colCodigoDescricao)
    If lErro <> SUCESSO Then gError 75422
    
    'Coloca FilialCompra Default na tela
    If gobjCOM.iFilialCompra > 0 Then
        FilialCompra.Text = gobjCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If
    Call FilialCompra_Validate(bSGECancelDummy)

    'Seleciona em Filial Empresa, a Matriz como Default
    iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126942
    
    FilialEmpresa.Text = iFilialEmpresa
    Call FilialEmpresa_Validate(bSGECancelDummy)
   
    QuantidadeItem.Format = FORMATO_ESTOQUE
    PontoPedido.Format = FORMATO_ESTOQUE
    QuantEstoque.Format = FORMATO_ESTOQUE
    QuantReservada.Format = FORMATO_ESTOQUE
    QuantEmpenhada.Format = FORMATO_ESTOQUE
    QuantEmPedido.Format = FORMATO_ESTOQUE
    QuantEmReq.Format = FORMATO_ESTOQUE
    QuantCancelada.Format = FORMATO_ESTOQUE
    QuantComprar.Format = FORMATO_ESTOQUE
    QuantComprarItem.Format = FORMATO_ESTOQUE
    QuantFaturada.Format = FORMATO_ESTOQUE
    QuantidadeItem.Format = FORMATO_ESTOQUE
    QuantReservadaPV.Format = FORMATO_ESTOQUE
    SaldoAtual.Format = FORMATO_ESTOQUE
    QuantSC.Format = FORMATO_ESTOQUE
    
    'Carrega Tipos de Tributação
    lErro = Carrega_TipoTributacao()
    If lErro <> SUCESSO Then gError 68107

    'Variaveis dos Grids da Tela
    Set objGridProdutos = New AdmGrid
    Set objGridItens = New AdmGrid
    Set objGridPedidos = New AdmGrid

    'Inicializa GridProdutos
    lErro = Inicializa_GridProdutos(objGridProdutos)
    If lErro <> SUCESSO Then gError 68108

    'Inicializa GridPedidos
    lErro = Inicializa_Grid_Pedidos(objGridPedidos)
    If lErro <> SUCESSO Then gError 68109

    'Inicializa GridItens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 68110

    'Atualiza a variável global para controle de frames e seta um tipo Padrao
    iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 68103 To 68110, 68195, 75421, 75422, 126942
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161444)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'objEventos
    Set objEventoBotaoEstoqueProduto = Nothing
    Set objEventoBotaoFiliaisFornProduto = Nothing
    Set objEventoClienteAte = Nothing
    Set objEventoClienteDe = Nothing
    Set objEventoPedidoAte = Nothing
    Set objEventoPedidoDe = Nothing
    Set objEventoRequisitante = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoBotaoCcl = Nothing
    Set objEventoFornecedor = Nothing
    
    'Grids
    Set objGridItens = Nothing
    Set objGridPedidos = Nothing
    Set objGridProdutos = Nothing

    'Variavel global da tela
    Set gobjGeracaoReqPedVenda = Nothing

    Exit Sub

End Sub

Function Carrega_ComboFiliais(colCodigoDescricao As AdmColCodigoNome) As Long
'Carrega as Combos (FilialEmpresa e FilialCompra com as Filiais Empresa passada na colecao

Dim lErro As Long
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161445)

    End Select

    Exit Function

End Function

Private Sub Ordenacao_Carrega()
'preenche a combo de ordenacao e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordenação dos Pedidos
    asOrdenacao(0) = "PedidosDeVenda.Codigo,PedidosDeVenda.Cliente,PedidosDeVenda.Filial"
    asOrdenacao(1) = "PedidosDeVenda.DataEmissao,PedidosDeVenda.Codigo"

    asOrdenacaoString(0) = "Código"
    asOrdenacaoString(1) = "Data de Emissão"

    'Carrega a Combobox Ordenacao
    For iIndice = 0 To 1

        Ordenados.AddItem asOrdenacaoString(iIndice)
        Ordenados.ItemData(Ordenados.NewIndex) = iIndice

    Next

    'Seleciona a opção Codigo + Cliente + Filial de seleção
    Ordenados.ListIndex = 0

    gobjGeracaoReqPedVenda.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
    
    Exit Sub

End Sub

Function Traz_Pedidos_Tela(gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda) As Long
'Traz os Pedidos de Venda lidos do BD para a tela

Dim lErro As Long
Dim iLinha As Integer
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objItemPV As New ClassItemPedido
Dim dtDataMenor As Date

On Error GoTo Erro_Traz_Pedidos_Tela

    gobjGeracaoReqPedVenda.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
    'Preenche o grid de pedidos com os pedidos lidos
    For Each objPedidoVenda In gobjGeracaoReqPedVenda.colPedVenda
        dtDataMenor = DATA_NULA
        
        iLinha = objGridPedidos.iLinhasExistentes + 1

        GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) = objPedidoVenda.lCodigo

        'Verifica se o Cliente está preenchido
        If objPedidoVenda.lCliente <> 0 Then

            objCliente.lCodigo = objPedidoVenda.lCliente

            'Lê o Cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 68220
            'Se não encontrou o Cliente ==> erro
            If lErro = 12293 Then gError 68221

            'Informa o NomeReduzido do Cliente no Grid
            GridPedido.TextMatrix(iLinha, iGrid_NomeRed_Col) = objCliente.sNomeReduzido

        End If

        'Verifica se a FilialCliente está preenchida
        If objPedidoVenda.iFilial <> 0 Then

            objFilialCliente.lCodCliente = objPedidoVenda.lCliente
            objFilialCliente.iCodFilial = objPedidoVenda.iFilial

            'Lê a Filial do Cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 68222
            'Se não encontrou ==>Erro
            If lErro = 12567 Then gError 68223

            GridPedido.TextMatrix(iLinha, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome

        End If

        If objPedidoVenda.dtDataEmissao <> DATA_NULA Then
            GridPedido.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objPedidoVenda.dtDataEmissao, "dd/mm/yyyy")
        End If

        For Each objItemPV In objPedidoVenda.colItensPedido
        
            If (dtDataMenor = DATA_NULA And objItemPV.dtDataEntrega <> DATA_NULA) Or (objItemPV.dtDataEntrega <> DATA_NULA And objItemPV.dtDataEntrega < dtDataMenor) Then
                dtDataMenor = objItemPV.dtDataEntrega
            End If
        Next
        
        If dtDataMenor <> DATA_NULA Then GridPedido.TextMatrix(iLinha, iGrid_DataEntrega_Col) = Format(dtDataMenor, "dd/mm/yyyy")

        'Atualiza o numero de linhas existentes do GridPedidos
        objGridPedidos.iLinhasExistentes = iLinha
    Next

    Traz_Pedidos_Tela = SUCESSO

    Exit Function

Erro_Traz_Pedidos_Tela:

    Traz_Pedidos_Tela = gErr

    Select Case gErr

        Case 68220, 68222
            'Erros tratados nas rotinas chamadas

        Case 68221
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case 68223
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.iCodFilial, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161446)

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
    If lErro <> SUCESSO Then gError 68112

    'Lê o Tipo de Tributação Padrão
    lErro = CF("TipoTributacaoPadrao_Le", iTipoTrib)
    If lErro <> SUCESSO And lErro <> 66597 Then gError 68113

    'Carrega Tipos de Tributação
    For iIndice = 1 To colTributacao.Count
        TipoTribItem.AddItem colTributacao(iIndice).iCodigo & SEPARADOR & colTributacao(iIndice).sNome
    Next

    'Seleciona Tipo de Tributação default
    For iIndice = 0 To TipoTribItem.ListCount - 1
        If Codigo_Extrai(TipoTribItem.List(iIndice)) = iTipoTrib Then
            TipoTribItem.ListIndex = iIndice
            Exit For
        End If
    Next

    giTipoTributacao = iTipoTrib

    Carrega_TipoTributacao = SUCESSO

    Exit Function

Erro_Carrega_TipoTributacao:

    Carrega_TipoTributacao = gErr

    Select Case gErr

        Case 68112, 68113
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161447)

    End Select

    Exit Function

End Function

Function Move_Requisicao_Memoria(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Move os dados do GridProdutos para o objRequisicaoCompra

Dim lErro As Long
Dim iIndice As Integer
Dim objItemRC As New ClassItemReqCompras
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim objRequisitante As New ClassRequisitante
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim sContaFormata As String
Dim iContaPreenchida As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_Move_Requisicao_Memoria

    For iIndice = 1 To objGridProdutos.iLinhasExistentes

        Set objItemRC = New ClassItemReqCompras
        
        If StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col)) <> 0 Then
        
''''            Call Recolhe_ItensPV(GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), objItemRC)
        
            'Coloca o Produto no formato do BD
            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 68224
    
            objItemRC.sProduto = sProdutoFormatado
            objItemRC.sDescProduto = GridProdutos.TextMatrix(iIndice, iGrid_DescProduto_Col)
            objItemRC.sUM = GridProdutos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
            objItemRC.dQuantComprar = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col))
            objItemRC.dQuantidade = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col))
            objItemRC.dQuantCancelada = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantCancelada_Col))
            
            'Verifica se o Almoxarifado foi preenchido
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) > 0 Then
    
                objAlmoxarifado.sNomeReduzido = GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
    
                'Lê o Almoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 68149
                'Se nao encontrou ==> erro
                If lErro = 25060 Then gError 68225
    
                If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 86170
         
                objItemRC.iAlmoxarifado = objAlmoxarifado.iCodigo
    
            End If
    
            'Verifica se o CentroCusto foi preenchido
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_CentroCusto_Col))) > 0 Then
    
                'Coloca Ccl no formato do BD
                lErro = CF("Ccl_Formata", GridProdutos.TextMatrix(iIndice, iGrid_CentroCusto_Col), sCclFormata, iCclPreenchida)
                If lErro <> SUCESSO Then gError 68150
    
                objItemRC.sCcl = sCclFormata
    
            End If
    
            'Verifica se a ContaContabil foi preenchida
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_ContaContabil_Col))) > 0 Then
    
                'Coloca a Conta Contabil no formato do BD
                lErro = CF("Conta_Formata", GridProdutos.TextMatrix(iIndice, iGrid_ContaContabil_Col), sContaFormata, iContaPreenchida)
                If lErro <> SUCESSO Then gError 68151
    
                objItemRC.sContaContabil = sContaFormata
    
            End If
            
            'Verifica se o Produto é exclusivo ou preferencial
            If GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col) = "Exclusivo" Then
                objItemRC.iExclusivo = FORNECEDOR_EXCLUSIVO
            Else
                objItemRC.iExclusivo = FORNECEDOR_PREFERENCIAL
            End If
    
            objItemRC.iTipoTributacao = Codigo_Extrai(GridProdutos.TextMatrix(iIndice, iGrid_TipoTribItem_Col))
            objRequisicaoCompra.iTipoTributacao = Codigo_Extrai(GridProdutos.TextMatrix(iIndice, iGrid_TipoTribItem_Col))
    
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Fornecedor_Col))) > 0 Then
                objFornecedor.sNomeReduzido = GridProdutos.TextMatrix(iIndice, iGrid_Fornecedor_Col)
    
                'Lê o fornecedor
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 68152
    
                'Se o Fornecedor não estiver cadastrado, Erro
                If lErro = 6681 Then gError 68153
    
                objItemRC.lFornecedor = objFornecedor.lCodigo
                
                If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_FilialForn_Col))) > 0 Then objItemRC.iFilial = Codigo_Extrai(GridProdutos.TextMatrix(iIndice, iGrid_FilialForn_Col))
                
                objFornecedorProdutoFF.lFornecedor = objItemRC.lFornecedor
                objFornecedorProdutoFF.iFilialForn = objItemRC.iFilial
                objFornecedorProdutoFF.sProduto = objItemRC.sProduto
                objFornecedorProdutoFF.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
                
                lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
                If lErro <> SUCESSO And lErro <> 54217 Then gError 86097
                If lErro <> SUCESSO Then gError 86098
                
            End If
            
            
            objItemRC.sObservacao = GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col)
            'Adiciona os Itens de Requisicao em colItens
            objRequisicaoCompra.colItens.Add objItemRC
            
        End If
        
    Next

    objRequisicaoCompra.dtData = gdtDataAtual
    If EnviaCompras.Value = vbChecked Then
        objRequisicaoCompra.dtDataEnvio = gdtDataHoje
    Else
        objRequisicaoCompra.dtDataEnvio = DATA_NULA
    End If
    
    objRequisicaoCompra.dtDataBaixa = DATA_NULA
    objRequisicaoCompra.dtDataLimite = DATA_NULA
    objRequisicaoCompra.iTipoDestino = TIPO_DESTINO_EMPRESA
    objRequisicaoCompra.iFilialDestino = giFilialEmpresa
    objRequisicaoCompra.iFilialEmpresa = giFilialEmpresa
    objRequisicaoCompra.iFilialCompra = Codigo_Extrai(FilialCompra.Text)
    objRequisicaoCompra.sDigitador = gsUsuario
    
    If Len(Trim(Requisitante.Text)) = 0 Then gError 68297
    
    objRequisitante.sNomeReduzido = Requisitante.Text
    
    'Lê o Requisitante
    lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
    If lErro <> SUCESSO And lErro <> 51152 Then gError 68314
    'Se não encontrou o Requisitante ==> erro
    If lErro = 51152 Then gError 68315
    
    objRequisicaoCompra.lRequisitante = objRequisitante.lCodigo
        
    Move_Requisicao_Memoria = SUCESSO

    Exit Function

Erro_Move_Requisicao_Memoria:

    Move_Requisicao_Memoria = gErr

    Select Case gErr

        Case 68149, 68150, 68151, 68152, 68224, 68314, 86097
            'Erro tratado na rotina chamada

        Case 68153
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case 68225
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
    
        Case 68297
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_PREENCHIDO", gErr)
            
        Case 68315
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO1", gErr, objRequisitante.sNomeReduzido)
        
        Case 86098
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
                    
        Case 86170
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161448)

    End Select

    Exit Function

End Function

Function Move_TabSelecao_Memoria(gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda) As Long
'Recolhe dados do TAB de Seleção

On Error GoTo Erro_Move_TabSelecao_Memoria

    Set gobjGeracaoReqPedVenda = New ClassGeracaoReqPedVenda

    'Verifica se Exibe Todos os Pedidos
    If ExibeTodos.Value = vbChecked Then
        gobjGeracaoReqPedVenda.iExibeTodos = vbChecked
    ElseIf ExibeTodos.Value = vbUnchecked Then
        gobjGeracaoReqPedVenda.iExibeTodos = vbUnchecked
    End If

    'PedidoInicial  e PedidoFinal
    If Len(Trim(PedidoInicial.Text)) > 0 And Len(Trim(PedidoFinal.Text)) > 0 Then
        If StrParaLong(PedidoInicial.Text) > StrParaLong(PedidoFinal.Text) Then gError 68190
    End If
    
    gobjGeracaoReqPedVenda.PedidoDe = StrParaLong(PedidoInicial.Text)
    gobjGeracaoReqPedVenda.PedidoAte = StrParaLong(PedidoFinal.Text)

    'ClienteDe e ClienteAte
    If Len(Trim(ClienteDe.Text)) > 0 And Len(Trim(ClienteAte.Text)) > 0 Then
        If StrParaLong(ClienteDe.Text) > StrParaLong(ClienteAte.Text) Then gError 68191
    End If
    
    gobjGeracaoReqPedVenda.ClienteDe = StrParaLong(ClienteDe.Text)
    gobjGeracaoReqPedVenda.ClienteAte = StrParaLong(ClienteAte.Text)

    'DataEmissaoDe e DataEmissaoAte
    If Len(Trim(DataEmissaoDe.ClipText)) > 0 And Len(Trim(DataEmissaoAte.ClipText)) > 0 Then
        If StrParaDate(DataEmissaoDe.Text) > StrParaDate(DataEmissaoAte.Text) Then gError 68192
    End If
    
    gobjGeracaoReqPedVenda.DataEmissaoDe = StrParaDate(DataEmissaoDe.Text)
    gobjGeracaoReqPedVenda.DataEmissaoAte = StrParaDate(DataEmissaoAte.Text)

    'DataEntregaDe e DataEntregaAte
    If Len(Trim(DataEntregaDe.ClipText)) > 0 And Len(Trim(DataEntregaAte.ClipText)) > 0 Then
        If StrParaDate(DataEntregaDe.Text) > StrParaDate(DataEntregaAte.Text) Then gError 68193
    End If
    
    gobjGeracaoReqPedVenda.DataEntregaDe = StrParaDate(DataEntregaDe.Text)
    gobjGeracaoReqPedVenda.DataEntregaAte = StrParaDate(DataEntregaAte.Text)

    If gobjGeracaoReqPedVenda.sOrdenacao = "" Then
        gobjGeracaoReqPedVenda.sOrdenacao = "PedidosDeVenda.Codigo"
    End If
    
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case 68190
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)

        Case 68191
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", gErr)

        Case 68192, 68193
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161449)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Requisição a partir de Pedidos de Venda"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoReqPedVenda"

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

Private Sub BotaoCcl_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_BotaoCcls_Click

    'Se nenhuma linha foi selecionada do Grid, Erro
    If GridProdutos.Row = 0 Then gError 75412

    'Verifica se o Produto foi preenchido
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) = 0 Then gError 75413

    'Verifica se o Ccl Foi preenchido
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_CentroCusto_Col))) > 0 Then

        sCclFormatada = String(STRING_CCL, 0)

        lErro = CF("Ccl_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_CentroCusto_Col), sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 75414

        objCcl.sCcl = sCclFormatada

    End If

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoBotaoCcl)

    Exit Sub

Erro_BotaoCcls_Click:

    Select Case gErr

        Case 75412
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 75413
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 75414

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161450)

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
    If lErro <> SUCESSO Then gError 84543 '(65562)
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub
    
Erro_FilialEmpresa_Click:
    
    Select Case gErr
        
        Case 84543 'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161451)
        
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84539 '61922

    'Se não encontrou o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 84540 '61923

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 84541 '61924

        'coloca na tela
        FilialEmpresa.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome
        
        lErro = AlmoxarifadoPadrao_Preenche(iCodigo)
        If lErro <> SUCESSO Then gError 86162

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 84542 '61925
        
    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True
    
    Select Case gErr

        Case 84539, 84540, 86162

        Case 84541
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialEmpresa.Text)
            
        Case 84542
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialEmpresa.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161452)

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
            If lErro <> SUCESSO Then gError 84544 '61915

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 84545 '61916

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

        Case 84544, 84545

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161453)

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

Private Sub objEventoBotaoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se nenhuma foi selecionada, erro
    If GridProdutos.Row = 0 Then gError 75415

    sContaEnxuta = String(STRING_CCL, 0)

    'Coloca a conta no formato conta enxuta
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 75416

    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_CentroCusto_Col) = sCclMascarado
    CentroCusto.PromptInclude = False
    CentroCusto.Text = sCclMascarado
    CentroCusto.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 75415
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 75416

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161454)

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

    If GridProdutos.Row = 0 Then gError 75408

    If GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col) = "" Then gError 75409

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 75410

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 75408
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 75409
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 75410

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161455)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta <> "" Then
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 75411
            
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
        
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 75411
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161456)

    End Select

    Exit Sub

End Sub

Private Sub ClienteAte_GotFocus()
    
Dim iTabSelecao As Integer

    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(ClienteAte, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub ClienteDe_GotFocus()

Dim iTabSelecao As Integer
    
    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(ClienteDe, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteDe_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(ClienteDe.Text)) > 0 Then
    
        objCliente.lCodigo = StrParaLong(ClienteDe.Text)
        
        'Lê o Cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 72381
        
        'Se não encontrou o Cliente ==> erro
        If lErro = 12293 Then gError 72382
        
    End If
    
    Exit Sub
    
Erro_ClienteDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 72381
            'Erro tratado na rotina chamada
            
        Case 72382
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161457)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteAte_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(ClienteAte.Text)) > 0 Then
    
        objCliente.lCodigo = StrParaLong(ClienteAte.Text)
        
        'Lê o Cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 72383
        
        'Se não encontrou o Cliente ==> erro
        If lErro = 12293 Then gError 72384
        
    End If
    
    Exit Sub
    
Erro_ClienteAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 72383
            'Erro tratado na rotina chamada
            
        Case 72384
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161458)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub DataEmissaoAte_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub DataEmissaoDe_GotFocus()
    
Dim iTabSelecao As Integer
    
    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub DataEntregaAte_GotFocus()
    
Dim iTabSelecao As Integer

    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaAte, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub DataEntregaDe_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaDe, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub objEventoBotaoEstoqueProduto_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto

    Set objEstoqueProduto = obj1

    'Preenche campo Almoxarifado
    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    Me.Show

    Exit Sub

End Sub

Private Sub PedidoFinal_GotFocus()
    
Dim iTabSelecao As Integer

    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoFinal, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub PedidoInicial_GotFocus()
    
Dim iTabSelecao As Integer
    
    iTabSelecao = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoInicial, iAlterado)
    iFrameSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub QuantComprarItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantComprarItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'Private Sub QuantSC_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub QuantSC_GotFocus()
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Private Sub QuantSC_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Private Sub QuantSC_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = QuantidadeItem
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub

Private Sub SelItem_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim lCodPedido As Long
Dim objItemPV As ClassItemPedido
Dim iItem As Integer
Dim iIndice As Integer, iIndice2 As Integer

On Error GoTo Erro_SelItem_Click

    If StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_SelItem_Col)) = DESMARCADO Then
        
''''        GridItens.TextMatrix(GridItens.Row, iGrid_QuantComprarItem_Col) = ""
        
        'Formata o produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItem_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 68267

        'Guarda o Codigo do Produto em objProduto
        objProduto.sCodigo = sProduto

        'Le o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 68268
        'Se nao encontrou => erro
        If lErro = 28030 Then gError 68269

        'Reflete no GridProdutos
        lErro = GridProdutos_Atualiza(GridItens.Row, objProduto, 0)
        If lErro <> SUCESSO Then gError 68270
    
    Else
    
        lCodPedido = StrParaLong(GridItens.TextMatrix(GridItens.Row, iGrid_PedidoItem_Col))
        iItem = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_ItemPed_Col))
        
        For iIndice = 1 To gobjGeracaoReqPedVenda.colPedVenda.Count
            If gobjGeracaoReqPedVenda.colPedVenda(iIndice).lCodigo = lCodPedido Then
                For iIndice2 = 1 To gobjGeracaoReqPedVenda.colPedVenda(iIndice).colItensPedido.Count
                    If gobjGeracaoReqPedVenda.colPedVenda(iIndice).colItensPedido(iIndice2).iItem = iItem Then
                        Set objItemPV = gobjGeracaoReqPedVenda.colPedVenda(iIndice).colItensPedido(iIndice2)
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
        
        lErro = GridProdutos_Preenche(objItemPV, iItem)
        
    End If
        
    Exit Sub

Erro_SelItem_Click:

    Select Case gErr

        Case 68267, 68268, 68270

        Case 68269
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161459)

    End Select

    Exit Sub

End Sub

Private Sub Text1_Change()

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

        'Limpa os almoxarifados e os Ccls do GridProduto
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161460)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Requisitante Then
            Call RequisitanteLabel_Click
        ElseIf Me.ActiveControl Is PedidoInicial Then
            Call LabelPedidoDe_Click
        ElseIf Me.ActiveControl Is PedidoFinal Then
            Call LabelPedidoAte_Click
        ElseIf Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoqueProd_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is CentroCusto Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call BotaoProdutoFiliaisForn_Click
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


Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os Pedidos do GridPedidos

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPedidos.iLinhasExistentes
        'Desmarca na tela a linha em questão
        GridPedido.TextMatrix(iLinha, iGrid_SelecionadoPed_Col) = GRID_CHECKBOX_INATIVO
    Next

    'Limpa o GridProdutos
    Call Grid_Limpa(objGridProdutos)

    'Limpa o GridItens
    Call Grid_Limpa(objGridItens)

    iAlterado = REGISTRO_ALTERADO

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridPedidos)

    Exit Sub

End Sub

Private Sub BotaoEstoqueProd_Click()
'Chama a tela EstoqueProdutoFilialLista

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iFilialEmpresa As Integer

On Error GoTo Erro_BotaoEstoqueProd_Click

    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = False Then Exit Sub

    If GridProdutos.Row = 0 Then Error 61884

    'Verifica se o Produto da linha atual do GridItens está preenchido
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) > 0 Then

        'Formata o Produto
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 68115

        iFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text)
        If Len(Trim(iFilialEmpresa)) = 0 Then gError 84612
    
        'adiciona produto e FiliaEmpresa de Destino em colSelecao
        colSelecao.Add sProdutoFormatado
        colSelecao.Add iFilialEmpresa
        
        Call Chama_Tela("AlmoxarifadoFilialLista", colSelecao, objEstoqueProduto, objEventoBotaoEstoqueProduto)

    End If

    Exit Sub

Erro_BotaoEstoqueProd_Click:

    Select Case gErr

        Case 84612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_ENTREGA_NAO_PREENCHIDA", Err)
        
        Case 61884
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 68115
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161461)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGeraReq_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGeraReq_Click

    'Grava a Requisicao de Compra
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 68127

    'Limpa a tela
    Call Limpa_Tela_GeracaoReqPedVenda

    Exit Sub

Erro_BotaoGeraReq_Click:

    Select Case gErr

        Case 68127
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161462)

    End Select

    Exit Sub

End Sub
Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iMarcado As Integer
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim objProduto As New ClassProduto
Dim colPedidoVenda As New Collection
Dim iQuantComprar As Integer
Dim dQuantComprarProduto As Double
Dim dQuantComprarItem As Double
Dim dQuant As Double
Dim dFator As Double
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Gravar_Registro

    'Se não houver nenhum Pedido de Venda no GridPedidos ==> erro
    If objGridPedidos.iLinhasExistentes = 0 Then gError 68128

    For iIndice = 1 To objGridPedidos.iLinhasExistentes

        'Verifica se o Pedido de Venda está selecionado
        If StrParaInt(GridPedido.TextMatrix(iIndice, iGrid_SelecionadoPed_Col)) = MARCADO Then

            iMarcado = iMarcado + 1
            For iIndice2 = 1 To objGridItens.iLinhasExistentes
                If StrParaLong(GridPedido.TextMatrix(iIndice, iGrid_Pedido_Col)) = StrParaLong(GridItens.TextMatrix(iIndice2, iGrid_PedidoItem_Col)) Then
                    'Verifica se a Quantidade a Comprar do Item do PedVenda é maior que zero
                    If StrParaDbl(GridItens.TextMatrix(iIndice2, iGrid_QuantComprarItem_Col)) > 0 Then
                        iQuantComprar = iQuantComprar + 1
                    End If
                End If
            Next
            'Se nenhum item de Pedido de Venda tiver quantidade a comprar maior que zero ==> erro
            If iQuantComprar = 0 Then gError 68255

        End If
    Next
'************************* Inicio do Trecho Copiado de ReqCompras ****************
    
    'Se não existir Pedido de Venda selecionado no GridPedidos ==> erro
    If iMarcado = 0 Then gError 84587 '68129

    For iIndice = 1 To objGridProdutos.iLinhasExistentes

        'Define a QuantComprar para o Produto no GridProdutos
        dQuantComprarProduto = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col))

        For iIndice2 = 1 To objGridItens.iLinhasExistentes

            If GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = GridItens.TextMatrix(iIndice2, iGrid_ProdutoItem_Col) Then
                'Calcula a QuantComprar para os Itens com o mesmo produto no GridItens
                dQuant = StrParaDbl(GridItens.TextMatrix(iIndice2, iGrid_QuantComprarItem_Col))
                
                lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 84588 '72511

                objProduto.sCodigo = sProdutoFormatado
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO Then gError 84589 '72509
                'Converte dQuant para UM de compra
                lErro = CF("UM_Conversao", objProduto.iClasseUM, GridItens.TextMatrix(iIndice2, iGrid_UM_Col), objProduto.sSiglaUMCompra, dFator)
                If lErro <> SUCESSO Then gError 84590 '72510

                dQuant = dQuant * dFator
                dQuantComprarItem = dQuant + dQuantComprarItem
            End If
        Next
        'Verifica se QuantComprar do GridProdutos é igual a QuantComprar do GridItens para o Produto em questao
        If Abs(dQuantComprarProduto - dQuantComprarItem) >= 0.0001 Then gError 84591 '68256
        dQuantComprarItem = 0
        
    Next
    
    'Verifica se o Grid foi preenchido
    If objGridItens.iLinhasExistentes = 0 Then gError 84592 '61971

    'Se o tipo destino for empresa
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then
    
        'Se a FilialEmpresa não estiver preenchida, erro
        If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 84593 '61972
    
        'Critica o preenchimento do Almoxarifado
        For iIndice = 1 To objGridProdutos.iLinhasExistentes
                
            'Formata o Produto
            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 84594 '67880
            
            objProduto.sCodigo = sProdutoFormatado
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 23080 Then gError 84595 '67881
            If lErro <> SUCESSO Then gError 84596 '67882
            
            'Se o produto possui Estoque e o almoxarifado não foi preenchido, erro
            If Not (objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) And Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 84597 '66587
                    
        Next
    
    'Se o tipo destino for Fornecedor
    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then
    
        'Se o Fornecedor não estiver preenchido, erro
        If Len(Trim(Fornec.Text)) = 0 Then gError 84598  '61973
        
        'Se a Filial do Fornecedor não estiver preenchida, erro
        If Len(Trim(FilialFornec.Text)) = 0 Then gError 84599 '61974
            
    End If

'**************************** Fim do Trecho ***********************
    
    'Recolhe os dados do TabProdutos e armazena em objRequisicaoCompras
    lErro = Move_Requisicao_Memoria(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 68146

    'Preenche colPedidoVenda com os Pedidos de Venda selecionados
    lErro = Carrega_Dados_Pedido(colPedidoVenda)
    If lErro <> SUCESSO Then gError 68147
    
    Call Recolhe_ItensPV(objRequisicaoCompras.colItens, colPedidoVenda)

    'Gera a Requisicao de Compra
    lErro = CF("GeracaoReqPedVenda_Grava", objRequisicaoCompras, colPedidoVenda)
    If lErro <> SUCESSO Then gError 68148

    'Avisa com que codigo a Requisicao de Compras foi gerada
    vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_REQUISICAOCOMPRA_GERADA", objRequisicaoCompras.lCodigo)
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 68128
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDVENDAS_GRID", gErr)

        Case 84587
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDVENDA_SELECIONADO", gErr)

        Case 68146, 68147, 68148, 84589, 84590, 84588
            'Erros tratados nas rotinas chamadas

        Case 68255
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_ITENSPED", gErr, GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col))

        Case 84591
            
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_PRODUTOS_DIFERENTE", gErr)

'**********************************************
        
        Case 84592
            
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDITENS_VAZIO", gErr)
         
        Case 84593
            
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)
        
        Case 84598
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO", gErr)
        
        Case 84599
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA", gErr)
                        
        Case 84594, 84595, 89156
        
        Case 84597
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
            
        Case 84596
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

'*************************************
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161463)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Function Carrega_Dados_Pedido(colPedidoVenda As Collection) As Long
'Carrega os dados dos Pedidos de Venda que estão selecionados no GridPedido em colPedidoVenda

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objItemPV As New ClassItemPedido
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim colReserva As New colReserva

On Error GoTo Erro_Carrega_Dados_Pedido

    Set objPedidoVenda = New ClassPedidoDeVenda
    
    For iIndice = 1 To objGridPedidos.iLinhasExistentes
    
        Set objPedidoVenda = New ClassPedidoDeVenda
        
        'Verifica se o Pedido está selecionado
        If StrParaInt(GridPedido.TextMatrix(iIndice, iGrid_SelecionadoPed_Col)) = MARCADO Then

            
            For iIndice2 = 1 To objGridItens.iLinhasExistentes

                'Procura pelos itens do Pedido de Venda que foi selecionado
                If StrParaLong(GridItens.TextMatrix(iIndice2, iGrid_PedidoItem_Col)) = StrParaLong(GridPedido.TextMatrix(iIndice, iGrid_Pedido_Col)) Then

                    'Verifica se o item do Pedido de Venda selecionado tem QuantComprar>0
                    If StrParaDbl(GridItens.TextMatrix(iIndice2, iGrid_QuantComprarItem_Col)) > 0 Then
                        
                        Set objItemPV = New ClassItemPedido
                        objItemPV.dQuantSC = StrParaDbl(GridItens.TextMatrix(iIndice2, iGrid_QuantComprarItem_Col))
                        objItemPV.iFilialEmpresa = giFilialEmpresa
                        objItemPV.lCodPedido = StrParaLong(GridItens.TextMatrix(iIndice2, iGrid_PedidoItem_Col))
                        objItemPV.iItem = StrParaInt(GridItens.TextMatrix(iIndice2, iGrid_ItemPed_Col))
                        objItemPV.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UM_Col)
                        objItemPV.dtDataEntrega = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataEntregaItem_Col))
                        objItemPV.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantidadeItem_Col))
                        objItemPV.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancelada_Col))
                        objItemPV.dQuantReservada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantReservadaPV_Col))
                        objItemPV.dQuantFaturada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col))
                        
                        
                        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
                        
                        objPedidoVenda.lCodigo = objItemPV.lCodPedido
                        
                        'Coloca o Produto no formato do BD
                        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice2, iGrid_ProdutoItem_Col), sProdutoFormatado, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 68303
                    
                        objItemPV.sProduto = sProdutoFormatado
                        objItemPV.sDescricao = GridItens.TextMatrix(iIndice2, iGrid_DescricaoProduto_Col)
                        Call Busca_ItemPV(objItemPV)
                        
                        'Adiciona os Itens do Pedido de Venda em objPedidoVenda.colItensPedido
                        objPedidoVenda.colItensPedido.Add objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.iItem, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sProdutoDescricao, objItemPV.dValorAbatComissao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, colReserva, objItemPV.sProdutoNomeReduzido, objItemPV.sUnidadeMed, objItemPV.sLote, objItemPV.sUMEstoque, objItemPV.iClasseUM, objItemPV.dQuantFaturada, objItemPV.dQuantAFaturar, objItemPV.sDescricao, objItemPV.iStatus, objItemPV.iControleEstoque, objItemPV.dQuantOP, objItemPV.dQuantSC, 0, 0, 0, 0, 0, 0, objItemPV.lNumIntDoc
                    End If
                End If
            Next
            'Adiciona o Pedido de Venda em colPedidoVenda
            If objPedidoVenda.colItensPedido.Count > 0 Then colPedidoVenda.Add objPedidoVenda
        End If

    Next

    Carrega_Dados_Pedido = SUCESSO

    Exit Function

Erro_Carrega_Dados_Pedido:

    Carrega_Dados_Pedido = gErr

        Select Case gErr

            Case 68278, 68296, 68303
                'Erros tratados nas rotinas chamadas
                
            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161464)

        End Select

    Exit Function

End Function
Private Sub BotaoMarcarTodos_Click()
'Marca todos os Pedidos do GridPedidos

Dim lErro As Long
Dim iIndice As Integer
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoMarcarTodos_Click

    For iIndice = 1 To objGridPedidos.iLinhasExistentes

        'Verifica se o Pedido desta linha não está selecionado
        If StrParaInt(GridPedido.TextMatrix(iIndice, iGrid_SelecionadoPed_Col)) = DESMARCADO Then

            'Seleciona o Pedido em questão
            GridPedido.TextMatrix(iIndice, iGrid_SelecionadoPed_Col) = vbChecked

            Set objPedidoVenda = gobjGeracaoReqPedVenda.colPedVenda(iIndice)

            'Traz os Itens do Pedido de Venda que foi selecionado para a tela
            lErro = Traz_ItensPed_Tela(objPedidoVenda)
            If lErro <> SUCESSO Then gError 68123

        End If

    Next

    Call Grid_Refresh_Checkbox(objGridPedidos)

    Exit Sub

Erro_BotaoMarcarTodos_Click:

    Select Case gErr

        Case 68123
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161465)

    End Select

    Exit Sub

End Sub

Private Function Traz_ItensPed_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long
'Traz para a tela os Itens do Pedido de Venda

Dim lErro As Long
Dim objItemPV As New ClassItemPedido
Dim iLinha As Integer
Dim iItem As Integer
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto
Dim colProdutos As New Collection
Dim bAchou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Traz_ItensPed_Tela


    'Preenche o Grid de Itens de Pedido de Venda
    For Each objItemPV In objPedidoVenda.colItensPedido
        iItem = iItem + 1
        objItemPV.iItem = iItem
        Call Busca_Produto(objItemPV.sProduto, colProdutos, objProduto, bAchou)
        
        If Not bAchou Then
        
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objItemPV.sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 62700
            If lErro <> SUCESSO Then gError 62701

            colProdutos.Add objProduto

        End If
                  
'        If objProduto.iCompras = PRODUTO_COMPRAVEL And (objItemPV.dQuantidade - objItemPV.dQuantSC - objItemPV.dQuantCancelada) >= 0.0001 Then
            
        If (objItemPV.dQuantidade - objItemPV.dQuantSC - objItemPV.dQuantCancelada) >= 0.0001 Then
            
            iLinha = objGridItens.iLinhasExistentes + 1
            GridItens.TextMatrix(iLinha, iGrid_SelItem_Col) = MARCADO
            GridItens.TextMatrix(iLinha, iGrid_PedidoItem_Col) = objPedidoVenda.lCodigo
            GridItens.TextMatrix(iLinha, iGrid_ItemPed_Col) = iItem
    
            'Mascara o Produto
            lErro = Mascara_MascararProduto(objItemPV.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 68213
    
            GridItens.TextMatrix(iLinha, iGrid_ProdutoItem_Col) = sProdutoMascarado
            GridItens.TextMatrix(iLinha, iGrid_DescricaoProduto_Col) = objItemPV.sProdutoDescricao
            GridItens.TextMatrix(iLinha, iGrid_UM_Col) = objItemPV.sUnidadeMed
    
            If objItemPV.dtDataEntrega <> DATA_NULA Then
                GridItens.TextMatrix(iLinha, iGrid_DataEntregaItem_Col) = Format(objItemPV.dtDataEntrega, "dd/mm/yyyy")
            End If
    
            GridItens.TextMatrix(iLinha, iGrid_QuantidadeItem_Col) = Formata_Estoque(objItemPV.dQuantidade)
            GridItens.TextMatrix(iLinha, iGrid_QuantCancelada_Col) = Formata_Estoque(objItemPV.dQuantCancelada)
            GridItens.TextMatrix(iLinha, iGrid_QuantReservadaPV_Col) = Formata_Estoque(objItemPV.dQuantReservada)
            GridItens.TextMatrix(iLinha, iGrid_QuantSC_Col) = Formata_Estoque(objItemPV.dQuantSC)
            GridItens.TextMatrix(iLinha, iGrid_QuantFaturada_Col) = Formata_Estoque(objItemPV.dQuantFaturada)
            GridItens.TextMatrix(iLinha, iGrid_QuantComprarItem_Col) = Formata_Estoque(objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantReservada)
        
            objGridItens.iLinhasExistentes = iLinha
    
            'Preenche o GridProdutos com o Produto de ItensPedidoVenda
            lErro = GridProdutos_Preenche(objItemPV, iLinha)
            If lErro <> SUCESSO Then gError 68214
        
        End If

    Next
    
    Call Grid_Refresh_Checkbox(objGridItens)
    
    Set objPedidoVenda = New ClassPedidoDeVenda

    Traz_ItensPed_Tela = SUCESSO

    Exit Function

Erro_Traz_ItensPed_Tela:

    Traz_ItensPed_Tela = gErr

    Select Case gErr

        Case 62700
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 68137, 68138, 68213, 68214, 62700
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161466)

    End Select

    Exit Function

End Function

Private Function GridProdutos_Preenche(objItemPV As ClassItemPedido, iItem As Integer) As Long
'Preenche o GridProdutos a partir do Produto do Item do Pedido de Venda

Dim lErro As Long
Dim sProdutoMascarado As String
Dim iAlmoxarifadoPadrao As Integer
Dim bAchou As Boolean, iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double, dFatorEst As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaEnxuta As String, sConta As String
Dim dSaldoAtual As Double, dQuantEmPedidoCompra As Double
Dim dQuantEmReqCompra As Double, dQuantComprarItem As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim dQuantEmEstoque As Double, dQuantEmpenhada As Double
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim iCont As Integer
Dim dQuantReservada As Double
Dim dSaldoDisponivel As Double

On Error GoTo Erro_GridProdutos_Preenche

    objProduto.sCodigo = objItemPV.sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 68216

    'Se não encontrou o Produto
    If lErro = 28030 Then gError 68217

    'Converte dQuant para UM de compra
    lErro = CF("UM_Conversao", objProduto.iClasseUM, GridItens.TextMatrix(iItem, iGrid_UM_Col), objProduto.sSiglaUMCompra, dFator)
    If lErro <> SUCESSO Then gError 68138
    
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objProduto.sSiglaUMCompra, dFatorEst)
    If lErro <> SUCESSO Then gError 86165
    
    'Mascara o Produto
    lErro = Mascara_MascararProduto(objItemPV.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 68215

    'Se já existem outros produtos  no GridProdutos
    If objGridProdutos.iLinhasExistentes <> 0 Then

        For iIndice = 1 To objGridProdutos.iLinhasExistentes

            'Verifica se o Produto já existe no gridProdutos
            If GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado Then
                bAchou = True
                
                'Guarda a quantidade do item para a linha em questão
                dQuantComprarItem = StrParaDbl(GridItens.TextMatrix(iItem, iGrid_QuantComprarItem_Col))
                dQuantComprarItem = dQuantComprarItem * dFator
                
                'Guarda a quantidade total de todas as linhas onde o produto aparece
                GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col) = Formata_Estoque(dQuantComprarItem + StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantComprar_Col)))
                                
                Exit For

            End If
        Next
    End If
    
    'Se o produto não está no GridProdutos ou o GridProdutos está vazio
    If (bAchou = False) Or objGridProdutos.iLinhasExistentes = 0 Then

        objGridProdutos.iLinhasExistentes = objGridProdutos.iLinhasExistentes + 1
        
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_Produto_Col) = sProdutoMascarado
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_DescProduto_Col) = objItemPV.sProdutoDescricao
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_QuantComprar_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iItem, iGrid_QuantComprarItem_Col)) * dFator)
        
        lErro = SaldoAtualProdComprado(objItemPV.sProduto, dSaldoAtual, dQuantReservada, dQuantEmReqCompra, dQuantEmPedidoCompra, dSaldoDisponivel, dQuantEmpenhada)
        If lErro <> SUCESSO Then gError 86164
        
        'Calcula Quantidade em Estoque do Produto
        dQuantEmEstoque = dSaldoDisponivel + dQuantReservada
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_QuantEstoque_Col) = Formata_Estoque(dQuantEmEstoque * dFatorEst)
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantReservada * dFatorEst)
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_QuantEmpenhada_Col) = Formata_Estoque(dQuantEmpenhada * dFatorEst)
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_QuantEmPedido_Col) = Formata_Estoque(dQuantEmPedidoCompra * dFatorEst)
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_QuantEmReq_Col) = Formata_Estoque(dQuantEmReqCompra * dFatorEst)
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_SaldoAtual_Col) = Formata_Estoque(dSaldoAtual * dFatorEst)
        
        objProdutoFilial.sProduto = objProduto.sCodigo
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        
        'Le o Produto
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 68319
        
        'Se não encontrou ==> Erro
        If lErro = 28261 Then gError 68320
        
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_PontoPedido_Col) = Formata_Estoque(objProdutoFilial.dPontoPedido * dFatorEst)
        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_LoteEconomico_Col) = Formata_Estoque(objProdutoFilial.dLoteEconomico * dFatorEst)
        
        If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True And Len(Trim(FilialEmpresa.Text)) > 0 Then
            
            lErro = CF("AlmoxarifadoPadrao_Le", Codigo_Extrai(FilialEmpresa.Text), objItemPV.sProduto, iAlmoxarifadoPadrao)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 86166
            If iAlmoxarifadoPadrao > 0 Then
            
                objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
    
                'Lê os dados do Almoxarifado a partir do código passado
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 68307
        
                GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            End If
        End If
        
        If Len(Trim(objProduto.sContaContabil)) > 0 Or Len(Trim(objAlmoxarifado.sContaContabil)) > 0 Then
            
            sConta = IIf(Len(Trim(objProduto.sContaContabil)) > 0, objProduto.sContaContabil, objAlmoxarifado.sContaContabil)
        
            lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
            If lErro <> SUCESSO Then gError 68308
            
            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True
            GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_ContaContabil_Col) = ContaContabil.Text
                
        End If
    
        'Seleciona Tipo de Tributação padrão
        For iCont = 0 To TipoTribItem.ListCount - 1
            If Codigo_Extrai(TipoTribItem.List(iCont)) = giTipoTributacao Then
                TipoTribItem.ListIndex = iCont
                GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_TipoTribItem_Col) = TipoTribItem.Text
                Exit For
            End If
        Next
    
        If Codigo_Extrai(FilialCompra.Text) <> giFilialEmpresa Then
            
            objProdutoFilial.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
            
            lErro = CF("ProdutoFilial_Le", objProdutoFilial)
            If lErro <> SUCESSO And lErro <> 28261 Then gError 68319
            
            'Se não encontrou ==> Erro
            If lErro = 28261 Then gError 68320
            
            'Coloca Nome Reduzido do Fornecedor no Grid
            If objProdutoFilial.lFornecedor > 0 Then
    
                objFornecedor.lCodigo = objProdutoFilial.lFornecedor
                lErro = CF("Fornecedor_Le", objFornecedor)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 75404
                If lErro = 12729 Then gError 75405
    
                GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
                           
                'Coloca Filial do Fornecedor no Grid
                If objProdutoFilial.iFilialForn > 0 Then
                    
                    objFilialFornecedor.iCodFilial = objProdutoFilial.iFilialForn
                    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
                    
                    lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                    If lErro <> SUCESSO And lErro <> 12929 Then gError 75406
                    If lErro = 12929 Then gError 75407
                    
                    GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_FilialForn_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
                End If
                                        
                'Preenche Exclusivo
                For iCont = 0 To Exclusivo.ListCount - 1
                    If Exclusivo.ItemData(iCont) = FORNECEDOR_PREFERENCIAL Then
                        GridProdutos.TextMatrix(objGridProdutos.iLinhasExistentes, iGrid_Exclusivo_Col) = Exclusivo.List(iCont)
                        Exit For
                    End If
                Next
                
            End If
        End If
                
    End If

    GridProdutos_Preenche = SUCESSO

    Exit Function

Erro_GridProdutos_Preenche:

    GridProdutos_Preenche = gErr

    Select Case gErr

        Case 68138, 68216, 86165, 68215, 86164, 68319, 86166, 68307, 68308, _
             68319, 75404, 75406
            'Erros tratados nas rotinas chamadas

        Case 68217
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 68320
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, objProduto.sCodigo, objProdutoFilial.iFilialEmpresa)
            
        Case 75405
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
        
        Case 75407
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161467)

    End Select

    Exit Function

End Function

Private Function Exclui_ItensPed_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long
'Exlui os Itens de Pedido de Venda do GridItens

Dim objItemPV As New ClassItemPedido
Dim sProd As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iIndice3 As Integer
Dim dQuantComprarItem As Double
Dim bAchou As Boolean
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim lErro As Long

On Error GoTo Erro_Exclui_ItensPed_Tela

    For iIndice = objGridItens.iLinhasExistentes To 1 Step -1

        'Procura os Itens do Pedido de Venda passado
        If objPedidoVenda.lCodigo = StrParaLong(GridItens.TextMatrix(iIndice, iGrid_PedidoItem_Col)) Then
            
            'Verifica se o Produto está presente no GridItens
            sProd = GridItens.TextMatrix(iIndice, iGrid_ProdutoItem_Col)
            
            lErro = CF("Produto_Formata", sProd, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 55555
            
            objProduto.sCodigo = sProdutoFormatado
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 23080 Then gError 11112
            If lErro <> SUCESSO Then gError 11113
            
            lErro = CF("UM_Conversao", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UM_Col), objProduto.sSiglaUMCompra, dFator)
            If lErro <> SUCESSO Then gError 66666
            
            dQuantComprarItem = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantComprarItem_Col)) * dFator
            
            'Exclui a linha do GridItens
            Call Grid_Exclui_Linha(objGridItens, iIndice)

            For iIndice3 = 1 To objGridProdutos.iLinhasExistentes

                'Verifica se o Produto do GridItens é o mesmo da linha do GridProdutos
                If sProd = GridProdutos.TextMatrix(iIndice3, iGrid_Produto_Col) Then

                


                    'Atualiza a QuantComprar do GridProdutos
                    GridProdutos.TextMatrix(iIndice3, iGrid_QuantComprar_Col) = Formata_Estoque(StrParaDbl(GridProdutos.TextMatrix(iIndice3, iGrid_QuantComprar_Col)) - dQuantComprarItem)

                End If
            Next

            For iIndice2 = 1 To objGridItens.iLinhasExistentes

                'Verifica se existe algum Item com o mesmo produto do item que foi excluido do GridItens
                If GridItens.TextMatrix(iIndice2, iGrid_ProdutoItem_Col) = sProd Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If bAchou = False Then
                For iIndice3 = 1 To objGridProdutos.iLinhasExistentes

                    If sProd = GridProdutos.TextMatrix(iIndice3, iGrid_Produto_Col) Then

                        'Exclui o produto do GridProdutos
                        Call Grid_Exclui_Linha(objGridProdutos, iIndice3)

                    End If
                Next

            End If

            bAchou = False

        End If

    Next

    'Se não existe mais nenhum item no GridItens
    If objGridItens.iLinhasExistentes = 0 Then
        Call Grid_Limpa(objGridProdutos)
    End If

    Exclui_ItensPed_Tela = SUCESSO

    Exit Function

Erro_Exclui_ItensPed_Tela:

    Exclui_ItensPed_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161468)

    End Select

    Exit Function

End Function

Private Sub BotaoPedido_Click()
'Chama a Tela de Produto

Dim objPedidoVenda As New ClassPedidoDeVenda

    'Verifica se o Pedido está selecionado
    If GridPedido.Row = 0 Then Exit Sub

    objPedidoVenda.lCodigo = StrParaLong(GridPedido.TextMatrix(GridPedido.Row, iGrid_Pedido_Col))
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    'Chama a tela PedidoVenda
    Call Chama_Tela("PedidoVenda", objPedidoVenda)

    Exit Sub

End Sub

Private Sub BotaoProdutoFiliaisForn_Click()
'Chama a tela FiliaisFornProdutoLista

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_BotaoProdutoFiliaisForn_Click

    'Verifica se existe alguma linha selecionada no GridItens
    If GridProdutos.Row = 0 Then Exit Sub

    'Verifica se o Produto da linha selecionada foi preenchido
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) = 0 Then Exit Sub
    If Len(Trim(FilialCompra.Text)) = 0 Then gError 75400

    'Formata o Produto
    lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 68114

    colSelecao.Add sProdutoFormatado
    colSelecao.Add Codigo_Extrai(FilialCompra.Text)

    'Chama a tela FiliaisFornProdutoLista
    Call Chama_Tela("FiliaisFornProdutoLista", colSelecao, objFornecedorProdutoFF, objEventoBotaoFiliaisFornProduto)

    Exit Sub

Erro_BotaoProdutoFiliaisForn_Click:

    Select Case gErr

        Case 75400
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", gErr)

        Case 68114
            'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161469)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()
'Chama a tela de Produto

Dim objProdutos As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer, lErro As Long
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    'Se existe alguma linha do GridProdutos selecionada
    If GridProdutos.Row = 0 Then Exit Sub
        
    'Formata o Produto
    lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 75399
   
    objProdutos.sCodigo = sProdutoFormatado

    'Chama a tela Produto
    Call Chama_Tela("Produto", objProdutos)

    Exit Sub

Erro_BotaoProdutos_Click:
    
    Select Case gErr
        
        Case 75399
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161470)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub QuantComprar_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantComprar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub QuantComprar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub QuantComprar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = QuantComprar
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = DescProduto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
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

Private Sub Selecionado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteEconomico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteEconomico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub LoteEconomico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub LoteEconomico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = LoteEconomico
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PontoPedido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub PontoPedido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub PontoPedido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub PontoPedido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = PontoPedido
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub SaldoAtual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SaldoAtual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub SaldoAtual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub SaldoAtual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = SaldoAtual
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantEmReq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub QuantEmReq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub QuantEmReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub QuantEmReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = QuantEmReq
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantEmPedido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantEmPedido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub QuantEmPedido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub QuantEmPedido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = QuantEmPedido
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantEmpenhada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantEmpenhada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub QuantEmpenhada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub QuantEmpenhada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = QuantEmpenhada
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantReservada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantReservada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub QuantReservada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub QuantReservada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = QuantReservada
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantEstoque_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantEstoque_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub QuantEstoque_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub QuantEstoque_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = QuantEstoque
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Selecionado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPedidos)

End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPedidos)

End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridPedidos)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub PedidoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PedidoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PedidoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PedidoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantFaturada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantFaturada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantFaturada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub QuantFaturada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantFaturada
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantCancelada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantCancelada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantCancelada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub QuantCancelada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantCancelada
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantidadeItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub QuantidadeItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantidadeItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantComprarItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub QuantComprarItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantComprarItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataEntregaItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntregaItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataEntregaItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataEntregaItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataEntregaItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UM_Change()

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

Private Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemPed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemPed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ItemPed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ItemPed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ItemPed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPedidos)

End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPedidos)

End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = DataEntrega
    lErro = Grid_Campo_Libera_Foco(objGridPedidos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPedidos)

End Sub

Private Sub DataEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPedidos)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = DataEmissao
    lErro = Grid_Campo_Libera_Foco(objGridPedidos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPedidos)

End Sub

Private Sub FilialCliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPedidos)

End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = FilialCliente
    lErro = Grid_Campo_Libera_Foco(objGridPedidos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NomeRed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeRed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPedidos)

End Sub

Private Sub NomeRed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPedidos)

End Sub

Private Sub NomeRed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = NomeRed
    lErro = Grid_Campo_Libera_Foco(objGridPedidos)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub Pedido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Pedido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPedidos)

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPedidos)

End Sub

Private Sub Pedido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = Pedido
    lErro = Grid_Campo_Libera_Foco(objGridPedidos)
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

Private Sub Exclusivo_Change()

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

Private Sub TipoTribItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTribItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub TipoTribItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub TipoTribItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = TipoTribItem
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

Private Sub CentroCusto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CentroCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub CentroCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub CentroCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = CentroCusto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ClienteAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ClienteDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se  DataEmissaoAte foi preenchida
    If Len(Trim(DataEmissaoAte.Text)) = 0 Then Exit Sub

    'Critica DataDe
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 68100

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68100
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161471)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Verifica se  DataEmissaoDe foi preenchida
    If Len(Trim(DataEmissaoDe.Text)) = 0 Then Exit Sub

    'Critica DataDe
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 68099

    Exit Sub

Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68099
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161472)

    End Select

    Exit Sub

End Sub

Private Sub DataEntregaAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntregaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntregaAte_Validate

    'Verifica se  DataEntregaAte foi preenchida
    If Len(Trim(DataEntregaAte.Text)) = 0 Then Exit Sub

    'Critica DataEntregaAte
    lErro = Data_Critica(DataEntregaAte.Text)
    If lErro <> SUCESSO Then gError 68102

    Exit Sub

Erro_DataEntregaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68102
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161473)

    End Select

    Exit Sub

End Sub

Private Sub DataEntregaDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntregaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntregaDe_Validate

    'Verifica se  DataEntregaDe foi preenchida
    If Len(Trim(DataEntregaDe.Text)) = 0 Then Exit Sub

    'Critica DataEntregaDe
    lErro = Data_Critica(DataEntregaDe.Text)
    If lErro <> SUCESSO Then gError 68101

    Exit Sub

Erro_DataEntregaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68101
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161474)

    End Select

    Exit Sub

End Sub

Private Sub EnviaCompras_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ExibeTodos_Click()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    
    If ExibeTodos.Value = vbChecked Then
        
        LabelPedidoDe.Enabled = False
        LabelPedidoAte.Enabled = False
        LabelClienteAte.Enabled = False
        LabelClienteDe.Enabled = False
        Label23.Enabled = False
        Label24.Enabled = False
        Label7(0).Enabled = False
        Label8(2).Enabled = False
        PedidoInicial.Enabled = False
        PedidoFinal.Enabled = False
        ClienteDe.Enabled = False
        ClienteAte.Enabled = False
        DataEmissaoDe.Enabled = False
        DataEmissaoAte.Enabled = False
        DataEntregaDe.Enabled = False
        DataEntregaAte.Enabled = False
        UpDownEmissaoDe.Enabled = False
        UpDownEmissaoAte.Enabled = False
        UpDownEntregaAte.Enabled = False
        UpDownEntregaDe.Enabled = False
               
        PedidoInicial.Text = ""
        PedidoFinal.Text = ""
        ClienteDe.Text = ""
        ClienteAte.Text = ""
        DataEmissaoDe.PromptInclude = False
        DataEmissaoDe.Text = ""
        DataEmissaoDe.PromptInclude = True
        DataEmissaoAte.PromptInclude = False
        DataEmissaoAte.Text = ""
        DataEmissaoAte.PromptInclude = True
        DataEntregaDe.PromptInclude = False
        DataEntregaDe.Text = ""
        DataEntregaDe.PromptInclude = True
        DataEntregaAte.PromptInclude = False
        DataEntregaAte.Text = ""
        DataEntregaAte.PromptInclude = True
    
    Else
    
        LabelPedidoDe.Enabled = True
        LabelPedidoAte.Enabled = True
        LabelClienteAte.Enabled = True
        LabelClienteDe.Enabled = True
        Label23.Enabled = True
        Label24.Enabled = True
        Label7(0).Enabled = True
        Label8(2).Enabled = True
        PedidoInicial.Enabled = True
        PedidoFinal.Enabled = True
        ClienteDe.Enabled = True
        ClienteAte.Enabled = True
        DataEmissaoDe.Enabled = True
        DataEmissaoAte.Enabled = True
        DataEntregaDe.Enabled = True
        DataEntregaAte.Enabled = True
        UpDownEmissaoDe.Enabled = True
        UpDownEmissaoAte.Enabled = True
        UpDownEntregaAte.Enabled = True
        UpDownEntregaDe.Enabled = True
    End If
    

End Sub

Private Sub FilialCompra_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCompra_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialCompra.Text)) = 0 Then Exit Sub

    'Verifica se é uma FilialEmpresa selecionada
    If FilialCompra.Text = FilialCompra.List(FilialCompra.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCompra, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 68140

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68139

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 68141

        'coloca na tela
        FilialCompra.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 68142

    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case gErr

        Case 68139, 68140
            
        Case 68141
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialCompra.Text)

        Case 68142
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialCompra.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161475)

    End Select

    Exit Sub

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

Private Sub GridProdutos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos)
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


Private Sub GridPedido_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPedidos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If

End Sub

Private Sub GridPedido_GotFocus()
    Call Grid_Recebe_Foco(objGridPedidos)
End Sub

Private Sub GridPedido_EnterCell()
    Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
End Sub

Private Sub GridPedido_LeaveCell()
    Call Saida_Celula(objGridPedidos)
End Sub

Private Sub GridPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridPedidos)
End Sub

Private Sub GridPedido_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPedidos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If

End Sub

Private Sub GridPedido_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridPedidos)
End Sub

Private Sub GridPedido_RowColChange()
    Call Grid_RowColChange(objGridPedidos)
End Sub

Private Sub GridPedido_Scroll()
    Call Grid_Scroll(objGridPedidos)
End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

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

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    objCliente.lCodigo = StrParaLong(ClienteAte.Text)

    'Chama a tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)


End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    objCliente.lCodigo = StrParaLong(ClienteDe.Text)

    'Chama a tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)

End Sub

Private Sub LabelPedidoAte_Click()
'Chama a tela PedidoVendaLista

Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection

    objPedidoVenda.lCodigo = StrParaLong(PedidoFinal.Text)

    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoVenda, objEventoPedidoAte)

    Exit Sub

End Sub

Private Sub LabelPedidoDe_Click()
'Chama a tela PedidoVendaLista

Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection

    'Verifica se PedidoDe está preenchido
    objPedidoVenda.lCodigo = StrParaLong(PedidoInicial.Text)

    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoVenda, objEventoPedidoDe)

    Exit Sub

End Sub

Private Sub objEventoBotaoFiliaisFornProduto_evSelecao(obj1 As Object)

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
    If lErro <> SUCESSO Then gError 68130

    'Preenche campo Fornecedor
    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Lê coleção de códigos e nomes da Filial do Fornecedor
    lErro = CF("FornecedorProdutoFF_Le_FilialForn", objFornecedorProdutoFF.sProduto, objFornecedor.lCodigo, Codigo_Extrai(FilialCompra.Text), colCodigoNome)
    If lErro <> SUCESSO Then gError 68131

    'Se não encontrou nenhuma Filial, erro
    If colCodigoNome.Count = 0 Then gError 68132

    If objFornecedorProdutoFF.iFilialForn > 0 Then

        For iIndice = 1 To colCodigoNome.Count
            If colCodigoNome.Item(iIndice).iCodigo = objFornecedorProdutoFF.iFilialForn Then
                Exit For
            End If
        Next

        If iIndice > colCodigoNome.Count Then gError 68133

    ElseIf objFornecedorProdutoFF.iFilialForn = 0 Then
        objFornecedorProdutoFF.iFilialForn = colCodigoNome.Item(1).iCodigo
    End If

    For iIndice = 1 To colCodigoNome.Count
        If colCodigoNome.Item(iIndice).iCodigo = objFornecedorProdutoFF.iFilialForn Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_FilialForn_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
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

        Case 68130, 68131, 68132

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161476)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim objCliente As New ClassCliente

    Set objCliente = obj1

    'Coloca o código do Cliente na tela
    ClienteAte.Text = objCliente.lCodigo

    Me.Show
End Sub

Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim objCliente As New ClassCliente

    Set objCliente = obj1

    'Coloca o código do Cliente na tela
    ClienteDe.Text = objCliente.lCodigo

    Me.Show

End Sub

Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)

Dim objPedidoVenda As New ClassPedidoDeVenda

    Set objPedidoVenda = obj1

    'Coloca o Código do Pedido na tela
    PedidoFinal.Text = objPedidoVenda.lCodigo

    Me.Show

End Sub

Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)

Dim objPedidoVenda As New ClassPedidoDeVenda

    Set objPedidoVenda = obj1

    'Coloca o código do pedido na tela
    PedidoInicial.Text = objPedidoVenda.lCodigo

    Me.Show

End Sub


Private Sub objEventoRequisitante_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    'Colocao Nome Reduzido do Requisitante na tela
    Requisitante.Text = objRequisitante.sNomeReduzido

    'Dispara o Validate de Requisitante
    Call Requisitante_Validate(bSGECancelDummy)

    Me.Show

End Sub

Private Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordenados_Click()

Dim lErro As Long
Dim colPedidos As New Collection
Dim colCampos As New Collection
Dim colPedidosSaida As New Collection
Dim colPedMarcados As New Collection
Dim iIndice As Integer
Dim iLinha As Integer

On Error GoTo Erro_Ordenados_Click

    If gobjGeracaoReqPedVenda.sOrdenacao = "" Then Exit Sub
    
    'Verifica se Ordenacao da tela é diferente de gsOrdenacao
    If Ordenados.Text <> gobjGeracaoReqPedVenda.sOrdenacao Then

        'Guarda os Pedidos de Venda marcados
        For iIndice = 1 To objGridPedidos.iLinhasExistentes
            If GridPedido.TextMatrix(iIndice, iGrid_SelecionadoPed_Col) = "1" Then
                colPedMarcados.Add CLng(GridPedido.TextMatrix(iIndice, iGrid_Pedido_Col))
            End If
        Next
    
        'Recolhe os itens do GridPedidos
        lErro = GridPedidos_Recolhe(colPedidos)
        If lErro <> SUCESSO Then gError 68134

        Call Monta_Colecao_Campos_Pedido(colCampos, Ordenados.ListIndex)

        lErro = Ordena_Colecao(colPedidos, colPedidosSaida, colCampos)
        If lErro <> SUCESSO Then gError 68135

        'Devolve os elementos ordenados para o  GridPedidos
        lErro = GridPedidos_Devolve(colPedidosSaida)
        If lErro <> SUCESSO Then gError 68136

        'Marca novamente os Pedidos de Venda
        For iIndice = 1 To colPedMarcados.Count
            For iLinha = 1 To objGridPedidos.iLinhasExistentes
                If CStr(colPedMarcados(iIndice)) = GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) Then
                    GridPedido.TextMatrix(iLinha, iGrid_SelecionadoPed_Col) = "1"
                End If
            Next
        Next
        
    End If
    
    Call Grid_Refresh_Checkbox(objGridPedidos)

    Exit Sub

Erro_Ordenados_Click:

    Select Case gErr

        Case 68134, 68135, 68136
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161477)

    End Select

    Exit Sub

End Sub

Private Function GridPedidos_Recolhe(colPedido As Collection) As Long
'Recolhe os itens do GridPedidos e adiciona em colPedido

Dim objPedVenda As New ClassPedidoDeVenda
Dim iIndice As Integer
Dim objItemPV As New ClassItemPedido
Dim colReserva As New colReserva

On Error GoTo Erro_GridPedidos_Recolhe

    Set colPedido = New Collection

    'Percorre todas as linhas do GridPedidos
    For iIndice = 1 To objGridPedidos.iLinhasExistentes

        Set objPedVenda = New ClassPedidoDeVenda

        objPedVenda.lCodigo = StrParaLong(GridPedido.TextMatrix(iIndice, iGrid_Pedido_Col))
        objPedVenda.iFilial = Codigo_Extrai(GridPedido.TextMatrix(iIndice, iGrid_FilialCliente_Col))
        objPedVenda.dtDataEmissao = StrParaDate(GridPedido.TextMatrix(iIndice, iGrid_DataEmissao_Col))
        objItemPV.dtDataEntrega = StrParaDate(GridPedido.TextMatrix(iIndice, iGrid_DataEntrega_Col))
        
        objPedVenda.colItensPedido.Add objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.iItem, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sProdutoDescricao, objItemPV.dValorAbatComissao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, colReserva, objItemPV.sProdutoNomeReduzido, objItemPV.sUnidadeMed, objItemPV.sLote, objItemPV.sUMEstoque, objItemPV.iClasseUM, objItemPV.dQuantFaturada, objItemPV.dQuantAFaturar, objItemPV.sDescricao, objItemPV.iStatus, objItemPV.iControleEstoque, objItemPV.dQuantOP, objItemPV.dQuantSC, 0, 0, 0, 0, 0, 0, objItemPV.lNumIntDoc
        
        'Adiciona em colPedido
        colPedido.Add objPedVenda

    Next

    GridPedidos_Recolhe = SUCESSO

    Exit Function

Erro_GridPedidos_Recolhe:

    GridPedidos_Recolhe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161478)

    End Select

    Exit Function

End Function
Sub Monta_Colecao_Campos_Pedido(colCampos As Collection, iOrdenacao As Integer)

    Select Case iOrdenacao

        Case 0
            colCampos.Add "lCodigo"
            colCampos.Add "lCliente"
            colCampos.Add "iFilial"

        Case 1
            colCampos.Add "dtDataEmissao"
            colCampos.Add "lCodigo"
            
    End Select

End Sub

Private Function GridPedidos_Devolve(colPedido As Collection) As Long
'Devolve os elementos ordenados para o GridPedidos

Dim lErro As Long
Dim objPedVenda As New ClassPedidoDeVenda
Dim iLinha As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objItemPV As New ClassItemPedido

On Error GoTo Erro_GridPedidos_Devolve

    Call Grid_Limpa(objGridPedidos)

    For Each objPedVenda In colPedido

        iLinha = iLinha + 1
        
        'Preenche o GridPedidos
        GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) = objPedVenda.lCodigo
        If objPedVenda.dtDataEmissao <> DATA_NULA Then GridPedido.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objPedVenda.dtDataEmissao, "dd/mm/yyyy")
        
        For Each objItemPV In objPedVenda.colItensPedido
            If objItemPV.dtDataEntrega <> DATA_NULA Then
                GridPedido.TextMatrix(iLinha, iGrid_DataEntrega_Col) = Format(objItemPV.dtDataEntrega, "dd/mm/yyyy")
            End If
        Next
        
        If objPedVenda.lCliente <> 0 Then

            objCliente.lCodigo = objPedVenda.lCliente

            'Lê o Cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO Then gError 68298

            GridPedido.TextMatrix(iLinha, iGrid_NomeRed_Col) = objCliente.sNomeReduzido
        
        End If

        If objPedVenda.iFilial <> 0 Then
        
            objFilialCliente.iCodFilial = objPedVenda.iFilial
            objFilialCliente.lCodCliente = objPedVenda.lCliente
            
            'Lê a Filial Cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 68299
            'Se não encontrou a Filial ==> erro
            If lErro = 12567 Then gError 68300
            
            GridPedido.TextMatrix(iLinha, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        
        End If
        
    Next

    Call Grid_Refresh_Checkbox(objGridPedidos)

    objGridPedidos.iLinhasExistentes = iLinha

    GridPedidos_Devolve = SUCESSO

    Exit Function

Erro_GridPedidos_Devolve:

    GridPedidos_Devolve = gErr

    Select Case gErr

        Case 68298, 68299
            'Erros tratados nas rotinas chamadas
        
        Case 68300
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161479)

    End Select

    Exit Function

End Function

Private Sub PedidoFinal_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoInicial_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Requisitante_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Requisitante_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Requisitante_Validate

    'Se o Requisitante não foi preenchido, sai da rotina
    If Len(Trim(Requisitante.Text)) = 0 Then Exit Sub
    
    lErro = TP_Requisitante_Le(Requisitante, objRequisitante)
    If lErro <> SUCESSO Then gError 68194

    Exit Sub

Erro_Requisitante_Validate:

    Cancel = True

    Select Case gErr

        Case 68194
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161480)

    End Select

    Exit Sub

End Sub

Private Sub RequisitanteLabel_Click()

Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

    'Se o Requisitante estiver preenchido
    objRequisitante.sNomeReduzido = Requisitante.Text

    'Chama o Browser que Lista os Requisitantes
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoRequisitante)

End Sub

Private Sub Selecionado_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Selecionado_Click

    Set objPedidoVenda = gobjGeracaoReqPedVenda.colPedVenda(GridPedido.Row)
    
    'Verifica se o Pedido de Venda está selecionado
    If StrParaInt(GridPedido.TextMatrix(GridPedido.Row, iGrid_SelecionadoPed_Col)) = MARCADO Then

        'Traz para a tela os Itens do Pedido de Venda que foi selecionado
        lErro = Traz_ItensPed_Tela(objPedidoVenda)
        If lErro <> SUCESSO Then gError 68124

    ElseIf StrParaInt(GridPedido.TextMatrix(GridPedido.Row, iGrid_SelecionadoPed_Col)) = DESMARCADO Then

        objPedidoVenda.lCodigo = StrParaLong(GridPedido.TextMatrix(GridPedido.Row, iGrid_Pedido_Col))
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa

        'Retira da tela os Itens do Pedido de Venda que foi desselecionado
        lErro = Exclui_ItensPed_Tela(objPedidoVenda)
        If lErro <> SUCESSO Then gError 68125

    End If

    Exit Sub

Erro_Selecionado_Click:

    Select Case gErr

        Case 68124, 68125, 68272
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161481)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual invisivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

        'Se foi clicado no TAB_Selecao
        If TabStrip1.SelectedItem.Index = TAB_Selecao Then iFrameSelecaoAlterado = 0

            'Se o frame anterior foi o de Seleção e ele foi alterado
            If iFrameSelecaoAlterado = REGISTRO_ALTERADO Then

                Set gobjGeracaoReqPedVenda = New ClassGeracaoReqPedVenda
                'Limpa o Grid de Pedidos
                Call Grid_Limpa(objGridPedidos)
                
                'Limpa o Grid de Itens de Pedidos
                Call Grid_Limpa(objGridItens)
                
                'Limpa o Grid de Produtos
                Call Grid_Limpa(objGridProdutos)

                'Recolhe os dados do TAB_Selecao
                lErro = Move_TabSelecao_Memoria(gobjGeracaoReqPedVenda)
                If lErro <> SUCESSO Then gError 68143

                'Busca no BD todos os Pedidos de Venda com as características definidas no Tab Selecao
                lErro = CF("GeracaoReqPedVenda_Le_Pedidos", gobjGeracaoReqPedVenda)
                If lErro <> SUCESSO Then gError 68144

                'Traz os dados dos Pedidos de Venda para a tela
                lErro = Traz_Pedidos_Tela(gobjGeracaoReqPedVenda)
                If lErro <> SUCESSO Then gError 68145

            End If

            iFrameSelecaoAlterado = 0

            iFrameAlterado = 0

        End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 68143, 68144, 68145
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161482)

    End Select

    Exit Sub

End Sub

Private Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UpDownEmissaoAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68117

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 68117
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161483)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68122

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 68122
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161484)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68116

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 68116
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161485)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68121

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 68121
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161486)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Limpa_Tela_GeracaoReqPedVenda()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_GeracaoReqPedVenda

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
    ExibeTodos.Value = vbUnchecked
    Requisitante.Text = ""

    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridProdutos)
    Call Grid_Limpa(objGridPedidos)

    Set gobjGeracaoReqPedVenda = Nothing

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_GeracaoReqPedVenda:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161487)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_DownClick

    lErro = Data_Up_Down_Click(DataEntregaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68118

    Exit Sub

Erro_UpDownEntregaAte_DownClick:

    Select Case gErr

        Case 68118
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161488)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_UpClick

    lErro = Data_Up_Down_Click(DataEntregaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68120

    Exit Sub

Erro_UpDownEntregaAte_UpClick:

    Select Case gErr

        Case 68120
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161489)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEntregaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_DownClick

    lErro = Data_Up_Down_Click(DataEntregaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68119

    Exit Sub

Erro_UpDownEntregaDe_DownClick:

    Select Case gErr

        Case 68119
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161490)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_UpClick

    lErro = Data_Up_Down_Click(DataEntregaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68119

    Exit Sub

Erro_UpDownEntregaDe_UpClick:

    Select Case gErr

        Case 68119
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161491)

    End Select

    Exit Sub

End Sub

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label7(Index), Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelPedidoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoAte, Source, X, Y)
End Sub

Private Sub LabelPedidoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelPedidoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoDe, Source, X, Y)
End Sub

Private Sub LabelPedidoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoDe, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub RequisitanteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RequisitanteLabel, Source, X, Y)
End Sub

Private Sub RequisitanteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RequisitanteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Busca_Produto(sProduto As String, colProdutos As Collection, objProduto As ClassProduto, bAchou As Boolean)

Dim objProdAux As ClassProduto

    bAchou = False
    
    For Each objProdAux In colProdutos
        If objProdAux.sCodigo = sProduto Then
            Set objProduto = objProdAux
            bAchou = True
            Exit For
        End If
    
    Next

    Exit Sub

End Sub

Private Sub SelItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub SelItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub SelItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = SelItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

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
            If lErro <> SUCESSO And lErro <> 28030 Then gError 84552 '(65558)
    
            'Se não encontrou o Produto --> Erro
            If lErro = 28030 Then gError 84553 '(65560)
            
            'Se o Produto possui Estoque
            If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                'Lê dados do seu almoxarifado Padrão
                lErro = CF("AlmoxarifadoPadrao_Le", iCodFilial, objProduto.sCodigo, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then gError 84554 '(61710)
    
                'Se encontrou
                If lErro = SUCESSO And iAlmoxarifadoPadrao <> 0 Then
    
                    objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
    
                    'Lê os dados do Almoxarifado a partir do código passado
                    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25056 Then gError 84555 '(65559)
            
                    'Se não encontrou, erro
                    If lErro = 25056 Then gError 84556 '(65561)
            
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
        
        Case 84552, 84555, 84554
        
        Case 84553
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 84556
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161492)
    
    End Select
    
    Exit Function
    
End Function

Function SaldoAtualProdComprado(sCodProduto As String, dSaldoAtual As Double, Optional dQuantReservada As Double, Optional dQuantEmReq As Double, Optional dQuantEmPed As Double, Optional dSaldoDisponivel As Double, Optional dQuantEmpenhada As Double)
'Calcula o Saldo Atual para o produto comprado passado por parâmetro.

Dim lErro As Long
Dim lComando As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProduto As New ClassProduto

On Error GoTo Erro_SaldoAtualProdComprado
    
    'Lê a quantidade do produto em Requisição
    lErro = CF("Produto_Le_QuantRequisitada", sCodProduto, dQuantEmReq)
    If lErro <> SUCESSO Then gError 84584
    
    objEstoqueProduto.sProduto = sCodProduto
    'Lê o estoque do produto para a filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 52979 Then gError 84585
    
    'Calcula a quantidade do produto em Pedido de compras
    lErro = CF("Produto_Le_QuantEmPedidoCompra", sCodProduto, dQuantEmPed)
    If lErro <> SUCESSO Then gError 84586
             
    'Calcula o Saldo Atual do produto
    With objEstoqueProduto
        dSaldoDisponivel = .dQuantDisponivel + .dQuantRecIndl
        dQuantEmpenhada = .dQuantEmpenhada
        dQuantReservada = .dQuantReservada
        dSaldoAtual = objEstoqueProduto.dQuantReservada + dSaldoDisponivel + dQuantEmPed + dQuantEmReq - .dQuantEmpenhada
    End With
       
    SaldoAtualProdComprado = SUCESSO
    
    Exit Function
    
Erro_SaldoAtualProdComprado:

    SaldoAtualProdComprado = gErr
    
    Select Case gErr
    
        Case 84584, 84585, 84586
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161493)
            
    End Select
    
    Exit Function
    
End Function

Sub Recolhe_ItensPV(colItensRC As Collection, colPedidoVenda As Collection)

Dim objItemPV As ClassItemPedido
Dim objPedidoVenda As ClassPedidoDeVenda
Dim objItemRC As ClassItemReqCompras
    
    For Each objItemRC In colItensRC
        For Each objPedidoVenda In colPedidoVenda
        
        
        
            For Each objItemPV In objPedidoVenda.colItensPedido
                If objItemPV.sProduto = objItemRC.sProduto Then objItemRC.colItemPV.Add objItemPV
            Next
        Next
    Next

    Exit Sub
    
End Sub


Sub Busca_ItemPV(objItemPV As ClassItemPedido)

Dim objPedidoVenda As ClassPedidoDeVenda
Dim objItemPedido As ClassItemPedido


    For Each objPedidoVenda In gobjGeracaoReqPedVenda.colPedVenda
        
        If objPedidoVenda.lCodigo = objItemPV.lCodPedido And objPedidoVenda.iFilialEmpresa = objItemPV.iFilialEmpresa Then
            For Each objItemPedido In objPedidoVenda.colItensPedido
                If objItemPedido.sProduto = objItemPV.sProduto Then
                    objItemPV.lNumIntDoc = objItemPedido.lNumIntDoc
                    Exit Sub
                End If
            Next
        End If
        
    Next
    
    Exit Sub
        
End Sub

Private Sub BotaoEntrega_Click()

Dim sProdutoTela As String
Dim dQuantidade As Double
Dim colDataEntrega As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemPedido As New ClassItemPedido
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim lErro As Long
Dim lCodPV As Long


On Error GoTo Erro_BotaoEntrega_Click

    If GridItens.Row = 0 Then gError 183247

    sProdutoTela = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItem_Col)

    If Len(sProdutoTela) = 0 Then gError 183248

    objPedidoVenda.lCodigo = StrParaLong(GridItens.TextMatrix(GridItens.Row, iGrid_PedidoItem_Col))
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    lErro = CF("PedidoDeVenda_Le_Completo", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 27633 Then gError 183251

    'Coloca o Produto no formato do BD
    lErro = CF("Produto_Formata", sProdutoTela, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 183250

    For Each objItemPedido In objPedidoVenda.colItensPedido
        If objItemPedido.sProduto = sProdutoFormatado Then
            Exit For
        End If
    Next

    Call Chama_Tela_Modal("DataEntrega", sProdutoTela, dQuantidade, objItemPedido.colItemPVEntrega, DESABILITA_TELA)

    Exit Sub

Erro_BotaoEntrega_Click:

    Select Case gErr

        Case 183247
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 183248
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 183249, 183250, 183251

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183252)

    End Select

    Exit Sub

End Sub


