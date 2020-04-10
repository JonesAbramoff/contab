VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoPedCompraGerCotOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7995
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   930
      Width           =   16530
      Begin VB.Frame Frame6 
         Caption         =   "Exibe Gerações de Pedidos de Cotação"
         Height          =   4428
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   8310
         Begin VB.Frame Frame9 
            Caption         =   "Código"
            Height          =   1665
            Left            =   4548
            TabIndex        =   10
            Top             =   432
            Width           =   3432
            Begin MSMask.MaskEdBox CodigoDe 
               Height          =   312
               Left            =   1512
               TabIndex        =   12
               Top             =   444
               Width           =   816
               _ExtentX        =   1429
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoAte 
               Height          =   312
               Left            =   1512
               TabIndex        =   14
               Top             =   1032
               Width           =   816
               _ExtentX        =   1429
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label Label14 
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
               Height          =   192
               Left            =   1056
               TabIndex        =   11
               Top             =   528
               Width           =   312
            End
            Begin VB.Label Label12 
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
               Height          =   192
               Left            =   1056
               TabIndex        =   13
               Top             =   1092
               Width           =   360
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Data"
            Height          =   1665
            Left            =   372
            TabIndex        =   3
            Top             =   408
            Width           =   3432
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   2376
               TabIndex        =   6
               Top             =   468
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   2376
               TabIndex        =   9
               Top             =   1008
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   1212
               TabIndex        =   5
               Top             =   456
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   1212
               TabIndex        =   8
               Top             =   1008
               Width           =   1176
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label17 
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
               Height          =   192
               Left            =   744
               TabIndex        =   4
               Top             =   516
               Width           =   312
            End
            Begin VB.Label Label13 
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
               Height          =   192
               Left            =   768
               TabIndex        =   7
               Top             =   1056
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Local de Entrega"
            Height          =   1635
            Left            =   372
            TabIndex        =   15
            Top             =   2364
            Width           =   7605
            Begin VB.Frame Frame4 
               Caption         =   "Tipo"
               Height          =   585
               Index           =   0
               Left            =   3390
               TabIndex        =   17
               Top             =   180
               Width           =   4065
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
                  Left            =   585
                  TabIndex        =   18
                  Top             =   225
                  Value           =   -1  'True
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
                  Left            =   2280
                  TabIndex        =   19
                  Top             =   225
                  Width           =   1335
               End
            End
            Begin VB.CheckBox SelecionaDestino 
               Caption         =   "Seleciona Local Entrega"
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
               Left            =   324
               TabIndex        =   16
               Top             =   315
               Width           =   2445
            End
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   675
               Index           =   0
               Left            =   3600
               TabIndex        =   20
               Top             =   885
               Width           =   3645
               Begin VB.ComboBox FilialEmpresa 
                  Enabled         =   0   'False
                  Height          =   288
                  Left            =   1230
                  TabIndex        =   22
                  Top             =   150
                  Width           =   2160
               End
               Begin VB.Label FilEmprDestLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "Filial:"
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
                  Height          =   195
                  Left            =   720
                  TabIndex        =   21
                  Top             =   180
                  Width           =   465
               End
            End
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Height          =   675
               Index           =   1
               Left            =   3570
               TabIndex        =   23
               Top             =   885
               Visible         =   0   'False
               Width           =   3645
               Begin VB.ComboBox FilialFornec 
                  Enabled         =   0   'False
                  Height          =   288
                  Left            =   1260
                  TabIndex        =   27
                  Top             =   360
                  Width           =   2160
               End
               Begin MSMask.MaskEdBox Fornecedor 
                  Height          =   300
                  Left            =   1245
                  TabIndex        =   25
                  Top             =   15
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   20
                  PromptChar      =   " "
               End
               Begin VB.Label FilFornDestLabel 
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
                  TabIndex        =   26
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
                  Left            =   150
                  MousePointer    =   14  'Arrow and Question
                  TabIndex        =   24
                  Top             =   60
                  Width           =   1035
               End
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7950
      Index           =   2
      Left            =   135
      TabIndex        =   28
      Top             =   1005
      Visible         =   0   'False
      Width           =   16605
      Begin VB.Frame frame 
         Caption         =   "Gerações Pedidos Cotação"
         Height          =   7110
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   14670
         Begin MSMask.MaskEdBox DescricaoCot 
            Height          =   225
            Left            =   4770
            TabIndex        =   36
            Top             =   195
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.OptionButton Selecionado 
            Height          =   240
            Left            =   1350
            TabIndex        =   33
            Top             =   240
            Width           =   1155
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   225
            Left            =   3660
            TabIndex        =   35
            Top             =   210
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
         Begin MSMask.MaskEdBox CodCotacao 
            Height          =   225
            Left            =   2505
            TabIndex        =   34
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridPedCotacao 
            Height          =   3270
            Left            =   225
            TabIndex        =   32
            Top             =   345
            Width           =   14265
            _ExtentX        =   25162
            _ExtentY        =   5768
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.ComboBox OrdenacaoPedCot 
         Height          =   315
         ItemData        =   "GeracaoPedCompraGerCotOcx.ctx":0000
         Left            =   2970
         List            =   "GeracaoPedCompraGerCotOcx.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   150
         Width           =   2325
      End
      Begin VB.Label Label32 
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
         Left            =   1860
         TabIndex        =   29
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7950
      Index           =   3
      Left            =   150
      TabIndex        =   37
      Top             =   1035
      Visible         =   0   'False
      Width           =   16605
      Begin VB.ComboBox OrdenacaoReq 
         Height          =   315
         ItemData        =   "GeracaoPedCompraGerCotOcx.ctx":0004
         Left            =   2610
         List            =   "GeracaoPedCompraGerCotOcx.ctx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   135
         Width           =   2325
      End
      Begin VB.Frame Frame7 
         Caption         =   "Requisições de Compra"
         Height          =   6705
         Left            =   45
         TabIndex        =   41
         Top             =   525
         Width           =   16500
         Begin VB.TextBox ObservacaoReq 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   51
            Top             =   2985
            Width           =   6810
         End
         Begin VB.CheckBox Urgente 
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
            Height          =   225
            Left            =   5910
            TabIndex        =   48
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox EscolhidoReq 
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
            Left            =   360
            TabIndex        =   43
            Top             =   315
            Width           =   975
         End
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   240
            Left            =   6345
            TabIndex        =   49
            Top             =   390
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialReq 
            Height          =   225
            Left            =   1155
            TabIndex        =   44
            Top             =   360
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CclReq 
            Height          =   225
            Left            =   270
            TabIndex        =   50
            Top             =   2985
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   225
            Left            =   3585
            TabIndex        =   46
            Top             =   360
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
         Begin MSMask.MaskEdBox CodigoReq 
            Height          =   225
            Left            =   2730
            TabIndex        =   45
            Top             =   360
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataReq 
            Height          =   225
            Left            =   4755
            TabIndex        =   47
            Top             =   375
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
         Begin MSFlexGridLib.MSFlexGrid GridRequisicoes 
            Height          =   2655
            Left            =   150
            TabIndex        =   42
            Top             =   285
            Width           =   16155
            _ExtentX        =   28496
            _ExtentY        =   4683
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoReqCompras 
         Caption         =   "Requisição de Compras..."
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
         Left            =   6375
         TabIndex        =   40
         Top             =   60
         Width           =   2040
      End
      Begin VB.CommandButton BotaoDesmarcarTodosReq 
         Caption         =   "Desmarcar Todos"
         Height          =   555
         Left            =   1665
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   7305
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodosReq 
         Caption         =   "Marcar Todos"
         Height          =   555
         Left            =   45
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":11EA
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   7305
         Width           =   1425
      End
      Begin VB.Label Label57 
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
         Left            =   1470
         TabIndex        =   38
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7935
      Index           =   4
      Left            =   135
      TabIndex        =   54
      Top             =   1035
      Visible         =   0   'False
      Width           =   16590
      Begin VB.Frame Frame5 
         Caption         =   "Itens de Requisições"
         Height          =   7215
         Left            =   45
         TabIndex        =   55
         Top             =   -30
         Width           =   16530
         Begin VB.TextBox DescProdutoItemRC 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4980
            MaxLength       =   50
            TabIndex        =   62
            Top             =   330
            Width           =   4000
         End
         Begin VB.CheckBox EscolhidoItem 
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
            Left            =   300
            TabIndex        =   57
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox ObservacaoItemRC 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6975
            MaxLength       =   255
            TabIndex        =   73
            Top             =   3315
            Width           =   2355
         End
         Begin MSMask.MaskEdBox CclItemRC 
            Height          =   225
            Left            =   2295
            TabIndex        =   69
            Top             =   3210
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ExclusivoItemRC 
            Height          =   225
            Left            =   5745
            TabIndex        =   72
            Top             =   3165
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   990
            TabIndex        =   68
            Top             =   3300
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   225
            Left            =   -90
            TabIndex        =   67
            Top             =   3255
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox QuantPedida 
            Height          =   225
            Left            =   1440
            TabIndex        =   66
            Top             =   2970
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox FilialFornItemRC 
            Height          =   225
            Left            =   4470
            TabIndex        =   71
            Top             =   3270
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
         Begin MSMask.MaskEdBox FornecedorItemRC 
            Height          =   225
            Left            =   3090
            TabIndex        =   70
            Top             =   3075
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
         Begin MSMask.MaskEdBox QuantidadeItemRC 
            Height          =   225
            Left            =   240
            TabIndex        =   65
            Top             =   2985
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox CodigoReqItem 
            Height          =   225
            Left            =   1830
            TabIndex        =   59
            Top             =   255
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialReqItem 
            Height          =   225
            Left            =   915
            TabIndex        =   58
            Top             =   285
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Item 
            Height          =   225
            Left            =   2970
            TabIndex        =   60
            Top             =   225
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UnidadeMedItemRC 
            Height          =   225
            Left            =   6480
            TabIndex        =   63
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
         Begin MSMask.MaskEdBox QuantComprarItemRC 
            Height          =   225
            Left            =   7665
            TabIndex        =   64
            Top             =   345
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
         Begin MSMask.MaskEdBox ProdutoItemRC 
            Height          =   225
            Left            =   3795
            TabIndex        =   61
            Top             =   300
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItensRequisicoes 
            Height          =   3135
            Left            =   90
            TabIndex        =   56
            Top             =   330
            Width           =   16365
            _ExtentX        =   28866
            _ExtentY        =   5530
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoDesmarcarTodosItensRC 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   1785
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":2204
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   7275
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodosItensRC 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   45
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":33E6
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   7275
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8070
      Index           =   5
      Left            =   90
      TabIndex        =   76
      Top             =   930
      Visible         =   0   'False
      Width           =   16695
      Begin VB.Frame FrameProdutos 
         BorderStyle     =   0  'None
         Height          =   7050
         Index           =   2
         Left            =   150
         TabIndex        =   90
         Top             =   375
         Visible         =   0   'False
         Width           =   16275
         Begin VB.TextBox DescProduto2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2430
            MaxLength       =   50
            TabIndex        =   93
            Top             =   270
            Width           =   4000
         End
         Begin MSMask.MaskEdBox FilialDestino 
            Height          =   225
            Left            =   540
            TabIndex        =   98
            Top             =   3540
            Width           =   1065
            _ExtentX        =   1879
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
         Begin MSMask.MaskEdBox Destino 
            Height          =   225
            Left            =   7050
            TabIndex        =   97
            Top             =   270
            Width           =   1065
            _ExtentX        =   1879
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
         Begin MSMask.MaskEdBox TipoDestinoProd 
            Height          =   225
            Left            =   6000
            TabIndex        =   96
            Top             =   300
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
         Begin MSMask.MaskEdBox Fornecedor2 
            Height          =   225
            Left            =   1725
            TabIndex        =   99
            Top             =   3555
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
         Begin MSMask.MaskEdBox UnidadeMed2 
            Height          =   225
            Left            =   3945
            TabIndex        =   94
            Top             =   270
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialForn2 
            Height          =   225
            Left            =   3615
            TabIndex        =   100
            Top             =   3585
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
         Begin MSMask.MaskEdBox Quantidade2 
            Height          =   225
            Left            =   4995
            TabIndex        =   95
            Top             =   270
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
         Begin MSMask.MaskEdBox Produto2 
            Height          =   225
            Left            =   1155
            TabIndex        =   92
            Top             =   270
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos2 
            Height          =   3195
            Left            =   285
            TabIndex        =   91
            Top             =   435
            Width           =   15810
            _ExtentX        =   27887
            _ExtentY        =   5636
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame FrameProdutos 
         BorderStyle     =   0  'None
         Height          =   7065
         Index           =   1
         Left            =   165
         TabIndex        =   78
         Top             =   375
         Width           =   16275
         Begin VB.CheckBox EscolhidoProduto 
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
            Left            =   540
            TabIndex        =   80
            Top             =   240
            Width           =   990
         End
         Begin VB.TextBox DescProduto1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2535
            MaxLength       =   50
            TabIndex        =   82
            Top             =   270
            Width           =   4000
         End
         Begin VB.CommandButton BotaoMarcarTodosProd 
            Caption         =   "Marcar Todos"
            Height          =   570
            Left            =   165
            Picture         =   "GeracaoPedCompraGerCotOcx.ctx":4400
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   6465
            Width           =   1425
         End
         Begin VB.CommandButton BotaoDesmarcarTodosProd 
            Caption         =   "Desmarcar Todos"
            Height          =   570
            Left            =   1815
            Picture         =   "GeracaoPedCompraGerCotOcx.ctx":541A
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   6465
            Width           =   1425
         End
         Begin MSMask.MaskEdBox QuantUrgente 
            Height          =   225
            Left            =   6195
            TabIndex        =   85
            Top             =   270
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   5
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
         Begin MSMask.MaskEdBox UnidadeMed1 
            Height          =   225
            Left            =   4005
            TabIndex        =   83
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialForn1 
            Height          =   225
            Left            =   3060
            TabIndex        =   87
            Top             =   2835
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
         Begin MSMask.MaskEdBox Fornecedor1 
            Height          =   225
            Left            =   1005
            TabIndex        =   86
            Top             =   2235
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
         Begin MSMask.MaskEdBox Quantidade1 
            Height          =   225
            Left            =   5130
            TabIndex        =   84
            Top             =   330
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox Produto1 
            Height          =   225
            Left            =   1260
            TabIndex        =   81
            Top             =   270
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos1 
            Height          =   6165
            Left            =   105
            TabIndex        =   79
            Top             =   120
            Width           =   16125
            _ExtentX        =   28443
            _ExtentY        =   10874
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoEditarProduto 
         Caption         =   "Produto..."
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
         Left            =   135
         TabIndex        =   101
         Top             =   7620
         Width           =   1395
      End
      Begin MSComctlLib.TabStrip TabProdutos 
         Height          =   7470
         Left            =   75
         TabIndex        =   77
         Top             =   60
         Width           =   16485
         _ExtentX        =   29078
         _ExtentY        =   13176
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Seleção"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Quantidades por Destino"
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
      Height          =   8010
      Index           =   6
      Left            =   150
      TabIndex        =   102
      Top             =   945
      Visible         =   0   'False
      Width           =   16650
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
         Height          =   345
         Left            =   6345
         TabIndex        =   105
         Top             =   105
         Width           =   2205
      End
      Begin VB.Frame Frame4 
         Caption         =   "Opção"
         Height          =   1170
         Index           =   1
         Left            =   11985
         TabIndex        =   141
         Top             =   6795
         Width           =   4500
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
            Height          =   375
            Left            =   480
            TabIndex        =   142
            Top             =   255
            Width           =   2670
         End
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
            Height          =   375
            Left            =   480
            TabIndex        =   143
            Top             =   690
            Width           =   2670
         End
      End
      Begin VB.ComboBox OrdenacaoCot 
         Height          =   315
         ItemData        =   "GeracaoPedCompraGerCotOcx.ctx":65FC
         Left            =   2310
         List            =   "GeracaoPedCompraGerCotOcx.ctx":65FE
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   150
         Width           =   2325
      End
      Begin VB.Frame FrameCotacoes 
         Caption         =   "Cotações"
         Height          =   6180
         Index           =   2
         Left            =   45
         TabIndex        =   106
         Top             =   510
         Width           =   16425
         Begin VB.ComboBox Moeda 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   1215
            Width           =   1608
         End
         Begin MSMask.MaskEdBox PrecoUnitarioReal 
            Height          =   228
            Left            =   2976
            TabIndex        =   149
            Top             =   1248
            Width           =   1692
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TaxaForn 
            Height          =   225
            Left            =   4245
            TabIndex        =   150
            Top             =   1380
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox Cotacao 
            Height          =   225
            Left            =   5550
            TabIndex        =   151
            Top             =   1380
            Width           =   1080
            _ExtentX        =   1905
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
         Begin VB.ComboBox MotivoEscolhaCot 
            Height          =   315
            Left            =   6360
            TabIndex        =   130
            Text            =   "MotivoEscolhaCot"
            Top             =   1905
            Width           =   1995
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
            Left            =   465
            TabIndex        =   108
            Top             =   240
            Width           =   840
         End
         Begin VB.TextBox DescProdutoCot 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2535
            MaxLength       =   50
            TabIndex        =   112
            Top             =   270
            Width           =   4000
         End
         Begin VB.ComboBox TipoTributacaoCot 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2235
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   990
            Width           =   2565
         End
         Begin MSMask.MaskEdBox DataCotacao 
            Height          =   225
            Left            =   4245
            TabIndex        =   114
            Top             =   270
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaICMS 
            Height          =   225
            Left            =   825
            TabIndex        =   110
            Top             =   135
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   228
            Left            =   2928
            TabIndex        =   124
            Top             =   2316
            Width           =   1476
            _ExtentX        =   2619
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PedCotacao 
            Height          =   228
            Left            =   7260
            TabIndex        =   131
            Top             =   2172
            Width           =   1308
            _ExtentX        =   2302
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
            Left            =   180
            TabIndex        =   118
            Top             =   2310
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeEntrega 
            Height          =   225
            Left            =   4935
            TabIndex        =   127
            Top             =   2235
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
         Begin MSMask.MaskEdBox ValorPresente 
            Height          =   228
            Left            =   4092
            TabIndex        =   126
            Top             =   2256
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   423
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
         Begin MSMask.MaskEdBox DataNecessidade 
            Height          =   225
            Left            =   3450
            TabIndex        =   125
            Top             =   2115
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   2445
            TabIndex        =   122
            Top             =   2325
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrazoEntrega 
            Height          =   225
            Left            =   1260
            TabIndex        =   121
            Top             =   2160
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
            Left            =   6585
            TabIndex        =   129
            Top             =   2295
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
         Begin MSMask.MaskEdBox CondPagto 
            Height          =   225
            Left            =   1245
            TabIndex        =   120
            Top             =   2310
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   30
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UnidadeMedCot 
            Height          =   225
            Left            =   4005
            TabIndex        =   113
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFornCot 
            Height          =   225
            Left            =   285
            TabIndex        =   119
            Top             =   2175
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
            Left            =   6195
            TabIndex        =   116
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
         Begin MSMask.MaskEdBox QuantidadeCot 
            Height          =   225
            Left            =   5160
            TabIndex        =   115
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox ProdutoCot 
            Height          =   225
            Left            =   1260
            TabIndex        =   111
            Top             =   270
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCotacoes 
            Height          =   1845
            Left            =   60
            TabIndex        =   107
            Top             =   300
            Width           =   16305
            _ExtentX        =   28760
            _ExtentY        =   3254
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox Preferencia 
            Height          =   225
            Left            =   6060
            TabIndex        =   128
            Top             =   2280
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox ValorItem 
            Height          =   255
            Left            =   1980
            TabIndex        =   123
            Top             =   2145
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
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
            Left            =   15
            TabIndex        =   109
            Top             =   0
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
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   3405
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":6600
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Numeração Automática"
         Top             =   7635
         Width           =   300
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   465
         Left            =   2265
         TabIndex        =   137
         Top             =   7110
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   820
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
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
         Left            =   1170
         TabIndex        =   103
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   132
         Top             =   6825
         Width           =   1845
      End
      Begin VB.Label TotalItens 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2265
         TabIndex        =   133
         Top             =   6780
         Width           =   1155
      End
      Begin VB.Label TaxaEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5835
         TabIndex        =   135
         Top             =   6795
         Width           =   1155
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
         Left            =   4335
         TabIndex        =   134
         Top             =   6825
         Width           =   1455
      End
      Begin VB.Label Label31 
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
         Left            =   990
         TabIndex        =   138
         Top             =   7680
         Width           =   1215
      End
      Begin VB.Label Concorrencia 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2250
         TabIndex        =   139
         Top             =   7635
         Width           =   1155
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
         Left            =   1290
         TabIndex        =   136
         Top             =   7170
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   15225
      ScaleHeight     =   480
      ScaleWidth      =   1575
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   90
      Width           =   1635
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":66EA
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":6868
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "GeracaoPedCompraGerCotOcx.ctx":6D9A
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8460
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   16800
      _ExtentX        =   29633
      _ExtentY        =   14923
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerações Ped Cotação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens de Requisições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cotações"
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
   Begin VB.Label Label3 
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
      Left            =   210
      TabIndex        =   153
      Top             =   135
      Width           =   975
   End
   Begin VB.Label Comprador 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1275
      TabIndex        =   152
      Top             =   135
      Width           =   2145
   End
End
Attribute VB_Name = "GeracaoPedCompraGerCotOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis Globais
Dim iFrameAtual As Integer
Dim giPodeAumentarQuant As Integer
Dim iFrameProdutoAtual As Integer
Dim iAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim iFornecedorAlterado As Integer
Dim iFrameSelecaoAlterado As Integer
Dim iFrameDestinoAtual As Integer
Dim gobjGeracaoPedCompraCot As ClassGeracaoPedCompraCot
Dim gcolRequisicaoCompra As Collection
Dim gcolItemConcorrencia As Collection
Dim gColCotacoes As Collection
Dim gsTipoTributacao As String
Dim iCotacaoAlterada As Integer
Dim gsOrdenacao As String

'GridPedCotações
Dim objGridCotacao As AdmGrid
Dim iGrid_SelecionadoPed_Col As Integer
Dim iGrid_CodigoPed_Col As Integer
Dim iGrid_DataPed_Col As Integer
Dim iGrid_DescricaoCot_Col As Integer

'GridRequisicoes
Dim objGridRequisicoes As AdmGrid
Dim iGrid_EscolhidoReq_Col As Integer
Dim iGrid_FilialReq_Col As Integer
Dim iGrid_CodigoReq_Col As Integer
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_DataReq_Col As Integer
Dim iGrid_Urgente_Col As Integer
Dim iGrid_Requisitante_Col As Integer
Dim iGrid_CclReq_Col As Integer
Dim iGrid_ObservacaoReq_Col As Integer

'GridItensRequisicoes
Dim objGridItensRequisicoes As AdmGrid
Dim iGrid_EscolhidoItem_Col As Integer
Dim iGrid_FilialReqItem_Col As Integer
Dim iGrid_CodigoReqItem_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_ProdutoItemRC_Col As Integer
Dim iGrid_DescProdutoItem_Col As Integer
Dim iGrid_UnidadeMedItem_Col As Integer
Dim iGrid_QuantComprarItem_Col As Integer
Dim iGrid_QuantidadeItem_Col As Integer
Dim iGrid_QuantPedida_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_CclItemRC_Col As Integer
Dim iGrid_FornecedorItemRC_Col As Integer
Dim iGrid_FilialFornItemRC_Col As Integer
Dim iGrid_ExclusivoItemRC_Col As Integer
Dim iGrid_ObservacaoItemRC_Col As Integer

'GridProdutos1
Dim objGridProdutos1 As AdmGrid
Dim iGrid_EscolhidoProduto_Col As Integer
Dim iGrid_Produto1_Col As Integer
Dim iGrid_DescProduto1_Col As Integer
Dim iGrid_UnidadeMed1_Col As Integer
Dim iGrid_Quantidade1_Col As Integer
Dim iGrid_QuantUrgente_Col As Integer
Dim iGrid_Fornecedor1_Col As Integer
Dim iGrid_FilialForn1_Col As Integer

'GridProdutos2
Dim objGridProdutos2 As AdmGrid
Dim iGrid_Produto2_Col As Integer
Dim iGrid_DescProduto2_Col As Integer
Dim iGrid_UnidadeMed2_Col As Integer
Dim iGrid_Quantidade2_Col As Integer
Dim iGrid_TipoDestino_Col As Integer
Dim iGrid_Destino_Col As Integer
Dim iGrid_FilialDestino_Col As Integer
Dim iGrid_Fornecedor2_Col As Integer
Dim iGrid_FilialForn2_Col As Integer

'GridCotacoes
Dim objGridCotacoes As AdmGrid
Dim iGrid_EscolhidoCot_Col As Integer
Dim iGrid_ProdutoCot_Col As Integer
Dim iGrid_DescProdutoCot_Col As Integer
Dim iGrid_UMCot_Col As Integer
Dim iGrid_QuantidadeCot_Col As Integer
Dim iGrid_FornecedorCot_Col As Integer
Dim iGrid_FilialFornCot_Col As Integer
Dim iGrid_CondPagtoCot_Col As Integer
Dim iGrid_PrecoUnitarioCot_Col As Integer
Dim iGrid_ValorPresenteCot_Col As Integer
Dim iGrid_TipoTributacaoCot_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_PedidoCot_Col As Integer
Dim iGrid_DataValidadeCot_Col As Integer
Dim iGrid_PrazoEntrega_Col As Integer
Dim iGrid_DataNecessidade_Col As Integer
Dim iGrid_QuantidadeEntrega_Col As Integer
Dim iGrid_Preferencia_Col As Integer
Dim iGrid_QuantComprarCot_Col As Integer
Dim iGrid_MotivoEscolhaCot_Col As Integer
Dim iGrid_DataEntrega_Col As Integer
Dim iGrid_ValorItem_Col As Integer
Dim iGrid_DataCotacaoCot_Col As Integer
Dim iGrid_Moeda_Col As Integer
Dim iGrid_PrecoUnitario_RS_Col As Integer
Dim iGrid_TaxaForn_Col As Integer
Dim iGrid_CotacaoMoeda_Col As Integer

'Eventos da Tela
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoPedCotacao As AdmEvento
Attribute objEventoPedCotacao.VB_VarHelpID = -1

'Constantes Públicas dos tabs
Private Const TAB_Selecao = 1
Private Const TAB_PedCotacao = 2
Private Const TAB_REQUISICOES = 3
Private Const TAB_ITENSREQ = 4
Private Const TAB_Produtos = 5
Private Const TAB_COTACOES = 6

'Property Variables:
Dim m_Caption As String
Event Unload()

Function Trata_Parametros(Optional objGeracaoPedCompraCot As ClassGeracaoPedCompraCot)

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuario As New ClassUsuario
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameProdutoAtual = 1
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    '###########################
    'Inserido por Wagner
    Call Formata_Controles
    '###########################


    'Inicializa as variáveis globais
    Set objEventoFornecedor = New AdmEvento
    Set objEventoPedCotacao = New AdmEvento

    Set objGridCotacao = New AdmGrid
    Set objGridRequisicoes = New AdmGrid
    Set objGridItensRequisicoes = New AdmGrid
    Set objGridProdutos1 = New AdmGrid
    Set objGridProdutos2 = New AdmGrid
    Set objGridCotacoes = New AdmGrid

    Set gcolRequisicaoCompra = New Collection
    Set gColCotacoes = New Collection

    objComprador.sCodUsuario = gsUsuario

    'Verifica se gsUsuario é comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 67044

    'Se gsUsuario nao é comprador==> erro
    If lErro = 50059 Then gError 67061
    giPodeAumentarQuant = objComprador.iAumentaQuant

    objUsuario.sCodUsuario = objComprador.sCodUsuario

    'Lê o usuário
    lErro = CF("Usuario_Le", objUsuario)
    If lErro <> SUCESSO And lErro <> 36347 Then gError 67045

    'Se não encontrou ==>erro
    If lErro = 36347 Then gError 67062

    'Coloca o Nome Reduzido do Comprador na tela
    Comprador.Caption = objUsuario.sNomeReduzido

    'Preenche as combos de Ordenação
    Call OrdemPedCotacao_Carrega
    Call OrdemRequisicao_Carrega
    Call OrdemCotacao_Carrega

    'Inicializa as máscaras dos Produtos
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoItemRC)
    If lErro <> SUCESSO Then gError 67046

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto1)
    If lErro <> SUCESSO Then gError 67047

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto2)
    If lErro <> SUCESSO Then gError 67048

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoCot)
    If lErro <> SUCESSO Then gError 67049

    'Inicializa mascara do Ccl
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 67050

    'Coloca as Quantidades da tela no formato de Estoque
    QuantComprarItemRC.Format = FORMATO_ESTOQUE
    QuantPedida.Format = FORMATO_ESTOQUE
    QuantRecebida.Format = FORMATO_ESTOQUE
    Quantidade1.Format = FORMATO_ESTOQUE
    Quantidade2.Format = FORMATO_ESTOQUE
    QuantidadeCot.Format = FORMATO_ESTOQUE
    QuantComprarCot.Format = FORMATO_ESTOQUE
    QuantidadeItemRC.Format = FORMATO_ESTOQUE

    'Carrega Motivos de Escolha
    lErro = Carrega_MotivoEscolha()
    If lErro <> SUCESSO Then gError 67051

    'Inicializa o GridConcorrencia
    lErro = Inicializa_Grid_Cotacao(objGridCotacao)
    If lErro <> SUCESSO Then gError 67053

    'Inicializa o GridRequisicoes
    lErro = Inicializa_Grid_Requisicoes(objGridRequisicoes)
    If lErro <> SUCESSO Then gError 67052

    'Inicializa o GridItensRequisicoes
    lErro = Inicializa_Grid_ItensRequisicoes(objGridItensRequisicoes)
    If lErro <> SUCESSO Then gError 67054

    'Inicializa o GridProdutos1
    lErro = Inicializa_Grid_Produtos1(objGridProdutos1)
    If lErro <> SUCESSO Then gError 67055

    'Inicializa o GridProdutos2
    lErro = Inicializa_Grid_Produtos2(objGridProdutos2)
    If lErro <> SUCESSO Then gError 67056

    'Inicializa o GridProdutos2
    lErro = Inicializa_Grid_Cotacoes(objGridCotacoes)
    If lErro <> SUCESSO Then gError 67057

    'Carrega Tipos de Tributação
    lErro = Carrega_TipoTributacao()
    If lErro <> SUCESSO Then gError 66122

    'Carrega a combo FilialEmpresa
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 63860
    
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError 108981

    'Coloca Taxa Financeira na tela
    TaxaEmpresa.Caption = Format(gobjCOM.dTaxaFinanceiraEmpresa, "Percent")

    SelecionaDestino.Value = vbChecked
    
    'Coloca FiliaEmpresa Default na Tela
    iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126948
    
    FilialEmpresa.Text = iFilialEmpresa
    
    Call FilialEmpresa_Validate(bSGECancelDummy)
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 67044, 67045, 67046, 67047, 67048, 67049, 67050, 67051, 67052, 67053, 67054, 67055, 67056, 67057, 67058, 67059, 67060, 67297, 70508, 108981, 126948
            'Erros tratados nas rotinas chamadas

        Case 67061
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objUsuario.sCodUsuario)

        Case 67062
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuario.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161100)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub OrdemPedCotacao_Carrega()
'preenche a combo OrdenacaoPedCot

    OrdenacaoPedCot.Clear

    OrdenacaoPedCot.AddItem "Código"
    OrdenacaoPedCot.AddItem "Data"

    'Seleciona código como ordenação default
    OrdenacaoPedCot.ListIndex = 0

End Sub

Private Sub OrdemRequisicao_Carrega()
'preenche a combo OrdenacaoReq

    OrdenacaoReq.Clear

    OrdenacaoReq.AddItem "Número"
    OrdenacaoReq.AddItem "Data Limite"
    OrdenacaoReq.AddItem "Data da Requisição"

    'Seleciona Número como ordenação default
    OrdenacaoReq.ListIndex = 0

End Sub

Private Sub OrdemCotacao_Carrega()

    OrdenacaoCot.Clear

    OrdenacaoCot.AddItem "Produto"
    OrdenacaoCot.AddItem "Fornecedor"

    'Seleciona Produto como ordenação default
    OrdenacaoCot.ListIndex = 0

End Sub

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 66721

    CclReq.Mask = sMascaraCcl
    CclItemRC.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = gErr

    Select Case gErr

        Case 66721
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161101)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Cotacao(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid de Pedidos de Cotação

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Cotacao

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Selecionado")
    objGridInt.colColuna.Add ("Cotação")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Descrição")

    'campos de edição do grid
    objGridInt.colCampo.Add (Selecionado.Name)
    objGridInt.colCampo.Add (CodCotacao.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (DescricaoCot.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_SelecionadoPed_Col = 1
    iGrid_CodigoPed_Col = 2
    iGrid_DataPed_Col = 3
    iGrid_DescricaoCot_Col = 4

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridPedCotacao

    GridPedCotacao.ColWidth(0) = 400
    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_COTACOES + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 22

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cotacao = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Cotacao:

    Inicializa_Grid_Cotacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161102)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_ItensRequisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_ItensRequisicoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Requisição")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Unid. Med.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Em Pedido")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Centro C/L")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusividade")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoItem.Name)
    objGridInt.colCampo.Add (FilialReqItem.Name)
    objGridInt.colCampo.Add (CodigoReqItem.Name)
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (ProdutoItemRC.Name)
    objGridInt.colCampo.Add (DescProdutoItemRC.Name)
    objGridInt.colCampo.Add (UnidadeMedItemRC.Name)
    objGridInt.colCampo.Add (QuantComprarItemRC.Name)
    objGridInt.colCampo.Add (QuantidadeItemRC.Name)
    objGridInt.colCampo.Add (QuantPedida.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CclItemRC.Name)
    objGridInt.colCampo.Add (FornecedorItemRC.Name)
    objGridInt.colCampo.Add (FilialFornItemRC.Name)
    objGridInt.colCampo.Add (ExclusivoItemRC.Name)
    objGridInt.colCampo.Add (ObservacaoItemRC.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoItem_Col = 1
    iGrid_FilialReqItem_Col = 2
    iGrid_CodigoReqItem_Col = 3
    iGrid_Item_Col = 4
    iGrid_ProdutoItemRC_Col = 5
    iGrid_DescProdutoItem_Col = 6
    iGrid_UnidadeMedItem_Col = 7
    iGrid_QuantComprarItem_Col = 8
    iGrid_QuantidadeItem_Col = 9
    iGrid_QuantPedida_Col = 10
    iGrid_QuantRecebida_Col = 11
    iGrid_Almoxarifado_Col = 12
    iGrid_CclItemRC_Col = 13
    iGrid_FornecedorItemRC_Col = 14
    iGrid_FilialFornItemRC_Col = 15
    iGrid_ExclusivoItemRC_Col = 16
    iGrid_ObservacaoItemRC_Col = 17

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItensRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensRequisicoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ItensRequisicoes:

    Inicializa_Grid_ItensRequisicoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161103)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Produtos1(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos1

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos1

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Unid. Med.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoProduto.Name)
    objGridInt.colCampo.Add (Produto1.Name)
    objGridInt.colCampo.Add (DescProduto1.Name)
    objGridInt.colCampo.Add (UnidadeMed1.Name)
    objGridInt.colCampo.Add (Quantidade1.Name)
    objGridInt.colCampo.Add (QuantUrgente.Name)
    objGridInt.colCampo.Add (Fornecedor1.Name)
    objGridInt.colCampo.Add (FilialForn1.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoProduto_Col = 1
    iGrid_Produto1_Col = 2
    iGrid_DescProduto1_Col = 3
    iGrid_UnidadeMed1_Col = 4
    iGrid_Quantidade1_Col = 5
    iGrid_QuantUrgente_Col = 6
    iGrid_Fornecedor1_Col = 7
    iGrid_FilialForn1_Col = 8

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos1

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_GERACAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos1 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos1:

    Inicializa_Grid_Produtos1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161104)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Produtos2(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos2

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos2

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Unid. Med.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Tipo Destino")
    objGridInt.colColuna.Add ("Destino")
    objGridInt.colColuna.Add ("Filial Destino")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edição do grid
    objGridInt.colCampo.Add (Produto2.Name)
    objGridInt.colCampo.Add (DescProduto2.Name)
    objGridInt.colCampo.Add (UnidadeMed2.Name)
    objGridInt.colCampo.Add (Quantidade2.Name)
    objGridInt.colCampo.Add (TipoDestinoProd.Name)
    objGridInt.colCampo.Add (Destino.Name)
    objGridInt.colCampo.Add (FilialDestino.Name)
    objGridInt.colCampo.Add (Fornecedor2.Name)
    objGridInt.colCampo.Add (FilialForn2.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Produto2_Col = 1
    iGrid_DescProduto2_Col = 2
    iGrid_UnidadeMed2_Col = 3
    iGrid_Quantidade2_Col = 4
    iGrid_TipoDestino_Col = 5
    iGrid_Destino_Col = 6
    iGrid_FilialDestino_Col = 7
    iGrid_Fornecedor2_Col = 8
    iGrid_FilialForn2_Col = 9

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos2

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_GERACAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos2 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos2:

    Inicializa_Grid_Produtos2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161105)

    End Select

    Exit Function

End Function

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
        End If
    Next

        Carrega_ComboFiliais = SUCESSO

    Exit Function

Erro_Carrega_ComboFiliais:

    Carrega_ComboFiliais = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161106)

    End Select

    Exit Function

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161107)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161108)

    End Select

    Exit Sub

End Sub

Private Sub OrdenacaoReq_GotFocus()
    gsOrdenacao = OrdenacaoReq.Text
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim iIndice As Integer
Dim colRequisicoes As New Collection

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

        'Se foi clicado no TAB_Selecao e o Grid de Cotações já está preenchido
        If TabStrip1.SelectedItem.Index = TAB_Selecao And objGridCotacao.iLinhasExistentes > 0 Then
            iFrameSelecaoAlterado = 0
        End If

        'Se o frame anterior foi o de Seleção e ele foi alterado
        If iFrameAtual <> TAB_Selecao And iFrameSelecaoAlterado = REGISTRO_ALTERADO Then

            Set gobjGeracaoPedCompraCot = New ClassGeracaoPedCompraCot
            Set gcolRequisicaoCompra = New Collection

            'Limpa a seleção atual
            Call Grid_Limpa(objGridCotacao)
            Call Grid_Limpa(objGridRequisicoes)
            Call Grid_Limpa(objGridItensRequisicoes)
            Call Grid_Limpa(objGridProdutos1)
            Call Grid_Limpa(objGridProdutos2)
            Call Grid_Limpa(objGridCotacoes)

            'Preenche Grid de Pedidos de Cotação
            lErro = Preenche_Cotacoes()
            If lErro <> SUCESSO Then gError 67137

            iFrameSelecaoAlterado = 0


        End If

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 67137, 67138, 67151, 67159, 67162, 67160, 70429

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161109)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Function Preenche_Cotacoes() As Long

Dim lErro As Long

On Error GoTo Erro_Preenche_Cotacoes

    'Recolhe os dados do TAB_Selecao
    lErro = Move_TabSelecao_Memoria()
    If lErro <> SUCESSO Then gError 67068

    Set gobjGeracaoPedCompraCot.colCotacao = New Collection

    'Pesquisa no BD os Pedidos de Cotação com as características passadas em gobjGeracaoPedCompra
    lErro = CF("GeracaoPedCompraCot_Le_GeracaoPedCotacao", gobjGeracaoPedCompraCot)
    If lErro <> SUCESSO Then gError 67069

    'Traz os Pedidos de Cotação para a tela
    lErro = Traz_PedCotacao_Tela()
    If lErro <> SUCESSO Then gError 67070

    Preenche_Cotacoes = SUCESSO

    Exit Function

Erro_Preenche_Cotacoes:

    Preenche_Cotacoes = gErr

    Select Case gErr

        Case 67068, 67070

        Case 67069 'Nenhum Ped. Cotação encontrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161110)

    End Select

    Exit Function

End Function

Function Move_TabSelecao_Memoria() As Long
'Recolhe dados do TAB de Seleção

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_TabSelecao_Memoria

    'DataDe
    gobjGeracaoPedCompraCot.dtDataDe = MaskedParaDate(DataDe)

    'DataAte
    gobjGeracaoPedCompraCot.dtDataAte = MaskedParaDate(DataAte)

    'Local de Entrega
    gobjGeracaoPedCompraCot.iSelecionaDestino = SelecionaDestino.Value

    If SelecionaDestino.Value = vbChecked Then

        'Tipo de Destino
        If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True Then
            gobjGeracaoPedCompraCot.iTipoDestino = TIPO_DESTINO_EMPRESA

            'Filial Empresa Destino
            gobjGeracaoPedCompraCot.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)

        Else

            gobjGeracaoPedCompraCot.iTipoDestino = TIPO_DESTINO_FORNECEDOR

            'Se o Fornecedor foi preenchido
            If Len(Trim(Fornecedor.Text)) > 0 Then

                'Fornecedor e Filial Destino
                objFornecedor.sNomeReduzido = Fornecedor.Text

                'Lê o código do Fornecedor
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 67180
                If lErro = 6681 Then gError 67181

                gobjGeracaoPedCompraCot.lFornCliDestino = objFornecedor.lCodigo
                gobjGeracaoPedCompraCot.iFilialDestino = Codigo_Extrai(FilialFornec.Text)

            End If

        End If

    End If

    'Código De
    gobjGeracaoPedCompraCot.lCodigoDe = StrParaLong(CodigoDe.Text)

    'Código Até
    gobjGeracaoPedCompraCot.lCodigoAte = StrParaLong(CodigoAte.Text)

    'Ordenação
    Select Case OrdenacaoPedCot.ListIndex

        Case 0
            gobjGeracaoPedCompraCot.sOrdenacaoCot = " Concorrencia.Codigo"

        Case 1
            gobjGeracaoPedCompraCot.sOrdenacaoCot = " Concorrencia.Data,  Concorrencia.Codigo"

    End Select

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case 67180

        Case 67181
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161111)

    End Select

    Exit Function

End Function

Private Function Traz_PedCotacao_Tela() As Long
'Preenche o Grid de Concorrências

Dim iIndice As Integer
Dim lErro As Long
Dim lCodPedSelecionado As Long
Dim colCampos As New Collection
Dim colCotOrdenada As New Collection

On Error GoTo Erro_Traz_PedCotacao_Tela

    Call Grid_Limpa(objGridCotacao)

    'Guarda o Pedido de Cotação selecionado
    If gobjGeracaoPedCompraCot.iCotacaoSel <> 0 Then
        lCodPedSelecionado = gobjGeracaoPedCompraCot.colCotacao(gobjGeracaoPedCompraCot.iCotacaoSel).lCodigo
    End If

    'Preenche colCampos para a ordenação
    Select Case OrdenacaoPedCot.ListIndex

        Case 0
            colCampos.Add "lCodigo"

        Case 1
            colCampos.Add "dtData"
            colCampos.Add "lCodigo"

    End Select

    'Ordena coleção
    Call Ordena_Colecao(gobjGeracaoPedCompraCot.colCotacao, colCotOrdenada, colCampos)

    'Guarda a nova coleção ordenada de Pedido de Cotações
    Set gobjGeracaoPedCompraCot.colCotacao = colCotOrdenada

    
    'Preenche Grid de Pedidos de Cotações
    For iIndice = 1 To gobjGeracaoPedCompraCot.colCotacao.Count
        GridPedCotacao.TextMatrix(iIndice, iGrid_CodigoPed_Col) = gobjGeracaoPedCompraCot.colCotacao(iIndice).lCodigo
        GridPedCotacao.TextMatrix(iIndice, iGrid_DataPed_Col) = Format(gobjGeracaoPedCompraCot.colCotacao(iIndice).dtData, "dd/mm/yyyy")
        GridPedCotacao.TextMatrix(iIndice, iGrid_DescricaoCot_Col) = gobjGeracaoPedCompraCot.colCotacao(iIndice).sDescricao
    Next

    objGridCotacao.iLinhasExistentes = gobjGeracaoPedCompraCot.colCotacao.Count

    'Seleciona novamente o Pedido de Cotação
    If lCodPedSelecionado <> 0 Then
        For iIndice = 1 To objGridCotacao.iLinhasExistentes
            If lCodPedSelecionado = CLng(GridPedCotacao.TextMatrix(iIndice, iGrid_CodigoPed_Col)) Then
                GridPedCotacao.TextMatrix(iIndice, iGrid_SelecionadoPed_Col) = MARCADO
                gobjGeracaoPedCompraCot.iCotacaoSel = iIndice
                Exit For
            End If
        Next
    End If

    Call Grid_Refresh_Checkbox(objGridCotacao)

    Exit Function

    Traz_PedCotacao_Tela = SUCESSO

Erro_Traz_PedCotacao_Tela:

    Traz_PedCotacao_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161112)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Pergunta se deseja salvar alterações
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 67066

    'Limpa a tela
    Call Limpa_Tela_GeracaoPedCompraCot

    Set gobjGeracaoPedCompraCot = New ClassGeracaoPedCompraCot

    'Preenche novamente o Grid de Pedidos de cotação
    lErro = Preenche_Cotacoes()
    If lErro <> SUCESSO Then gError 67067

    iCotacaoAlterada = 0
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 67066, 67067
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161113)

    End Select

    Exit Sub

End Sub

Private Sub BotaoReqCompras_Click()

Dim lErro As Long
Dim objRequisicaoCompra As New ClassRequisicaoCompras

On Error GoTo Erro_BotaoReqCompras_Click

    If gcolRequisicaoCompra.Count = 0 Then Exit Sub

    'Se nennhuma linha foi selecionada, Erro
    If GridRequisicoes.Row = 0 Then gError 66998

    objRequisicaoCompra.lCodigo = CLng(GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_CodigoReq_Col))
    objRequisicaoCompra.iFilialEmpresa = Codigo_Extrai(GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_FilialReq_Col))

    Call Chama_Tela("ReqComprasCons", objRequisicaoCompra)

    Exit Sub

Erro_BotaoReqCompras_Click:

    Select Case gErr

        Case 66998
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161114)

    End Select

    Exit Sub

End Sub

Private Sub CodigoAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodigoAte_Validate

    'Se o código final não foi preenchido, sai da rotina
    If Len(Trim(CodigoAte.Text)) = 0 Then Exit Sub

    'Se o código inicial for maior que o final, erro
    If StrParaLong(CodigoDe.Text) > StrParaLong(CodigoAte.Text) And Len(Trim(CodigoDe.Text)) > 0 Then gError 67071

    Exit Sub

Erro_CodigoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 67071
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161115)

    End Select

    Exit Sub

End Sub

Private Sub CodigoDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodigoDe_Validate

    'Se o código inicial não foi preenchido, sai da rotina
    If Len(Trim(CodigoDe.ClipText)) = 0 Then Exit Sub

    'Se o código inicial for maior que o final, erro
    If StrParaLong(CodigoDe.Text) > StrParaLong(CodigoAte.Text) And Len(Trim(CodigoAte.Text)) > 0 Then gError 67073

    Exit Sub

Erro_CodigoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 67073
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161116)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub CodigoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoDe, iAlterado)

End Sub

Private Sub CodigoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoAte, iAlterado)

End Sub

Private Sub FilialEmpresa_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub FilialFornec_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornec_Click()

    iAlterado = REGISTRO_ALTERADO

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

Private Sub BotaoPedCotacao_Click()

Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPedCotacao_Click

    'Se nenhuma linha foi selecionada, sai da rotina
    If GridCotacoes.Row = 0 Then gError 89444

    objPedidoCotacao.lCodigo = StrParaLong(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_PedidoCot_Col))
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("PedidoCotacaoCons", objPedidoCotacao)

    Exit Sub

Erro_BotaoPedCotacao_Click:

    Select Case gErr
    
        Case 89444
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161117)
            
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoPedCotacao_evSelecao(obj1 As Object)

    Me.Show

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 67074

    'Se a data inicial for maior que a final erro
    If Len(Trim(DataAte.ClipText)) > 0 And DataDe.Text > DataAte.Text Then gError 67075

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 67074
            'Erro tratado na rotina chamada

        Case 67075
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161118)

    End Select

    Exit Sub

End Sub

Private Sub TabProdutos_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabProdutos.SelectedItem.Index <> iFrameProdutoAtual Then

        If TabStrip_PodeTrocarTab(iFrameProdutoAtual, TabProdutos, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameProdutos(TabProdutos.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameProdutos(iFrameProdutoAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameProdutoAtual = TabProdutos.SelectedItem.Index

    End If

End Sub


Private Sub TipoDestino_Click(Index As Integer)

    'Se o TipoDestino for o mesmo já selecionado, sai da rotina
    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna invisivel o FrameDestino com índice igual a iFrameDestinoAtual
    FrameTipoDestino(iFrameTipoDestinoAtual).Visible = False

    'Torna visível o FrameDestino com índice igual a Index
    FrameTipoDestino(Index).Visible = True

    'Armazena novo valor de giFrameDestinoAtual
    iFrameTipoDestinoAtual = Index

    iFrameSelecaoAlterado = REGISTRO_ALTERADO


End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 67076

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 67076
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161119)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 67077

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 67077
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161120)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataAte está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataAte informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 67078

    'Se a data inicial for maior que a final erro
    If Len(Trim(DataDe.Text)) > 0 And StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 67079

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 67078
            'Erro tratado na rotina chamada

        Case 67079
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161121)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataAte
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 67080

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 67080
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161122)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta um dia em DataAte
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 67081

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 67081
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161123)

    End Select

    Exit Sub

End Sub

Private Sub SelecionaDestino_Click()

Dim iIndice As Integer
Dim bCancel As Boolean

    'Verifica se SelecionaDestino estiver desmarcado
    If SelecionaDestino.Value = vbUnchecked Then

        'Desabilita todos os TipoDestino
        TipoDestino(TIPO_DESTINO_EMPRESA).Enabled = False
        TipoDestino(TIPO_DESTINO_FORNECEDOR).Enabled = False
        FornecedorLabel.Enabled = False
        FilEmprDestLabel.Enabled = False
        FilFornDestLabel.Enabled = False

        'Limpa os campos do Frame Destino()
        FilialEmpresa.Text = ""
        Fornecedor.Text = ""
        FilialFornec.ListIndex = -1

    'Verifica se SelecionaDestino está marcado
    ElseIf SelecionaDestino.Value = vbChecked Then

        'Haabilita todos os TipoDestino
        TipoDestino(TIPO_DESTINO_EMPRESA).Enabled = True
        TipoDestino(TIPO_DESTINO_FORNECEDOR).Enabled = True
        FornecedorLabel.Enabled = True
        FilEmprDestLabel.Enabled = True
        FilFornDestLabel.Enabled = True

        Fornecedor.Enabled = True
        FilialFornec.Enabled = True
        FilialEmpresa.Enabled = True

        'Se nenhuma FilialEmpresa estiver selecionada
        If FilialEmpresa.ListIndex = -1 Then FilialEmpresa.Text = giFilialEmpresa
        Call FilialEmpresa_Validate(bCancel)

    End If

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Exit Sub

End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresa_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialEmpresa.Text)) = 0 Then Exit Sub

    'Verifica se é uma FilialEmpresa selecionada
    If FilialEmpresa.Text = FilialEmpresa.List(FilialEmpresa.ListIndex) Then Exit Sub

    'Tenta selecionar a FilialEmpresa na combo FilialEmpresa
    lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 67082

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'preeenche objFilialEmpresa
        objFilialEmpresa.iCodFilial = iCodigo

        'Le a FilialEmpresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 67083

        'Se nao encontrou => erro
        If lErro = 27378 Then gError 67084

        If lErro = SUCESSO Then

            'Coloca na tela o codigo e o nome da FilialEmpresa
            FilialEmpresa.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

        End If

    End If

    'Se nao encontrou e nao era codigo
    If lErro = 6731 Then gError 67085

    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True

    Select Case gErr

        Case 67082, 67083
            'Erros tratados nas rotinas chamadas

        Case 67084
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 67085
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA1", gErr, FilialEmpresa.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161124)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se Fornecedor foi alterado
    If iFornecedorAlterado = 0 Then Exit Sub

    'Verifica se o Fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Le o Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 67086

        'Le as Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then gError 67087

        'Preenche a combo FilialFornec
        Call CF("Filial_Preenche", FilialFornec, colCodigoNome)

        'Seleciona a filial na combo de FilialFornec
        Call CF("Filial_Seleciona", FilialFornec, iCodFilial)

    End If

    'Se o Fornecedor nao estiver preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a combo FilialForn
        FilialFornec.Clear

    End If

    iFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 67086, 67087
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161125)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornec_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FilialFornec_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornec.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornec.ListIndex >= 0 Then Exit Sub

    'Tenta selecionar na combo de FilialFornec
    lErro = Combo_Seleciona(FilialFornec, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 67088

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 67092

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 67089

        'Se nao existir
        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = sFornecedor

            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 67090

            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            gError 67091

        End If

        'Coloca na tela o código e o nome da FilialForn
        FilialFornec.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 67093

    Exit Sub

Erro_FilialFornec_Validate:

    Cancel = True

    Select Case gErr

        Case 67088, 67089, 67090 'Tratados nas Rotinas chamadas

        Case 67091
            'Pergunta se deseja criar nova filial para o fornecedor em questao
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela FiliaisFornecedores
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 67092
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 67093
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornec.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161126)

    End Select

    Exit Sub

End Sub

Private Sub OrdenacaoPedCot_Click()

Dim lErro As Long

On Error GoTo Erro_OrdenacaoPedCot_Click

    'Se não existem Pedidos de cotação, sai da rotina
    If objGridCotacao.iLinhasExistentes = 0 Then Exit Sub

    'Preenche o Grid de Pedido de Cotações
    lErro = Traz_PedCotacao_Tela()
    If lErro <> SUCESSO Then gError 62801
    
    Exit Sub

Erro_OrdenacaoPedCot_Click:

    Select Case gErr
    
        Case 62801
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161127)
            
    End Select

    Exit Sub

End Sub

Private Sub ObservacaoReq_GotFocus()
    gsOrdenacao = OrdenacaoReq.Text
End Sub

Private Sub OrdenacaoCot_GotFocus()
    gsOrdenacao = OrdenacaoCot.Text
End Sub
Private Sub OrdenacaoReq_Click()

Dim lErro As Long
Dim colReqCompraSaida As New Collection
Dim colCampos As New Collection

On Error GoTo Erro_OrdenacaoReq_Click

    If gsOrdenacao = "" Then Exit Sub

    'Verifica se OrdenacaoReq da tela é diferente de gsOrdenacao
    If OrdenacaoReq.Text <> gsOrdenacao Then

        Call Monta_Colecao_Campos_Requisicao(colCampos, OrdenacaoReq.ListIndex)
        'Ordena
        lErro = Ordena_Colecao(gcolRequisicaoCompra, colReqCompraSaida, colCampos)
        If lErro <> SUCESSO Then gError 63908

        Set gcolRequisicaoCompra = colReqCompraSaida

    End If

    'COloca as Requsiicoes na tela ordenadamente
    lErro = GridRequisicoes_Preenche()
    If lErro <> SUCESSO Then gError 62750
    
    'Coloca os itens na tela de acordo com a ordem das requisições.
    lErro = GridItensReq_Preenche()
    If lErro <> SUCESSO Then gError 62751

    Exit Sub

Erro_OrdenacaoReq_Click:

    Select Case gErr

        Case 62750, 62751, 63907 To 63909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161128)

    End Select

    Exit Sub

End Sub

Sub Monta_Colecao_Campos_Requisicao(colCampos As Collection, iOrdenacao As Integer)

    Select Case iOrdenacao

        Case 0

            colCampos.Add "iFilialEmpresa"
            colCampos.Add "lCodigo"

        Case 1

            colCampos.Add "dtDataLimite"
            colCampos.Add "iFilialEmpresa"
            colCampos.Add "lCodigo"

        Case 2

            colCampos.Add "dtData"
            colCampos.Add "iFilialEmpresa"
            colCampos.Add "lCodigo"

    End Select
    
    Exit Sub

End Sub

Private Function Traz_Requisicoes_Tela(gobjGeracaoPedCompraCot As ClassGeracaoPedCompraCot) As Long
'Preenche as Requisições a partir da coleção passada

Dim lErro As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim iIndice As Integer
Dim objCotacao As New ClassCotacao

On Error GoTo Erro_Traz_Requisicoes_Tela

    'Preenche o grid de requisiçõe scom dados das requisições
    lErro = GridRequisicoes_Preenche()
    If lErro <> SUCESSO Then gError 62745

    'Para cada item de concorrência gerado
    For iIndice = 1 To gcolItemConcorrencia.Count
        'Busca as cotacoes para o item de concorrência
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iIndice)
        If lErro <> SUCESSO Then gError 62749
    Next

    'Preenche o grid de itens de Requisição
    lErro = GridItensReq_Preenche()
    If lErro <> SUCESSO Then gError 62746
    
    Traz_Requisicoes_Tela = SUCESSO

    Exit Function

Erro_Traz_Requisicoes_Tela:

    Traz_Requisicoes_Tela = gErr

    Select Case gErr

        Case 62745, 62746, 62749
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161129)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcarTodosReq_Click()
'Marca todas CheckBox do GridRequisicoes

Dim lErro As Long
Dim iItem As Integer
Dim iLinha As Integer
Dim iIndice As Integer
Dim colItens As New Collection
Dim objItemRC As ClassItemReqCompras
Dim objItemConc As New ClassItemConcorrencia
Dim objReqCompras As ClassRequisicaoCompras

On Error GoTo Erro_BotaoMarcarTodosReq_Click
    
    If gcolRequisicaoCompra.Count = 0 Then Exit Sub
    
    Set gcolItemConcorrencia = New Collection
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridRequisicoes.iLinhasExistentes
        
        'Marca na tela a linha em questão
        GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_ATIVO
        Set objReqCompras = gcolRequisicaoCompra(iLinha)
        objReqCompras.iSelecionado = MARCADO
        
        'Para cada Item
        For Each objItemRC In objReqCompras.colItens
            'Seleciona o item
            objItemRC.iSelecionado = True

            lErro = ItensConcorrencia_Cria_Altera(objItemRC)
            If lErro <> SUCESSO Then gError 62752
        
            Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConc, iItem, objItemRC)
            
            Call Adiciona_Codigo(colItens, iItem)
        
        Next

    Next
    
    'ATualiza as cotações
    For iIndice = 1 To colItens.Count
       lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, colItens(iIndice))
       If lErro <> SUCESSO Then gError 62767
    Next
    
    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridRequisicoes)
    
    'Preenche o grid de itens
    lErro = GridItensReq_Preenche()
    If lErro <> SUCESSO Then gError 62753
    
    'Preenche o grid de Produtos
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62754
    
    Exit Sub
    
Erro_BotaoMarcarTodosReq_Click:

    Select Case gErr
    
        Case 62752, 62753, 62754, 62767
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161130)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodosItensRC_Click()
'Desmarca todas CheckBox do GridItensRequisicoes

Dim iIndice As Integer
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras

    If gcolRequisicaoCompra.Count = 0 Then Exit Sub
    
    'Desmarca na coleção todos os itens
    For Each objReqCompras In gcolRequisicaoCompra
        For Each objItemRC In objReqCompras.colItens
            objItemRC.iSelecionado = DESMARCADO
        Next
    Next
    
    'Desmarca no grid todos os itens
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes
        GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col) = DESMARCADO
    Next

    'Limpa a coleção de itens de concorrência
    Set gcolItemConcorrencia = New Collection
    
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)
    
    Call Grid_Limpa(objGridProdutos1)
    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)
    
    Call Calcula_TotalItens
    
    Exit Sub

End Sub
Private Sub BotaoDesmarcarTodosReq_Click()
'Desmarca todas CheckBox do GridRequisicoes

Dim iLinha As Integer

    If gcolRequisicaoCompra.Count = 0 Then Exit Sub

    Set gcolItemConcorrencia = New Collection
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridRequisicoes.iLinhasExistentes
    
        'Desmarca na tela a linha em questão
        GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_INATIVO
        gcolRequisicaoCompra(iLinha).iSelecionado = DESMARCADO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridRequisicoes)
    
    Call Grid_Limpa(objGridItensRequisicoes)
    Call Grid_Limpa(objGridProdutos1)
    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)

    Call Calcula_TotalItens

    Exit Sub

End Sub
Private Sub BotaoDesmarcarTodosProd_Click()
'Desmarca todas CheckBox do GridProdutos1
Dim iIndice As Integer

    'Marca todos os Itens do GridProdutos1
    For iIndice = 1 To objGridProdutos1.iLinhasExistentes
        GridProdutos1.TextMatrix(iIndice, iGrid_EscolhidoProduto_Col) = DESMARCADO
        gcolItemConcorrencia(iIndice).iEscolhido = DESMARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridProdutos1)

    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)
    
    Call Calcula_TotalItens
    
    Exit Sub

End Sub
Private Sub BotaoMarcarTodosProd_Click()
'Marca todas CheckBox do GridProdutos1

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoMarcarTodosProd_Click

    'Marca todos os Itens do GridProdutos1
    For iIndice = 1 To objGridProdutos1.iLinhasExistentes
        GridProdutos1.TextMatrix(iIndice, iGrid_EscolhidoProduto_Col) = GRID_CHECKBOX_ATIVO
        gcolItemConcorrencia(iIndice).iEscolhido = MARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridProdutos1)
    
    'Preenche o grid de produtos
    lErro = GridProdutos2_Preenche()
    If lErro <> SUCESSO Then gError 62759

    'Preenche o grid de cotações
    lErro = GridCotacoes_Preenche()
    If lErro <> SUCESSO Then gError 62760

    Exit Sub

Erro_BotaoMarcarTodosProd_Click:

    Select Case gErr

        Case 62759, 62760

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161131)

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

            'Se for o Grid Pedido de Cotação
            Case GridPedCotacao.Name

                lErro = Saida_Celula_GridPedCotacao(objGridInt)
                If lErro <> SUCESSO Then gError 67100

            'se for o GridRequisicoes
            Case GridRequisicoes.Name

                lErro = Saida_Celula_GridRequisicoes(objGridInt)
                If lErro <> SUCESSO Then gError 67101

            'se for o GridItensReq
            Case GridItensRequisicoes.Name

                lErro = Saida_Celula_GridItensReq(objGridInt)
                If lErro <> SUCESSO Then gError 67102

            'se for o GridProdutos1
            Case GridProdutos1.Name

                lErro = Saida_Celula_GridProdutos1(objGridInt)
                If lErro <> SUCESSO Then gError 67103


            'se for o GridProdutos2
            Case GridProdutos2.Name

                lErro = Saida_Celula_GridProdutos2(objGridInt)
                If lErro <> SUCESSO Then gError 67104

            'se for o GridCotacoes
            Case GridCotacoes.Name

                lErro = Saida_Celula_GridCotacoes(objGridInt)
                If lErro <> SUCESSO Then gError 67105

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 67106

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 67100, 67101, 67102, 67106, 67103, 67104, 67105
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161132)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridPedCotacao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridPedCotacao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Selecionado
        Case iGrid_SelecionadoPed_Col
            lErro = Saida_Celula_SelecionadoPed(objGridInt)
            If lErro <> SUCESSO Then gError 67106

    End Select

    Saida_Celula_GridPedCotacao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridPedCotacao:

    Saida_Celula_GridPedCotacao = gErr

    Select Case gErr

        Case 67106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161133)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_SelecionadoPed(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SelecionadoPed

    Set objGridInt.objControle = Selecionado

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67107

    Saida_Celula_SelecionadoPed = SUCESSO

    Exit Function

Erro_Saida_Celula_SelecionadoPed:

    Saida_Celula_SelecionadoPed = gErr

    Select Case gErr

        Case 67107
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161134)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridRequisicoes(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridRequisicoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoReq
        Case iGrid_EscolhidoReq_Col
            lErro = Saida_Celula_EscolhidoReq(objGridInt)
            If lErro <> SUCESSO Then gError 67108

    End Select

    Saida_Celula_GridRequisicoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridRequisicoes:

    Saida_Celula_GridRequisicoes = gErr

    Select Case gErr

        Case 67108

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161135)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoReq(objGridInt As AdmGrid) As Long
'Faz a saida de célula de EscolhidoReq

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoReq

    Set objGridInt.objControle = EscolhidoReq

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63962

    Exit Function

Erro_Saida_Celula_EscolhidoReq:

    Select Case gErr

        Case 63962
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161136)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItensReq(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensReq

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoItem
        Case iGrid_EscolhidoItem_Col
            lErro = Saida_Celula_EscolhidoItem(objGridInt)
            If lErro <> SUCESSO Then gError 67110

        'QuantComprarItemReq
        Case iGrid_QuantComprarItem_Col
            lErro = Saida_Celula_QuantComprarItemReq(objGridInt)
            If lErro <> SUCESSO Then gError 67111

    End Select

    Saida_Celula_GridItensReq = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensReq:

    Saida_Celula_GridItensReq = gErr

    Select Case gErr

        Case 67110, 67111

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161137)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoItem

    Set objGridInt.objControle = EscolhidoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67112

    Saida_Celula_EscolhidoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_EscolhidoItem:

    Saida_Celula_EscolhidoItem = gErr

    Select Case gErr

        Case 67112
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161138)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantComprarItemReq(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim dQuantPosterior As Double
Dim dQuantAnterior As Double
Dim iIndice1 As Integer, iItem As Integer
Dim iIndice2 As Integer
Dim bAchou As Boolean, objProduto As New ClassProduto
Dim dQuantDiferenca As Double, dFator As Double
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_Saida_Celula_QuantComprarItemReq

    Set objGridInt.objControle = QuantComprarItemRC
    
    'Guarda a quantidade anterior do grid
    dQuantAnterior = StrParaDbl(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_QuantComprarItem_Col))

    'Se quantidade estiver preenchida
    If Len(Trim(QuantComprarItemRC.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(QuantComprarItemRC.Text)
        If lErro <> SUCESSO Then gError 63964
        
        'Guarda a qt alterada
        dQuantPosterior = StrParaDbl(QuantComprarItemRC.Text)

    Else
        gError 62799
    End If
    
    'Calula a diferença entre a quant anterior e a atual
    dQuantDiferenca = Round(dQuantPosterior - dQuantAnterior, 2)
        
    'Se houve alteração na quantidade
    If dQuantDiferenca <> 0 Then
        
        'Localiza o item e a requisição da linha selecionada
        For iIndice1 = 1 To gcolRequisicaoCompra.Count
            Set objReqCompra = gcolRequisicaoCompra(iIndice1)
            
            For iIndice2 = 1 To objReqCompra.colItens.Count
                
                Set objItemRC = objReqCompra.colItens(iIndice2)
                
                If objItemRC.iItem = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_Item_Col)) And _
                   objReqCompra.lCodigo = StrParaLong(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_CodigoReqItem_Col)) And _
                   objReqCompra.iFilialEmpresa = Codigo_Extrai(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_FilialReqItem_Col)) Then
                    'Achou
                    bAchou = True
                    Exit For
                End If
            Next
            'Se já achou --> sai
            If bAchou Then Exit For
        Next
        
        
        'Verifica se a quantidade digitada é maior que a quant que falta comprar do itemrc
        If dQuantPosterior > objItemRC.dQuantidade - objItemRC.dQuantCancelada - objItemRC.dQuantPedida - objItemRC.dQuantRecebida Then gError 63965
        
        'Localiza o ItemConcorrência vinculado ao Item RC
        Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConcorrencia, iItem, objItemRC)
        
        objProduto.sCodigo = objItemConcorrencia.sProduto
        
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 23080 Then gError 62756
        If lErro <> SUCESSO Then gError 62757
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 62758
        
        'Converte a quantidade p\ UM de compra
        dQuantDiferenca = dQuantDiferenca * dFator
        
        objItemRC.dQuantComprar = dQuantPosterior
        objItemRC.dQuantNaConcorrencia = objItemRC.dQuantComprar * dFator
                
        'Se a quantidade foi aumentada
        If dQuantDiferenca > 0 Then
            'Aumenta a quantidade do item de concorrência
            lErro = ItemConcorrencia_Inclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC, dQuantDiferenca)
            If lErro <> SUCESSO Then gError 62759
            
        'Se a quantidade foi diminuida
        ElseIf iItem > 0 Then
        
            'Diminui a quantidade no item de concorrência
            lErro = ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC, Abs(dQuantDiferenca))
            If lErro <> SUCESSO Then gError 62760
            
        End If
        
        'Atualiza as cotações
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iItem)
        If lErro <> SUCESSO Then gError 62771
            
        'Preenche o grid de produtos
        lErro = Grids_Produto_Preenche()
        If lErro <> SUCESSO Then gError 62761
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63966
    
    Saida_Celula_QuantComprarItemReq = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprarItemReq:

    Saida_Celula_QuantComprarItemReq = gErr

    Select Case gErr

        Case 62756, 63964, 63966, 62758, 62759, 62760, 62761, 62771
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 62757
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 62799
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 63965
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_SUPERIOR_MAXIMA", gErr, dQuantPosterior, objItemRC.dQuantidade - objItemRC.dQuantCancelada - objItemRC.dQuantPedida - objItemRC.dQuantRecebida)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161139)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridProdutos1(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProdutos1

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoProduto
        Case iGrid_EscolhidoProduto_Col
            lErro = Saida_Celula_EscolhidoProduto(objGridInt)
            If lErro <> SUCESSO Then gError 67116

    End Select

    Saida_Celula_GridProdutos1 = SUCESSO

    Exit Function

Erro_Saida_Celula_GridProdutos1:

    Saida_Celula_GridProdutos1 = gErr

    Select Case gErr

        Case 67116

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161140)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoProduto

    Set objGridInt.objControle = EscolhidoProduto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67117

    Saida_Celula_EscolhidoProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_EscolhidoProduto:

    Saida_Celula_EscolhidoProduto = gErr

    Select Case gErr

        Case 67117
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161141)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridProdutos2(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProdutos2

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Quantidade2
        Case iGrid_Quantidade2_Col
            lErro = Saida_Celula_Quantidade2(objGridInt)
            If lErro <> SUCESSO Then gError 67118

    End Select

    Saida_Celula_GridProdutos2 = SUCESSO

    Exit Function

Erro_Saida_Celula_GridProdutos2:

    Saida_Celula_GridProdutos2 = gErr

    Select Case gErr

        Case 67118

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161142)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade2(objGridInt As AdmGrid) As Long

Dim lErro As Long, dQuantidade As Double
Dim iIndice As Integer, dQuantTotalRC As Double
Dim sFornecedor As String, iFilial As Integer
Dim sProduto As String, dQuantAnterior As Double
Dim dQuantDiferenca As Double, iItem As Integer

On Error GoTo Erro_Saida_Celula_Quantidade2

    Set objGridInt.objControle = Quantidade2
    
    dQuantAnterior = StrParaDbl(GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Quantidade2_Col))

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade2.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade2.Text)
        If lErro <> SUCESSO Then gError 63963

        dQuantidade = CDbl(Quantidade2.Text)

        'Coloca o valor Formatado na tela
        Quantidade2.Text = Formata_Estoque(dQuantidade)
    Else
        gError 62744
    End If

    'Calcula a diferença entre a quant anterior e essa
    dQuantDiferenca = StrParaDbl(Formata_Estoque(dQuantidade - dQuantAnterior))
    
    'Guarda campos da linha em questão de GridProdutos2
    sProduto = GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col)
    sFornecedor = GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Fornecedor2_Col)
    iFilial = Codigo_Extrai(GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_FilialForn2_Col))

    'Atualiza o valor da coleção de qt suplementares
    ' e verifica se a qt digitada é < que a qt dos itens req
    For iIndice = 1 To objGridProdutos1.iLinhasExistentes
        If sProduto = GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) And sFornecedor = GridProdutos1.TextMatrix(iIndice, iGrid_Fornecedor1_Col) And iFilial = Codigo_Extrai(GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col)) Then
            lErro = Atualiza_QuantSupl(gcolItemConcorrencia(iIndice), dQuantDiferenca, GridProdutos2.Row)
            If lErro <> SUCESSO Then gError 63965
            Exit For
        End If
    Next

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63964

    'Se a quant foi alterada
    If dQuantDiferenca <> 0 Then
    
        'Atualiza a quantidade a comprar no GridProdutos1
        For iIndice = 1 To objGridProdutos1.iLinhasExistentes
            If sProduto = GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) And sFornecedor = GridProdutos1.TextMatrix(iIndice, iGrid_Fornecedor1_Col) And iFilial = Codigo_Extrai(GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col)) Then
                GridProdutos1.TextMatrix(iIndice, iGrid_Quantidade1_Col) = Formata_Estoque(StrParaDbl(GridProdutos1.TextMatrix(iIndice, iGrid_Quantidade1_Col)) + dQuantDiferenca)
                
                'Se a qt foi diminuida
                If dQuantDiferenca < 0 Then
                    'Exclui a quant no item de conc
                    lErro = ItemConcorrencia_Exclui_QuantComprar(gcolItemConcorrencia(iIndice), iIndice, , , Abs(dQuantDiferenca))
                    If lErro <> SUCESSO Then gError 62761
                'Senão
                Else
                    'Inclui a quant no item de conc
                    lErro = ItemConcorrencia_Inclui_QuantComprar(gcolItemConcorrencia(iIndice), iIndice, , , dQuantDiferenca)
                    If lErro <> SUCESSO Then gError 62762
                End If
                
                'Atualiza as cotaçõe spara a nova quantidade
                lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iIndice)
                If lErro <> SUCESSO Then gError 62763
                
                Exit For
            End If
        Next

        'Preenche o grid de Cotações
        lErro = GridCotacoes_Preenche()
        If lErro <> SUCESSO Then gError 62764
    End If
    
    Saida_Celula_Quantidade2 = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade2:

    Saida_Celula_Quantidade2 = gErr

    Select Case gErr

        Case 63963, 63964, 62761, 62762, 62763, 62764, 63965
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 62744
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161143)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridCotacoes(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridCotacoes que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCotacoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoCot
        Case iGrid_EscolhidoCot_Col
            lErro = Saida_Celula_EscolhidoCot(objGridInt)
            If lErro <> SUCESSO Then gError 63945

        'QuantComprarCot
        Case iGrid_QuantComprarCot_Col
            lErro = Saida_Celula_QuantComprarCot(objGridInt)
            If lErro <> SUCESSO Then gError 63946

        'Preço Unitário
        Case iGrid_PrecoUnitarioCot_Col
            lErro = Saida_Celula_PrecoUnitarioCot(objGridInt)
            If lErro <> SUCESSO Then gError 70459

        'MotivoEscolhaCot
        Case iGrid_MotivoEscolhaCot_Col
            lErro = Saida_Celula_MotivoEscolhaCot(objGridInt)
            If lErro <> SUCESSO Then gError 63947

    End Select

    Saida_Celula_GridCotacoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridCotacoes:

    Saida_Celula_GridCotacoes = gErr

    Select Case gErr

        Case 63945, 63946, 63947, 70459

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161144)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoCot(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoCot

    Set objGridInt.objControle = EscolhidoCot

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67124

    Saida_Celula_EscolhidoCot = SUCESSO

    Exit Function

Erro_Saida_Celula_EscolhidoCot:

    Saida_Celula_EscolhidoCot = gErr

    Select Case gErr

        Case 67124
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161145)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitarioCot(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objCotItemConc As New ClassCotacaoItemConc
Dim dValorPresente As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Saida_Celula_PrecoUnitarioCot

    Set objGridInt.objControle = PrecoUnitario

    'Se o Preço unitário estiver preenchido
    If Len(Trim(PrecoUnitario.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 70482

    End If
        
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    
    objCotItemConc.dPrecoAjustado = StrParaDbl(PrecoUnitario.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 70483

    'Se a condição de pagamento não for a vista
    If Codigo_Extrai(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_CondPagtoCot_Col)) <> COD_A_VISTA And PercentParaDbl(TaxaEmpresa.Caption) > 0 Then
        
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConc.sCondPagto)
        
        'Recalcula o Valor Presente
        lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
        If lErro <> SUCESSO Then gError 62736
        
        If objCotItemConc.iMoeda <> MOEDA_REAL Then
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format(dValorPresente * objCotItemConc.dTaxa, ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente * objCotItemConc.dTaxa
        Else
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format(dValorPresente, ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente
        End If
        
    ElseIf Codigo_Extrai(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_CondPagtoCot_Col)) = COD_A_VISTA Then
        
        If objCotItemConc.iMoeda <> MOEDA_REAL Then
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format((StrParaDbl(PrecoUnitario.Text)) * objCotItemConc.dTaxa, ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente * objCotItemConc.dTaxa
        Else
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format((StrParaDbl(PrecoUnitario.Text)), ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente
        End If
        
    End If
    
    If objCotItemConc.iMoeda <> MOEDA_REAL Then
        'Atualiza o valor desse item alterado
        GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar * objCotItemConc.dTaxa, "STANDARD")
    Else
        'Atualiza o valor desse item alterado
        GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar, "STANDARD")
    End If
    
    'Atuliza o valor dos itens selecionados
    Call Calcula_TotalItens
    
    Saida_Celula_PrecoUnitarioCot = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitarioCot:

    Saida_Celula_PrecoUnitarioCot = gErr

    Select Case gErr

        Case 62736, 70482, 70483
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161146)

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
    
    'Localiza o ItemCotacao selecionado
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    
    'Atualiza a quantidade a comprar
    objCotItemConc.dQuantidadeComprar = dQuantidade
    'Atualiza o valor do item
    GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar, "STANDARD")
    
    'recalcula o total
    Call Calcula_TotalItens
    
    Saida_Celula_QuantComprarCot = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprarCot:

    Saida_Celula_QuantComprarCot = gErr

    Select Case gErr

        Case 63739, 63740
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161147)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MotivoEscolhaCot(objGridInt As AdmGrid) As Long
'Faz a saida de celula de MotivoEscolha

Dim lErro As Long
Dim iCodigo As Integer
Dim objCotItemConc As ClassCotacaoItemConc

On Error GoTo Erro_Saida_Celula_MotivoEscolhaCot

    Set objGridInt.objControle = MotivoEscolhaCot

    'Verifica se o MotivoEscolhaCot está preenchido
    If Len(Trim(MotivoEscolhaCot.Text)) > 0 Then

        'Verifica se MotivoEscolhaCot não está selecionado
        If MotivoEscolhaCot.ListIndex = -1 Then
                        
            If UCase(MotivoEscolhaCot.Text) = UCase(MOTIVO_EXCLUSIVO_DESCRICAO) Then gError 62715
            
            'Seleciona o MotivoEscolhaCot na combobox
            lErro = Combo_Item_Seleciona(MotivoEscolhaCot)
            If lErro <> SUCESSO And lErro <> 12250 Then gError 63741

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63743

    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)

    objCotItemConc.sMotivoEscolha = GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_MotivoEscolhaCot_Col)

    Saida_Celula_MotivoEscolhaCot = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoEscolhaCot:

    Saida_Celula_MotivoEscolhaCot = gErr

    Select Case gErr

        Case 62715
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_EXCLUSIVO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63741, 63743
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161148)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_GeracaoPedCompraCot()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Limpa_Tela_GeracaoPedCompraCot

    'Função genérica que limpa a tela
    Call Limpa_Tela(Me)

    'Limpa Frame de Seleção
    DataDe.PromptInclude = False
    DataDe.Text = ""
    DataDe.PromptInclude = True

    DataAte.PromptInclude = False
    DataAte.Text = ""
    DataAte.PromptInclude = True

    SelecionaDestino.Value = vbUnchecked

    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True

    CodigoDe.PromptInclude = False
    CodigoDe.Text = ""
    CodigoDe.PromptInclude = True

    CodigoAte.PromptInclude = False
    CodigoAte.Text = ""
    CodigoAte.PromptInclude = True

    'Limpa os Grids
    Call Grid_Limpa(objGridCotacao)
    Call Grid_Limpa(objGridRequisicoes)
    Call Grid_Limpa(objGridItensRequisicoes)
    Call Grid_Limpa(objGridProdutos1)
    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)

    'Limpa Frame de Cotações
    Concorrencia.Caption = ""

    Set gobjGeracaoPedCompraCot = Nothing
    Set gcolRequisicaoCompra = New Collection
    Set gColCotacoes = New Collection

    Call Calcula_TotalItens
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iCotacaoAlterada = 0

    Exit Sub

Erro_Limpa_Tela_GeracaoPedCompraCot:

    Select Case gErr

        Case 66916

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161149)

    End Select

    Exit Sub

End Sub

Function Gravar_Pedidos() As Long

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim colPedidoCompra As New Collection
Dim objCotacao As ClassCotacao

On Error GoTo Erro_Gravar_Pedidos

    GL_objMDIForm.MousePointer = vbHourglass

    Set objCotacao = gobjGeracaoPedCompraCot.colCotacao(gobjGeracaoPedCompraCot.iCotacaoSel)

    'Recolhe os dados da tela
    lErro = Move_Concorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63920

    'Atualiza a Concorrencia no Banco de Dados
    lErro = CF("Concorrencia_Grava", objConcorrencia)
    If lErro <> SUCESSO Then gError 63921

    'Carrega em colPedidoCompras os Pedidos de Compra gerados a partir de diferentes Fornecedores e FiliaisFornecedores
    lErro = Carrega_Dados_Pedidos(objConcorrencia, colPedidoCompra)
    If lErro <> SUCESSO Then gError 63922

    'Grava o Pedido de Compras
    lErro = CF("PedCompra_Concorrencia_Grava", objConcorrencia, colPedidoCompra)
    If lErro <> SUCESSO Then gError 63923

    '#####################################
    'Inserido por Wagner
    If colPedidoCompra.Count > 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_CODIGO_PEDCOMPRA_GRAVADO", colPedidoCompra.Item(1).lCodigo, colPedidoCompra.Item(colPedidoCompra.Count).lCodigo)
    End If
    '#####################################

    'Limpa a tela
    Call Limpa_Tela_GeracaoPedCompraCot

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Pedidos = SUCESSO

    Exit Function

Erro_Gravar_Pedidos:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Pedidos = gErr

    Select Case gErr

        Case 63919 To 63923

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161150)

    End Select

    Exit Function

End Function

Private Function Move_Concorrencia_Memoria(objConcorrencia As ClassConcorrencia) As Long
'Recolhe os dados da tela e armazena em objConcorrencia

Dim lErro As Long
Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador
Dim objFornecedor As New ClassFornecedor
Dim iLinha As Integer

On Error GoTo Erro_Move_Concorrencia_Memoria
    
    If gcolRequisicaoCompra.Count > 0 Then
    
        'Verifica se o GridRequisicoes está vazio
        If objGridRequisicoes.iLinhasExistentes = 0 Then gError 63924
        
        'Verifica se existe algum Item de Requisicao selecionado
        For iLinha = 1 To objGridItensRequisicoes.iLinhasExistentes
            If GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_ATIVO Then
                Exit For
            End If
        Next
    
        If iLinha > objGridItensRequisicoes.iLinhasExistentes Then gError 63925
    End If
    
    'Verifica se existe algum Item de Requisicao selecionado
    For iLinha = 1 To objGridProdutos1.iLinhasExistentes
        If GridProdutos1.TextMatrix(iLinha, iGrid_EscolhidoProduto_Col) = GRID_CHECKBOX_ATIVO Then
            Exit For
        End If
    Next

    If iLinha > objGridProdutos1.iLinhasExistentes Then gError 63749
    
    If SelecionaDestino.Value = vbChecked Then
        'Verifica o Tipo de Destino selecionado é FilialEmpresa
        If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then
    
            'Verifica se a FilialEmpresa está preenchida
            If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 63746
            
            objConcorrencia.iTipoDestino = TIPO_DESTINO_EMPRESA
            objConcorrencia.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)
    
        'Verifica se o TipoDestino é Fornecedor
        ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then
    
            'Verifica se o Fornecedor está preenchido
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 63747
    
            'Verifica se a Filial do Fornecedor está preenchida
            If Len(Trim(FilialFornec.Text)) = 0 Then gError 63748
    
            objConcorrencia.iTipoDestino = TIPO_DESTINO_FORNECEDOR
            objConcorrencia.iFilialDestino = Codigo_Extrai(FilialFornec.Text)
            
            'Lê o Fornecedor
            objFornecedor.sNomeReduzido = Fornecedor.Text
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 63775
                
            'Se o Fornecedor não estiver cadastrado, Erro
            If lErro = 6681 Then gError 70491
            objConcorrencia.lFornCliDestino = objFornecedor.lCodigo
        End If
    Else
        objConcorrencia.iTipoDestino = TIPO_DESTINO_AUSENTE
    End If
    
    'Verifica se o GridProdutos está vazio
    If objGridProdutos1.iLinhasExistentes = 0 Then gError 63749
    
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

        Case 63924
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_SELECIONADA", gErr)

        Case 63925
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_REQUISICAO_NAO_SELECIONADO", gErr)

        Case 63746
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)

        Case 63749
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEMCONC_SELECIONADO", gErr)

        Case 63747
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO", gErr)
        
        Case 63748
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA", gErr)
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161151)

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
        If objItemConcorrencia.iEscolhido = MARCADO Then
            'verifica se a quantidade foi preenchida
            If objItemConcorrencia.dQuantidade = 0 Then gError 63750
            
            'valida a quantidade do item de concorrência
            lErro = Valida_Quantidade(objItemConcorrencia, iItem)
            If lErro <> SUCESSO Then gError 70492
    
            objConcorrencia.colItens.Add objItemConcorrencia
        End If
    Next

    Move_ItensConcorrencia_Memoria = SUCESSO

    Exit Function

Erro_Move_ItensConcorrencia_Memoria:

    Move_ItensConcorrencia_Memoria = gErr

    Select Case gErr

        Case 63750
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr)

        Case 70492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161152)

    End Select

    Exit Function

End Function

Function Carrega_Dados_Pedidos(objConcorrencia As ClassConcorrencia, colPedidoCompras As Collection) As Long
'Carrega em colPedidoCompras os Pedidos de Compra gerados a partir de diferentes Fornecedores e FiliaisFornecedores

Dim lErro As Long, bAchou As Boolean
Dim iIndice As Integer, objItemPC As ClassItemPedCompra
Dim dTotalItens As Double, lNumIntOriginal As Long
Dim objFornecedor As New ClassFornecedor
Dim objItemCotacao As ClassItemCotacao
Dim objCotItemConc As ClassCotacaoItemConc
Dim colItensCotacao As New Collection
Dim objQuantSupl As New ClassQuantSuplementar
Dim objPedidoCompra As ClassPedidoCompras
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colPedCompraGeral As New Collection
Dim colPedCompraExclu As New Collection
Dim objItemConcorrencia As ClassItemConcorrencia
Dim dQuantSupl As Double
Dim colCotItemConcAux As Collection
Dim colProdutos As New Collection

On Error GoTo Erro_Carrega_Dados_Pedidos
        
    Call Inicializa_QuantAssocia_ItenRC(gcolRequisicaoCompra)
    
    'Para cada item da concorrência
    For Each objItemConcorrencia In objConcorrencia.colItens
        
        If objItemConcorrencia.lFornecedor > 0 And objItemConcorrencia.iFilial > 0 Then
            Set colPedidoCompras = colPedCompraExclu
        Else
            Set colPedidoCompras = colPedCompraGeral
        End If
        
        Call Transfere_Dados_Cotacoes(objItemConcorrencia.colCotacaoItemConc, colCotItemConcAux)
        
        'Para cada destino do item de concorrencia
        For Each objQuantSupl In objItemConcorrencia.colQuantSuplementar
            
            dQuantSupl = objQuantSupl.dQuantidade
        
            For Each objCotItemConc In colCotItemConcAux
                            
                If (objCotItemConc.iEscolhido = MARCADO) And (objCotItemConc.dQuantidadeComprar > 0) Then
                                        
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
                           objPedidoCompra.iCondicaoPagto = Codigo_Extrai(objCotItemConc.sCondPagto) And _
                           objPedidoCompra.iTipoDestino = objQuantSupl.iTipoDestino And _
                           objPedidoCompra.lFornCliDestino = objQuantSupl.lFornCliDestino And _
                           objPedidoCompra.iFilialDestino = objQuantSupl.iFilialDestino Then
                           
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
                        objPedidoCompra.iFilialDestino = objQuantSupl.iFilialDestino
                        objPedidoCompra.iTipoDestino = objQuantSupl.iTipoDestino
                        objPedidoCompra.lFornCliDestino = objQuantSupl.lFornCliDestino
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
                    objItemPC.dtDataLimite = objItemConcorrencia.dtDataNecessidade
                    objItemPC.iStatus = ITEM_PED_COMPRAS_ABERTO
                    objItemPC.iTipoOrigem = TIPO_ORIGEM_COTACAOITEMCONC
                    objItemPC.sDescProduto = objItemConcorrencia.sDescricao
                    objItemPC.sProduto = objItemConcorrencia.sProduto
                    objItemPC.sUM = objCotItemConc.sUMCompra
                    objItemPC.lNumIntOrigem = objCotItemConc.lNumIntDoc
                    
                    If dQuantSupl <= objCotItemConc.dQuantidadeComprar Then
                        objItemPC.dQuantidade = dQuantSupl
                        objCotItemConc.dQuantidadeComprar = objCotItemConc.dQuantidadeComprar - dQuantSupl
                        dQuantSupl = 0
                    Else
                        objItemPC.dQuantidade = objCotItemConc.dQuantidadeComprar
                        dQuantSupl = dQuantSupl - objCotItemConc.dQuantidadeComprar
                        objCotItemConc.dQuantidadeComprar = 0
                    End If
                    
                    objPedidoCompra.colItens.Add objItemPC
                    
                    'Vincula qt a comprar de ItensRC do mesmo destino do PC ao ItemPC gerado
                    lErro = Inclui_Quant_ItemReqCompra(objItemPC, objItemConcorrencia, objQuantSupl, gcolRequisicaoCompra, colProdutos)
                    If lErro <> SUCESSO Then gError 86150
                    
                    'Adiciona o item de cotação na coleção de itens de cotacao
                    lErro = colItensCotacao_Adiciona(objCotItemConc.lItemCotacao, colItensCotacao)
                    If lErro <> SUCESSO Then gError 62726
                
                End If
                
                If dQuantSupl = 0 Then Exit For
            Next
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

        Case 63799, 70484, 62726, 62727, 86150
            'Erros tratados nas rotinas chamadas

        Case 63800, 70485, 76246
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161153)

    End Select

    Exit Function

End Function

Private Sub BotaoGeraPedidos_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGeraPedidos_Click

    'Chama Gravar_Pedidos
    lErro = Gravar_Pedidos()
    If lErro <> SUCESSO Then gError 67223

    'Limpa a tela
    Call Limpa_Tela_GeracaoPedCompraCot

    iAlterado = 0

    Exit Sub

Erro_BotaoGeraPedidos_Click:

    Select Case gErr

        Case 67223

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161154)

    End Select

    Exit Sub

End Sub

Private Sub OrdenacaoCot_Click()

Dim lErro As Long

On Error GoTo Erro_Ordenacao_Click

    If gsOrdenacao = "" Then Exit Sub

    If gsOrdenacao <> OrdenacaoCot.Text Then
    
        gsOrdenacao = OrdenacaoCot.Text
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161155)

    End Select

    Exit Sub

End Sub

Sub Calcula_Preferencia(objCotItemConc As ClassCotacaoItemConc, sProduto As String, dQuantComprar As Double)
'Calcula a Preferência

Dim iIndice As Integer
Dim dQuantPreferencial As Double
Dim dQuantComprarItem As Double
    
    dQuantPreferencial = 0
    
    If dQuantComprar = 0 Then Exit Sub
    
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes
    
        If StrParaInt(GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col)) = MARCADO Then
        
            If GridItensRequisicoes.TextMatrix(iIndice, iGrid_ProdutoItemRC_Col) = sProduto And _
              GridItensRequisicoes.TextMatrix(iIndice, iGrid_FilialFornItemRC_Col) = objCotItemConc.sFilial And _
              GridItensRequisicoes.TextMatrix(iIndice, iGrid_FornecedorItemRC_Col) = objCotItemConc.sFornecedor And _
              GridItensRequisicoes.TextMatrix(iIndice, iGrid_ExclusivoItemRC_Col) = "Preferencial" Then
                
                Call Busca_QuantComprar_ItemReq(StrParaLong(GridItensRequisicoes.TextMatrix(iIndice, iGrid_CodigoReqItem_Col)), Codigo_Extrai(GridItensRequisicoes.TextMatrix(iIndice, iGrid_FilialReqItem_Col)), StrParaInt(GridItensRequisicoes.TextMatrix(iIndice, iGrid_Item_Col)), dQuantComprarItem)
              
                dQuantPreferencial = dQuantPreferencial + dQuantComprarItem
            End If
        End If
    Next
            
    objCotItemConc.dPreferencia = dQuantPreferencial / dQuantComprar
    
    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera objetos globais
    Set gobjGeracaoPedCompraCot = Nothing

    Set objGridCotacao = Nothing
    Set objGridRequisicoes = Nothing
    Set objGridItensRequisicoes = Nothing
    Set objGridProdutos1 = Nothing
    Set objGridProdutos2 = Nothing
    Set objGridCotacoes = Nothing

    Set objEventoFornecedor = Nothing
    Set objEventoPedCotacao = Nothing

    Set gcolRequisicaoCompra = Nothing
    Set gColCotacoes = Nothing
    Set gcolItemConcorrencia = Nothing

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
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
        If objItemConcorrencia.iEscolhido = MARCADO Then
            'Coloca na coleção as cotações que aparecem na tela
             For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
                    
                Set objCotItemConcAux = New ClassCotacaoItemConcAux
                
                Set objCotItemConcAux.objCotacaoItemConc = objItemCotItemConc
                objCotItemConcAux.sCondPagto = objItemCotItemConc.sCondPagto
                objCotItemConcAux.sDescricao = objItemConcorrencia.sDescricao
                objCotItemConcAux.sFilial = objItemCotItemConc.sFilial
                objCotItemConcAux.sFornecedor = objItemCotItemConc.sFornecedor
                objCotItemConcAux.sProduto = objItemConcorrencia.sProduto
                objCotItemConcAux.dtDataNecessidade = objItemConcorrencia.dtDataNecessidade
                
                colGeracao.Add objCotItemConcAux
             Next
        End If
    Next
    
    'Carrega os campos base para a ordenação utilizados na rotina de ordenação
    Call Monta_Colecao_Campos_Cotacao(colCampos, OrdenacaoCot.ListIndex)

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
        Produto1.PromptInclude = False
        Produto1.Text = sProdutoMascarado
        Produto1.PromptInclude = True
        
        GridCotacoes.TextMatrix(iIndice, iGrid_ProdutoCot_Col) = Produto1.Text
        GridCotacoes.TextMatrix(iIndice, iGrid_DescProdutoCot_Col) = objCotItemConcAux.sDescricao
        GridCotacoes.TextMatrix(iIndice, iGrid_CondPagtoCot_Col) = objCotItemConcAux.objCotacaoItemConc.sCondPagto
        
        GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar)

        GridCotacoes.TextMatrix(iIndice, iGrid_UMCot_Col) = objCotItemConcAux.objCotacaoItemConc.sUMCompra
        GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitarioCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        
        If objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha <> MOTIVO_EXCLUSIVO_DESCRICAO Then
            Call Calcula_Preferencia(objCotItemConcAux.objCotacaoItemConc, Produto1.Text, objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar)
            GridCotacoes.TextMatrix(iIndice, iGrid_Preferencia_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPreferencia, "Percent")
        Else
            GridCotacoes.TextMatrix(iIndice, iGrid_Preferencia_Col) = "Exclusivo"
        End If
        
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
                                          
        GridCotacoes.TextMatrix(iIndice, iGrid_ValorPresenteCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dValorPresente, ValorPresente.Format) 'Alterado por Wagner
        
        If objCotItemConcAux.objCotacaoItemConc.iMoeda <> MOEDA_REAL Then
            GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado * objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar * objCotItemConcAux.objCotacaoItemConc.dTaxa, "STANDARD")
        Else
            GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado * objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar, "STANDARD")
        End If
        
        GridCotacoes.TextMatrix(iIndice, iGrid_FornecedorCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFornecedor
        GridCotacoes.TextMatrix(iIndice, iGrid_FilialFornCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFilial
        GridCotacoes.TextMatrix(iIndice, iGrid_PedidoCot_Col) = objCotItemConcAux.objCotacaoItemConc.lPedCotacao
        If objCotItemConcAux.objCotacaoItemConc.dQuantEntrega > 0 Then GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeEntrega_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantEntrega)
        
        'Data da Cotacao
        If objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataCotacaoCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao, "dd/mm/yyyy")
        End If
    
        For iIndice2 = 0 To TipoTributacaoCot.ListCount - 1
            If objCotItemConcAux.objCotacaoItemConc.iTipoTributacao = TipoTributacaoCot.ItemData(iIndice2) Then
                GridCotacoes.TextMatrix(iIndice, iGrid_TipoTributacaoCot_Col) = TipoTributacaoCot.List(iIndice2)
                Exit For
            End If
        Next
        
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaIPI, "Percent")
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaICMS, "Percent")
        
        'Data de Validade
        If objCotItemConcAux.objCotacaoItemConc.dtDataValidade <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataValidadeCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dtDataValidade, "dd/mm/yyyy")
        End If

        'Prazo de Entrega
        If objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega <> 0 Then
            GridCotacoes.TextMatrix(iIndice, iGrid_PrazoEntrega_Col) = objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega
            GridCotacoes.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(DateAdd("d", objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega, Date), "dd/mm/yyyy")
        End If

        'Data de Entrega
        If objCotItemConcAux.objCotacaoItemConc.dtDataEntrega <> DATA_NULA Then
        End If
                
        'Quantidade a comprar Máxima
        GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantCotada)

        'Motivo escolha
        GridCotacoes.TextMatrix(iIndice, iGrid_MotivoEscolhaCot_Col) = objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha
        
        If objCotItemConcAux.dtDataNecessidade <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataNecessidade_Col) = Format(objCotItemConcAux.dtDataNecessidade, "dd/mm/yyyy")
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161156)

    End Select

    Exit Function

End Function

Function GridItensReq_Preenche() As Long
'Preenche o GridItensRequisicoes com os Itens da Requisicao passada como parametro

Dim lErro As Long
Dim objRequisicao As ClassRequisicaoCompras
Dim objItemReqCompras As New ClassItemReqCompras
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iLinha As Integer
Dim sCclMascarado As String
Dim objFilialEmpresa As New AdmFiliais
Dim sProdutoMascarado As String
Dim objlCodigoNome As AdmlCodigoNome
Dim colFiliais As New AdmCollCodigoNome
Dim colAlmoxarifados As New AdmCollCodigoNome
Dim colFornecedor As New AdmCollCodigoNome
Dim colFilialForn As New Collection
Dim iPosicao As Integer
Dim objObservacao As New ClassObservacao

On Error GoTo Erro_GridItensReq_Preenche

    'Limpa o grid de itens
    Call Grid_Limpa(objGridItensRequisicoes)
    
    'Para cada requisicao
    For Each objRequisicao In gcolRequisicaoCompra
        'Se a req está selecionada
        If objRequisicao.iSelecionado = MARCADO Then
            'Para cada item
            For Each objItemReqCompras In objRequisicao.colItens
        
                iLinha = iLinha + 1
                'BUsca a filial da req na colfiliais
                Call Busca_Na_Colecao(colFiliais, objRequisicao.iFilialEmpresa, iPosicao)
            
                If iPosicao = 0 Then
               
                    objFilialEmpresa.iCodFilial = objRequisicao.iFilialEmpresa
                    'Lê a FilialEmpresa
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
                    If lErro <> SUCESSO And lErro <> 27378 Then gError 68059
        
                    'Se não encontrou a filial ==>erro
                    If lErro = 27378 Then gError 68060
        
                    Set objlCodigoNome = New AdmlCodigoNome
                    
                    objlCodigoNome.lCodigo = objFilialEmpresa.iCodFilial
                    objlCodigoNome.sNome = objFilialEmpresa.sNome
                    
                    colFiliais.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
        
                Else
                    Set objlCodigoNome = colFiliais(iPosicao)
                End If
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = objItemReqCompras.iSelecionado
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialReqItem_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_CodigoReqItem_Col) = objRequisicao.lCodigo
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_Item_Col) = objItemReqCompras.iItem
        
                'Mascara o Produto
                lErro = Mascara_RetornaProdutoEnxuto(objItemReqCompras.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 68064
        
                ProdutoItemRC.PromptInclude = False
                ProdutoItemRC.Text = sProdutoMascarado
                ProdutoItemRC.PromptInclude = True
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_ProdutoItemRC_Col) = ProdutoItemRC.Text
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_DescProdutoItem_Col) = objItemReqCompras.sDescProduto
                
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_UnidadeMedItem_Col) = objItemReqCompras.sUM
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantidadeItem_Col) = Formata_Estoque(objItemReqCompras.dQuantidade)
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantPedida_Col) = Formata_Estoque(objItemReqCompras.dQuantPedida)
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantRecebida_Col) = Formata_Estoque(objItemReqCompras.dQuantRecebida)
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantComprarItem_Col) = Formata_Estoque(objItemReqCompras.dQuantComprar)
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantComprarItem_Col) = Formata_Estoque(objItemReqCompras.dQuantidade - objItemReqCompras.dQuantRecebida - objItemReqCompras.dQuantPedida - objItemReqCompras.dQuantCancelada)
        
                If objItemReqCompras.iAlmoxarifado <> 0 Then
                    
                    Call Busca_Na_Colecao(colAlmoxarifados, objItemReqCompras.iAlmoxarifado, iPosicao)
                
                    If iPosicao = 0 Then
                
                        objAlmoxarifado.iCodigo = objItemReqCompras.iAlmoxarifado
            
                        'Lê o almoxarifado
                        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25056 Then gError 63984
            
                        'Se não encontrou ==> Erro
                        If lErro = 25056 Then gError 63985
        
                        Set objlCodigoNome = New AdmlCodigoNome
                        
                        objlCodigoNome.lCodigo = objAlmoxarifado.iCodigo
                        objlCodigoNome.sNome = objAlmoxarifado.sNomeReduzido
                        
                        colAlmoxarifados.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
        
                    Else
                        Set objlCodigoNome = colAlmoxarifados(iPosicao)
                    End If
                
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = objlCodigoNome.sNome
                
                End If
        
                If Len(Trim(objItemReqCompras.sCcl)) > 0 Then
        
                    'Mascara o Ccl
                    lErro = Mascara_MascararCcl(objItemReqCompras.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then gError 63986
                    
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_CclItemRC_Col) = sCclMascarado
                End If
        
                If objItemReqCompras.lFornecedor <> 0 And objItemReqCompras.iFilial <> 0 Then
                    
                    Call Busca_Na_Colecao(colFornecedor, objItemReqCompras.lFornecedor, iPosicao)
        
                    If iPosicao = 0 Then
        
                        objFornecedor.lCodigo = objItemReqCompras.lFornecedor
            
                        'Lê o Fornecedor
                        lErro = CF("Fornecedor_Le", objFornecedor)
                        If lErro <> SUCESSO And lErro <> 12729 Then gError 63987
            
                        'Se não encontrou o Fornecedor==> Erro
                        If lErro = 12729 Then gError 63988
                        
                        Set objlCodigoNome = New AdmlCodigoNome
                    
                        objlCodigoNome.lCodigo = objFornecedor.lCodigo
                        objlCodigoNome.sNome = objFornecedor.sNomeReduzido
                        
                        colFornecedor.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
                    
                    Else
                        Set objlCodigoNome = colFornecedor(iPosicao)
                    End If
        
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_FornecedorItemRC_Col) = objlCodigoNome.sNome
        
                    Call Busca_FilialForn(colFilialForn, objItemReqCompras.lFornecedor, objItemReqCompras.iFilial, iPosicao)
                    
                    If iPosicao = 0 Then
                        Set objFilialFornecedor = New ClassFilialFornecedor
                        objFilialFornecedor.iCodFilial = objItemReqCompras.iFilial
                        objFilialFornecedor.lCodFornecedor = objItemReqCompras.lFornecedor
                        
                        'Lê a FilialFornecedor
                        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                        If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
            
                        'Se não encontrou==>Erro
                        If lErro = 12929 Then gError 63990
                    Else
                        Set objFilialFornecedor = colFilialForn(iPosicao)
                    End If
        
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialFornItemRC_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
                
                    If objItemReqCompras.iExclusivo = MARCADO Then
                        GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusivoItemRC_Col) = "Exclusivo"
                    Else
                        GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusivoItemRC_Col) = "Preferencial"
                    End If
                    
                End If
        
                'Verifica se Observacao está preenchida
                If Len(Trim(objItemReqCompras.sObservacao)) = 0 And objItemReqCompras.lObservacao > 0 Then
        
                    objObservacao.lNumInt = objItemReqCompras.lObservacao
        
                    'Lê a observacao
                    lErro = CF("Observacao_Le", objObservacao)
                    If lErro <> SUCESSO And lErro <> 53827 Then gError 63577
        
                    'Se não encontrou a Observacao ==> erro
                    If lErro = 53827 Then gError 63578
                    
                    objItemReqCompras.sObservacao = objObservacao.sObservacao
        
                End If
                'Verifica se Observacao está preenchida
                If Len(Trim(objItemReqCompras.sObservacao)) = 0 And objItemReqCompras.lObservacao > 0 Then
        
                    objObservacao.lNumInt = objItemReqCompras.lObservacao
        
                    'Lê a observacao
                    lErro = CF("Observacao_Le", objObservacao)
                    If lErro <> SUCESSO And lErro <> 53827 Then gError 63577
        
                    'Se não encontrou a Observacao ==> erro
                    If lErro = 53827 Then gError 63578
                    
                    objItemReqCompras.sObservacao = objObservacao.sObservacao
        
                End If
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_ObservacaoItemRC_Col) = objItemReqCompras.sObservacao
        
            Next
        End If
    Next
    
    'Atualiza o número de linhas existentes do GridItensRequisicoes
    objGridItensRequisicoes.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)
    
    GridItensReq_Preenche = SUCESSO

    Exit Function

Erro_GridItensReq_Preenche:

    GridItensReq_Preenche = gErr

    Select Case gErr

        Case 63982, 63984, 63986, 63987, 63989, 68059, 68064, 63577
            'Erros tratados nas rotinas chamadas

        Case 63983
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO", gErr, objRequisicao.lNumIntDoc, objItemReqCompras.lReqCompra)

        Case 63985
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 63988
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63990
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFornecedor.lCodigo)

        Case 68060
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 63578
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161157)

    End Select

    Exit Function
    
End Function

Function Grids_Produto_Preenche() As Long

Dim iLinha1 As Integer, iLinha2 As Integer
Dim objItemConc As New ClassItemConcorrencia
Dim sProdutoEnxuto As String
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objQuantSupl As ClassQuantSuplementar
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colItensSaida As New Collection
Dim colCampos As New Collection
Dim colFilForn As New Collection
Dim colFornec As New AdmCollCodigoNome
Dim objCodNome As New AdmlCodigoNome
Dim iPosicao As Integer

On Error GoTo Erro_Grids_Produto_Preenche
    
    'Limpa o grid de produtos1
    Call Grid_Limpa(objGridProdutos1)
    
    colCampos.Add "sProduto"
    colCampos.Add "lFornecedor"
    colCampos.Add "iFilial"
    
    'Ordena os itens de concorrência por produto
    lErro = Ordena_Colecao(gcolItemConcorrencia, colItensSaida, colCampos)
    If lErro <> SUCESSO Then gError 63808

    Set gcolItemConcorrencia = colItensSaida
    
    iLinha1 = 0
    iLinha2 = 0
    
    'Para cada item de concorrência
    For Each objItemConc In gcolItemConcorrencia
        
        iLinha1 = iLinha1 + 1
        'Preenche a seleção
        GridProdutos1.TextMatrix(iLinha1, iGrid_EscolhidoProduto_Col) = objItemConc.iEscolhido
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemConc.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 62778
        
        Produto1.PromptInclude = False
        Produto1.Text = sProdutoEnxuto
        Produto1.PromptInclude = True
        
        'Preenche o produto
        GridProdutos1.TextMatrix(iLinha1, iGrid_Produto1_Col) = Produto1.Text
        GridProdutos1.TextMatrix(iLinha1, iGrid_DescProduto1_Col) = objItemConc.sDescricao
        GridProdutos1.TextMatrix(iLinha1, iGrid_UnidadeMed1_Col) = objItemConc.sUM
        GridProdutos1.TextMatrix(iLinha1, iGrid_Quantidade1_Col) = Formata_Estoque(objItemConc.dQuantidade)
        GridProdutos1.TextMatrix(iLinha1, iGrid_Urgente_Col) = Formata_Estoque(objItemConc.dQuantUrgente)
        
        'Se o Fornecedor está preenchido
        If objItemConc.lFornecedor > 0 And objItemConc.iFilial > 0 Then
            
            'verifica se esse forn já foi lido
            Call Busca_Na_Colecao(colFornec, objItemConc.lFornecedor, iPosicao)
        
            If iPosicao = 0 Then
                objFornecedor.lCodigo = objItemConc.lFornecedor
                
                lErro = CF("Fornecedor_Le", objFornecedor)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 62779
                If lErro <> SUCESSO Then gError 62780
                            
                Set objCodNome = New AdmlCodigoNome
                
                objCodNome.lCodigo = objFornecedor.lCodigo
                objCodNome.sNome = objFornecedor.sNomeReduzido
                
                colFornec.Add objCodNome.lCodigo, objCodNome.sNome
            Else
                Set objCodNome = colFornec(iPosicao)
            End If
            
            'Preenche o fornecedor
            GridProdutos1.TextMatrix(iLinha1, iGrid_Fornecedor1_Col) = objCodNome.sNome
            
            'Verifica se essa filial já foi lida
            Call Busca_FilialForn(colFilForn, objItemConc.lFornecedor, objItemConc.iFilial, iPosicao)
            
            If iPosicao = 0 Then
                Set objFilialFornecedor = New ClassFilialFornecedor
                objFilialFornecedor.lCodFornecedor = objItemConc.lFornecedor
                objFilialFornecedor.iCodFilial = objItemConc.iFilial
                
                lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
                
                'Se não encontrou==>Erro
                If lErro = 12929 Then gError 63990
                
                colFilForn.Add objFilialFornecedor
            Else
                Set objFilialFornecedor = colFilForn(iPosicao)
            End If
            'Preenche a filial
            GridProdutos1.TextMatrix(iLinha1, iGrid_FilialForn1_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        
        End If
    Next
    
    objGridProdutos1.iLinhasExistentes = iLinha1
    
    Call Grid_Refresh_Checkbox(objGridProdutos1)
    
    'Preenche o grid de produtos 2
    lErro = GridProdutos2_Preenche
    If lErro <> SUCESSO Then gError 62781
    
    Call GridCotacoes_Preenche
    
    Grids_Produto_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grids_Produto_Preenche:

    Grids_Produto_Preenche = gErr
    
    Select Case gErr
        
        Case 63808, 62778, 62779, 63989, 62781
        
        Case 62780
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case 63990
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161158)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Pedidos de Compra por Geração de Cotações"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoPedCompraGerCot"

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


Private Sub Concorrencia_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Concorrencia, Source, X, Y)
End Sub

Private Sub Concorrencia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Concorrencia, Button, Shift, X, Y)
End Sub
Private Sub TaxaEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TaxaEmpresa, Source, X, Y)
End Sub

Private Sub TaxaEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TaxaEmpresa, Button, Shift, X, Y)
End Sub

Private Sub GridPedCotacao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCotacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCotacao, iAlterado)
    End If

End Sub

Private Sub GridPedCotacao_EnterCell()

    Call Grid_Entrada_Celula(objGridCotacao, iAlterado)

End Sub

Private Sub GridPedCotacao_GotFocus()

    Call Grid_Recebe_Foco(objGridCotacao)

End Sub

Private Sub GridPedCotacao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCotacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCotacao, iAlterado)
    End If

End Sub

Private Sub GridPedCotacao_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

    Call Grid_Trata_Tecla1(KeyCode, objGridCotacao)
    
End Sub

Private Sub GridPedCotacao_LeaveCell()
    
    Call Saida_Celula(objGridCotacao)

End Sub

Private Sub GridPedCotacao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCotacao)

End Sub

Private Sub GridPedCotacao_RowColChange()

    Call Grid_RowColChange(objGridCotacao)

End Sub

Private Sub GridPedCotacao_Scroll()

    Call Grid_Scroll(objGridCotacao)

End Sub

Private Sub Selecionado_Click()

    iAlterado = REGISTRO_ALTERADO

    'Se a linha do Grid selecionada foi diferente da última que tinha sido selecionada
    If GridPedCotacao.Row <> gobjGeracaoPedCompraCot.iCotacaoSel Then
        iCotacaoAlterada = REGISTRO_ALTERADO
    End If

    If gobjGeracaoPedCompraCot.iCotacaoSel = GridPedCotacao.Row Then Exit Sub
    gobjGeracaoPedCompraCot.iCotacaoSel = GridPedCotacao.Row
    
    Call Traz_Cotacao_Tela

    Exit Sub

End Sub

Private Sub Selecionado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacao)

End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacao)

End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacao.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridCotacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRequisicoes, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If
    
    Exit Sub

End Sub

Private Sub GridRequisicoes_GotFocus()
    Call Grid_Recebe_Foco(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_EnterCell()
    Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
End Sub

Private Sub GridRequisicoes_LeaveCell()
    Call Saida_Celula(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRequisicoes)
    
End Sub

Private Sub GridRequisicoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If
   
    Exit Sub
    
End Sub

Private Sub GridRequisicoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_RowColChange()
    Call Grid_RowColChange(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_Scroll()
    Call Grid_Scroll(objGridRequisicoes)
End Sub
Private Sub EscolhidoReq_Click()
    
    iAlterado = REGISTRO_ALTERADO
    Call Requisicoes_Atualiza
    
End Sub

Private Sub EscolhidoReq_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)
End Sub

Private Sub EscolhidoReq_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)
End Sub

Private Sub EscolhidoReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = EscolhidoReq
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridItensRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridItensRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensRequisicoes, iAlterado)
    End If
   
    Exit Sub

End Sub

Private Sub GridItensRequisicoes_GotFocus()
    Call Grid_Recebe_Foco(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_EnterCell()
    Call Grid_Entrada_Celula(objGridItensRequisicoes, iAlterado)
End Sub

Private Sub GridItensRequisicoes_LeaveCell()
    Call Saida_Celula(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_KeyDown(KeyCode As Integer, Shift As Integer)
   
    Call Grid_Trata_Tecla1(KeyCode, objGridItensRequisicoes)
    
End Sub

Private Sub GridItensRequisicoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensRequisicoes, iAlterado)
    End If
    
    Exit Sub
        
End Sub

Private Sub GridItensRequisicoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_RowColChange()
    Call Grid_RowColChange(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_Scroll()
    Call Grid_Scroll(objGridItensRequisicoes)
End Sub
Private Sub EscolhidoItem_Click()
    
    iAlterado = REGISTRO_ALTERADO
                        
    Call Atualiza_ItensReq

End Sub

Private Sub EscolhidoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub EscolhidoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub EscolhidoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = EscolhidoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantComprarItemRC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantComprarItemRC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub QuantComprarItemRC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub QuantComprarItemRC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = QuantComprarItemRC
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridProdutos1_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)
    End If

End Sub

Private Sub GridProdutos1_GotFocus()
    Call Grid_Recebe_Foco(objGridProdutos1)
End Sub

Private Sub GridProdutos1_EnterCell()
    Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)
End Sub

Private Sub GridProdutos1_LeaveCell()
    Call Saida_Celula(objGridProdutos1)
End Sub

Private Sub GridProdutos1_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos1)
    
End Sub

Private Sub GridProdutos1_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)
    End If

End Sub

Private Sub GridProdutos1_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridProdutos1)
End Sub

Private Sub GridProdutos1_RowColChange()
    Call Grid_RowColChange(objGridProdutos1)
End Sub

Private Sub GridProdutos1_Scroll()
    Call Grid_Scroll(objGridProdutos1)
End Sub

Private Sub EscolhidoProduto_Click()
    
Dim lErro As Long
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_EscolhidoProduto_Click

    iAlterado = REGISTRO_ALTERADO

    'Se tiver clicado em escolhido
    If GridProdutos1.Col = iGrid_EscolhidoProduto_Col And objGridProdutos1.iLinhasExistentes > 0 Then
        
        'Pega o item de concorrência clicado
        Set objItemConcorrencia = gcolItemConcorrencia(GridProdutos1.Row)
        'Atualiza a escolha
        objItemConcorrencia.iEscolhido = GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_EscolhidoProduto_Col)

        'Repreenche o grid de produtos
        lErro = GridProdutos2_Preenche
        If lErro <> SUCESSO Then gError 62758
        
        Call Indica_Melhores
        Call GridCotacoes_Preenche
        
    End If

    Exit Sub
    
Erro_EscolhidoProduto_Click:
    
    Select Case gErr

        Case 62758

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161159)

    End Select

    Exit Sub

End Sub

Private Sub EscolhidoProduto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridProdutos1)
End Sub

Private Sub EscolhidoProduto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos1)
End Sub

Private Sub EscolhidoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos1.objControle = EscolhidoProduto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'GridProdutos2
Private Sub GridProdutos2_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)
    End If

End Sub

Private Sub GridProdutos2_EnterCell()

    Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)

End Sub

Private Sub GridProdutos2_GotFocus()

    Call Grid_Recebe_Foco(objGridProdutos2)

End Sub

Private Sub GridProdutos2_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos2, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)
    End If

End Sub

Private Sub GridProdutos2_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridProdutos2_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos2)

    Exit Sub

Erro_GridProdutos2_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161160)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos2_LeaveCell()

    Call Saida_Celula(objGridProdutos2)

End Sub

Private Sub GridProdutos2_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridProdutos2)

End Sub

Private Sub GridProdutos2_RowColChange()

    Call Grid_RowColChange(objGridProdutos2)

End Sub

Private Sub GridProdutos2_Scroll()

    Call Grid_Scroll(objGridProdutos2)

End Sub

Private Sub Quantidade2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade2_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos2)

End Sub

Private Sub Quantidade2_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos2)

End Sub

Private Sub Quantidade2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos2.objControle = Quantidade2
    lErro = Grid_Campo_Libera_Foco(objGridProdutos2)
    If lErro <> SUCESSO Then Cancel = True

End Sub

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

Private Sub EscolhidoCot_Click()
 
Dim objCotItemConc As ClassCotacaoItemConc

    iAlterado = REGISTRO_ALTERADO
    
    'Localiza a cotação correspondente
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    'Atuzaliza a escolha
    objCotItemConc.iEscolhido = GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_EscolhidoCot_Col)
    'Recalcula o total dos itens selecionados
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

Private Sub MotivoEscolhaCot_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MotivoEscolhaCot_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub MotivoEscolhaCot_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub MotivoEscolhaCot_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = MotivoEscolhaCot
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

Private Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Carrega_TipoTributacao() As Long
'Carrega Tipos de Tributação

Dim lErro As Long
Dim colTributacao As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoTributacao

    'Lê os Tipos de Tributação associadas a Compras
    lErro = CF("TiposTributacaoCompras_Le", colTributacao)
    If lErro <> SUCESSO Then gError 66123

    'Carrega Tipos de Tributação
    For iIndice = 1 To colTributacao.Count
        TipoTributacaoCot.AddItem colTributacao(iIndice).iCodigo & SEPARADOR & colTributacao(iIndice).sNome
        TipoTributacaoCot.ItemData(TipoTributacaoCot.NewIndex) = colTributacao(iIndice).iCodigo
    Next

    Carrega_TipoTributacao = SUCESSO

    Exit Function

Erro_Carrega_TipoTributacao:

    Carrega_TipoTributacao = gErr

    Select Case gErr

        Case 66123

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161161)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialEmpresa() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o Código e o Nome de toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 63873

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

        Case 63873
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161162)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Requisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Requisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Requisicoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("Data RC")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Requisitante")
    objGridInt.colColuna.Add ("Centro C/L")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoReq.Name)
    objGridInt.colCampo.Add (FilialReq.Name)
    objGridInt.colCampo.Add (CodigoReq.Name)
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (DataReq.Name)
    objGridInt.colCampo.Add (Urgente.Name)
    objGridInt.colCampo.Add (Requisitante.Name)
    objGridInt.colCampo.Add (CclReq.Name)
    objGridInt.colCampo.Add (ObservacaoReq.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoReq_Col = 1
    iGrid_FilialReq_Col = 2
    iGrid_CodigoReq_Col = 3
    iGrid_DataLimite_Col = 4
    iGrid_DataReq_Col = 5
    iGrid_Urgente_Col = 6
    iGrid_Requisitante_Col = 7
    iGrid_CclReq_Col = 8
    iGrid_ObservacaoReq_Col = 9

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REQUISICOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Requisicoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Requisicoes:

    Inicializa_Grid_Requisicoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161163)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Cotacoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Cotacoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Cotacoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Cond. Pagto")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Taxa Forn.")
    objGridInt.colColuna.Add ("Cotação")
    objGridInt.colColuna.Add ("Preço Unitário (R$)")
    objGridInt.colColuna.Add ("Preferência")
    objGridInt.colColuna.Add ("Quant. Cotada")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Valor Presente (R$)")
    objGridInt.colColuna.Add ("Valor Item (R$)")

    objGridInt.colColuna.Add ("Tipo Tributacao")
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
    objGridInt.colCampo.Add (CondPagto.Name)
    objGridInt.colCampo.Add (FornecedorCot.Name)
    objGridInt.colCampo.Add (FilialFornCot.Name)
    objGridInt.colCampo.Add (Moeda.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (TaxaForn.Name)
    objGridInt.colCampo.Add (Cotacao.Name)
    objGridInt.colCampo.Add (PrecoUnitarioReal.Name)
    objGridInt.colCampo.Add (Preferencia.Name)
    objGridInt.colCampo.Add (QuantidadeCot.Name)
    objGridInt.colCampo.Add (QuantComprarCot.Name)
    objGridInt.colCampo.Add (UnidadeMedCot.Name)

    objGridInt.colCampo.Add (ValorPresente.Name)
    objGridInt.colCampo.Add (ValorItem.Name)
    objGridInt.colCampo.Add (TipoTributacaoCot.Name)
    objGridInt.colCampo.Add (AliquotaIPI.Name)
    objGridInt.colCampo.Add (AliquotaICMS.Name)
    objGridInt.colCampo.Add (PedCotacao.Name)
    objGridInt.colCampo.Add (DataCotacao.Name)
    objGridInt.colCampo.Add (DataValidade.Name)
    objGridInt.colCampo.Add (PrazoEntrega.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (DataNecessidade.Name)
    objGridInt.colCampo.Add (QuantidadeEntrega.Name)
    objGridInt.colCampo.Add (MotivoEscolhaCot.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoCot_Col = 1
    iGrid_ProdutoCot_Col = 2
    iGrid_DescProdutoCot_Col = 3
    iGrid_CondPagtoCot_Col = 4
    iGrid_FornecedorCot_Col = 5
    iGrid_FilialFornCot_Col = 6
    iGrid_Moeda_Col = 7
    iGrid_PrecoUnitarioCot_Col = 8
    iGrid_TaxaForn_Col = 9
    iGrid_CotacaoMoeda_Col = 10
    iGrid_PrecoUnitario_RS_Col = 11
    iGrid_Preferencia_Col = 12
    iGrid_QuantidadeCot_Col = 13
    iGrid_QuantComprarCot_Col = 14
    iGrid_UMCot_Col = 15
    iGrid_ValorPresenteCot_Col = 16
    iGrid_ValorItem_Col = 17
    iGrid_TipoTributacaoCot_Col = 18
    iGrid_AliquotaIPI_Col = 19
    iGrid_AliquotaICMS_Col = 20
    iGrid_PedidoCot_Col = 21
    iGrid_DataCotacaoCot_Col = 22
    iGrid_DataValidadeCot_Col = 23
    iGrid_PrazoEntrega_Col = 24
    iGrid_DataEntrega_Col = 25
    iGrid_DataNecessidade_Col = 26
    iGrid_QuantidadeEntrega_Col = 27
    iGrid_MotivoEscolhaCot_Col = 28

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridCotacoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_COTACOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridCotacoes.Width = 8295
    GridCotacoes.ColWidth(0) = 350

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cotacoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Cotacoes:

    Inicializa_Grid_Cotacoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161164)

    End Select

    Exit Function

End Function

Private Function Carrega_MotivoEscolha() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_MotivoEscolha

    'Lê o Código e o Nome de todo MotivoEscolha do BD
    lErro = CF("Cod_Nomes_Le", "Motivo", "Codigo", "Motivo", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 63875

    'Carrega a combo de Motivo Escolha com código e nome
    For Each objCodigoNome In colCodigoNome

        'Verifica se o MotivoEscolha é diferente de Exclusividade
        If objCodigoNome.iCodigo <> MOTIVO_EXCLUSIVO Then

            MotivoEscolhaCot.AddItem objCodigoNome.sNome
            MotivoEscolhaCot.ItemData(MotivoEscolhaCot.NewIndex) = objCodigoNome.iCodigo

        End If

    Next

    Carrega_MotivoEscolha = SUCESSO

    Exit Function

Erro_Carrega_MotivoEscolha:

    Carrega_MotivoEscolha = gErr

    Select Case gErr

        Case 63875
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161165)

    End Select

    Exit Function

End Function

Private Sub BotaoGravaConcorrencia_Click()
'Grava a Concorrencia

Dim lErro As Long

On Error GoTo Erro_BotaoGravaConcorrencia_Click
    
    'Insere ou Altera uma concorrencia no BD
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 63672
    
    Exit Sub

Erro_BotaoGravaConcorrencia_Click:

    Select Case gErr

        Case 63672

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161166)

    End Select

    Exit Sub

End Sub

Function GridRequisicoes_Preenche() As Long

Dim lErro As Long
Dim objRequisicao As New ClassRequisicaoCompras
Dim iLinha As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim colRequisitantes As New AdmCollCodigoNome
Dim colFiliais As New AdmCollCodigoNome
Dim objlCodigoNome As AdmlCodigoNome
Dim iPosicao As Integer

On Error GoTo Erro_GridRequisicoes_Preenche

    'Limpa o Grid de Requisições
    Call Grid_Limpa(objGridRequisicoes)

    If gcolRequisicaoCompra.Count > 0 Then

        'Preenche o GridRequisicoes
        For Each objRequisicao In gcolRequisicaoCompra

            iLinha = objGridRequisicoes.iLinhasExistentes + 1
    
            Call Busca_Na_Colecao(colFiliais, objRequisicao.iFilialEmpresa, iPosicao)
    
            If iPosicao = 0 Then
    
                objFilialEmpresa.iCodFilial = objRequisicao.iFilialEmpresa
    
                'Lê a FilialEmpresa
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 63976
    
                'Se não encontrou ==>Erro
                If lErro = 27378 Then gError 63977
    
                Set objlCodigoNome = New AdmlCodigoNome
    
                objlCodigoNome.lCodigo = objFilialEmpresa.iCodFilial
                objlCodigoNome.sNome = objFilialEmpresa.sNome
    
                colFiliais.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
    
            Else
    
                Set objlCodigoNome = colFiliais(iPosicao)
    
            End If
    
            'Preenche a Filial de Requisicao com código e nome reduzido
            GridRequisicoes.TextMatrix(iLinha, iGrid_FilialReq_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
            GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoReq_Col) = objRequisicao.lCodigo
    
            'Verifica se DataLimite é diferente de Data Nula
            If objRequisicao.dtDataLimite <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataLimite_Col) = Format(objRequisicao.dtDataLimite, "dd/mm/yyyy")
    
            'Verifica se Data é diferente de Data Nula
            If objRequisicao.dtData <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataReq_Col) = Format(objRequisicao.dtData, "dd/mm/yyyy")
    
            GridRequisicoes.TextMatrix(iLinha, iGrid_Urgente_Col) = objRequisicao.lUrgente
    
            Call Busca_Na_Colecao(colRequisitantes, objRequisicao.lRequisitante, iPosicao)
            
            If iPosicao = 0 Then
                objRequisitante.lCodigo = objRequisicao.lRequisitante
        
                'Lê o requisitante
                lErro = CF("Requisitante_Le", objRequisitante)
                If lErro <> SUCESSO And lErro <> 49084 Then gError 63978
        
                'Se não encontrou o Requisitante ==> Erro
                If lErro = 49084 Then gError 63979
                
                Set objlCodigoNome = New AdmlCodigoNome
                
                objlCodigoNome.lCodigo = objRequisitante.lCodigo
                objlCodigoNome.sNome = objRequisitante.sNomeReduzido
                
                colRequisitantes.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
                
            Else
                Set objlCodigoNome = colRequisitantes(iPosicao)
            End If
            
            'Preenche o Requisitante com o código e o nome reduzido
            GridRequisicoes.TextMatrix(iLinha, iGrid_Requisitante_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
    
            'Se o Ccl está preenchida
            If Len(Trim(objRequisicao.sCcl)) > 0 Then
    
                'Mascara o Produto
                lErro = Mascara_MascararCcl(objRequisicao.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError 63980
    
                'Preenche o Ccl
                GridRequisicoes.TextMatrix(iLinha, iGrid_CclReq_Col) = sCclMascarado
    
            End If
    
            'Preenche a Observacao
            GridRequisicoes.TextMatrix(iLinha, iGrid_ObservacaoReq_Col) = objRequisicao.sObservacao
    
            'Selecionado
            GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = objRequisicao.iSelecionado
                       
            objGridRequisicoes.iLinhasExistentes = iLinha
        
        Next
    
        Call Grid_Refresh_Checkbox(objGridRequisicoes)

    End If
    
    GridRequisicoes_Preenche = SUCESSO
    
    Exit Function
    
Erro_GridRequisicoes_Preenche:

    GridRequisicoes_Preenche = gErr
    
    Select Case gErr
    
        Case 63976, 63978, 63980
            'Erros tratados nas rotinas chamadas

        Case 63977
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 63979
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161167)

    End Select
        
    Exit Function
        
End Function
Function ItensConcorrencia_Cria_Altera(objItemRC As ClassItemReqCompras) As Long

Dim lErro As Long
Dim lForn As Long
Dim dFator As Double
Dim bAchou As Boolean
Dim iFilForn As Integer
Dim iPosicao As Integer
Dim objProduto As New ClassProduto
Dim objReqCompra As New ClassRequisicaoCompras
Dim objQuantSupl As ClassQuantSuplementar
Dim dQuantComprar As Double
Dim objCotItemConc As ClassCotacaoItemConc
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objItemConcorrencia As ClassItemConcorrencia
Dim dQuantReq As Double

On Error GoTo Erro_ItensConcorrencia_Cria_Altera
    
    objProduto.sCodigo = objItemRC.sProduto
    
    'Lê os dados do produto envolvido
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 23080 Then gError 62775
    If lErro <> SUCESSO Then gError 62776
    
    'Se o item Rc for exclusivo
    If objItemRC.iExclusivo = MARCADO Then
        'guarda o fornc e filial do item de conc
        lForn = objItemRC.lFornecedor: iFilForn = objItemRC.iFilial
    'Senão
    Else
        'O item não estará vinculado a filial fornecedor
        lForn = 0: iFilForn = 0
    End If
        
    'Verica se já existe um item de concorrência com os dados
    'determinados pelo item de requisição
    bAchou = False
    iPosicao = 0
    For Each objItemConcorrencia In gcolItemConcorrencia
        iPosicao = iPosicao + 1
        If objItemConcorrencia.sProduto = objItemRC.sProduto And _
           objItemConcorrencia.lFornecedor = lForn And _
           objItemConcorrencia.iFilial = iFilForn Then
           'Encontrou o item de concorrência
           bAchou = True
           Exit For
        End If
    Next

    'Busca os dados da requisição de compra ligada ao ItemRC passado
    Call Obtem_ReqCompra(gcolRequisicaoCompra, objItemRC.lReqCompra, objReqCompra)
    
    'Faz a conversão da quantidade a comprar do item para UM compra
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
    If lErro <> SUCESSO Then gError 62777
    
    dQuantComprar = objItemRC.dQuantComprar * dFator
    objItemRC.dQuantNaConcorrencia = dQuantComprar
    
    'Se o item concorrência já existe
    If bAchou Then
        'recolhe o item de concorrência
        Set objItemConcorrencia = gcolItemConcorrencia(iPosicao)
        
        objItemConcorrencia.sDescricao = objProduto.sDescricao
        
        bAchou = False
        iPosicao = 0
        'Verifica se já um registro de quant suplementar para o tipo de destino do ItemRC
        For Each objQuantSupl In objItemConcorrencia.colQuantSuplementar
            iPosicao = iPosicao + 1
            If objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino And _
               objQuantSupl.iTipoDestino = objReqCompra.iTipoDestino And _
               objQuantSupl.lFornCliDestino = objReqCompra.lFornCliDestino Then
                'encontrou
                bAchou = True
                Exit For
            End If
        Next
        
        'Se encontrou registro de quant supl.
        If bAchou Then
            'Atualiza a quantidade suplementar
            Set objQuantSupl = objItemConcorrencia.colQuantSuplementar(iPosicao)
            objQuantSupl.dQuantidade = objQuantSupl.dQuantidade + dQuantComprar
            objQuantSupl.dQuantRequisitada = objQuantSupl.dQuantRequisitada + dQuantComprar
        'Senão
        Else
            'cria um novo registro de quant suplementar
            Set objQuantSupl = New ClassQuantSuplementar

            objQuantSupl.dQuantidade = dQuantComprar
            objQuantSupl.dQuantRequisitada = dQuantComprar
            objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino
            objQuantSupl.iTipoDestino = objReqCompra.iTipoDestino
            objQuantSupl.lFornCliDestino = objReqCompra.lFornCliDestino
                    
            objItemConcorrencia.colQuantSuplementar.Add objQuantSupl
        End If
                
    ' Se não
    Else
        'Cria um novo item de concorrência
        Set objItemConcorrencia = New ClassItemConcorrencia
        
        objItemConcorrencia.iEscolhido = MARCADO
        objItemConcorrencia.iFilial = iFilForn
        objItemConcorrencia.lFornecedor = lForn
        objItemConcorrencia.sProduto = objProduto.sCodigo
        objItemConcorrencia.sDescricao = objProduto.sDescricao
        objItemConcorrencia.sUM = objProduto.sSiglaUMCompra
        objItemConcorrencia.dtDataNecessidade = DATA_NULA
        
        'Cria um registro de quant suplementar p\ o destino da Req do ItemRC
        Set objQuantSupl = New ClassQuantSuplementar
        
        objQuantSupl.dQuantidade = dQuantComprar
        objQuantSupl.dQuantRequisitada = dQuantComprar
        objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino
        objQuantSupl.iTipoDestino = objReqCompra.iTipoDestino
        objQuantSupl.lFornCliDestino = objReqCompra.lFornCliDestino
                
        objItemConcorrencia.colQuantSuplementar.Add objQuantSupl
        
        'Adiciona o novo item de concorrência na coleção global
        gcolItemConcorrencia.Add objItemConcorrencia
        
    End If
        
    If objReqCompra.dtDataLimite <> DATA_NULA Then
        If (objItemConcorrencia.dtDataNecessidade = DATA_NULA) Or (objReqCompra.dtDataLimite < objItemConcorrencia.dtDataNecessidade) Then objItemConcorrencia.dtDataNecessidade = objReqCompra.dtDataLimite
    End If
    
    If objReqCompra.lUrgente = MARCADO Then objItemConcorrencia.dQuantUrgente = objItemConcorrencia.dQuantUrgente + dQuantComprar
    
    'Cria o link entre o item de req e o item de concorrência
    Set objItemRCItemConc = New ClassItemRCItemConcorrencia
    
    objItemRCItemConc.dQuantidade = dQuantComprar
    objItemRCItemConc.lItemReqCompra = objItemRC.lNumIntDoc
    
    objItemConcorrencia.colItemRCItemConcorrencia.Add objItemRCItemConc
    
    'Atualiza a quantidade do item de concorrência
    objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade + dQuantComprar

    ItensConcorrencia_Cria_Altera = SUCESSO
    
    Exit Function
    
Erro_ItensConcorrencia_Cria_Altera:

    ItensConcorrencia_Cria_Altera = gErr
    
    Select Case gErr
    
        Case 62775, 62777
        
        Case 62776
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161168)
            
    End Select
    
    Exit Function

End Function

Function ItensConcorrencia_Atualiza(objReqCompra As ClassRequisicaoCompras, objItemRC As ClassItemReqCompras)

Dim lErro As Long
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objItemRCOutros As ClassItemReqCompras
Dim objReqCompraOutras As ClassRequisicaoCompras
Dim iItem As Integer

On Error GoTo Erro_ItensConcorrencia_Atualiza
    
    'Localiza o item de concorrência correspondente
    Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConcorrencia, iItem, objItemRC)
    
    'Se a requisição está sendo desmarcada
    If objReqCompra.iSelecionado = DESMARCADO Then
        'Se o item da requisição está marcado
        If objItemRC.iSelecionado = MARCADO And iItem > 0 Then
            lErro = ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC)
            If lErro <> SUCESSO Then gError 62782
            
        End If
    'se a requisicao está marcada
    Else
        
        If objItemRC.iSelecionado = MARCADO Then
            'Inclui os dados do item de requisicao
            lErro = ItensConcorrencia_Cria_Altera(objItemRC)
            If lErro <> SUCESSO Then gError 62782
                    
        ElseIf iItem > 0 Then
            
            Set objItemConcorrencia = gcolItemConcorrencia(iItem)
            
            lErro = ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC)
            If lErro <> SUCESSO Then gError 62783
        
        End If
    
    End If
    
    Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConcorrencia, iItem, objItemRC)

    If iItem > 0 Then
        'Renova as cotacoes dos itens alterados
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iItem)
        If lErro <> SUCESSO Then gError 62784
    End If
    
    ItensConcorrencia_Atualiza = SUCESSO
        
    Exit Function
    
Erro_ItensConcorrencia_Atualiza:

    ItensConcorrencia_Atualiza = gErr
    
    Select Case gErr
    
        Case 62782, 62783, 62784
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161169)
            
    End Select
    
    Exit Function
    
End Function

Function ItemConcorrencia_Atualiza_Cotacoes(colItemConcorrencia As Collection, iItem As Integer) As Long
'Atualiza as cotações para o item passado

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim bPrecisa_Ler As Boolean
Dim objItemConcorrencia As ClassItemConcorrencia
Dim iTipoTributacao As Integer
Dim lItemMaior As Long
Dim lNumIntItem As Long
Dim objCotItemConc As ClassCotacaoItemConc
Dim objItemRC As New ClassItemReqCompras
Dim objReqCompra As New ClassRequisicaoCompras
Dim iIndice As Integer

On Error GoTo Erro_ItemConcorrencia_Atualiza_Cotacoes

    bPrecisa_Ler = True

    'recolhe o Item de concorrência
    Set objItemConcorrencia = gcolItemConcorrencia(iItem)
    
    lItemMaior = 1
    
    If objItemConcorrencia.colItemRCItemConcorrencia.Count > 0 Then
        lNumIntItem = objItemConcorrencia.colItemRCItemConcorrencia(1).lItemReqCompra
    
        For iIndice = 1 To objItemConcorrencia.colItemRCItemConcorrencia.Count
            If objItemConcorrencia.colItemRCItemConcorrencia(iIndice).dQuantidade > objItemConcorrencia.colItemRCItemConcorrencia(lItemMaior).dQuantidade Then
                lItemMaior = iIndice
                lNumIntItem = objItemConcorrencia.colItemRCItemConcorrencia(iIndice).lItemReqCompra
            End If
        Next
    End If
    'Lê o Produto
    objProduto.sCodigo = objItemConcorrencia.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 23080 Then gError 62791
    If lErro <> SUCESSO Then gError 62792

    If objProduto.iConsideraQuantCotAnt <> PRODUTO_CONSIDERA_QUANT_COTACAO_ANTERIOR And _
       objItemConcorrencia.colCotacaoItemConc.Count > 0 Then bPrecisa_Ler = False

    If bPrecisa_Ler Then
        
        Set objItemConcorrencia.colCotacaoItemConc = New Collection
                
        lErro = CF("Cotacoes_Produto_Le", objItemConcorrencia.colCotacaoItemConc, objProduto, objItemConcorrencia.dQuantidade, gobjGeracaoPedCompraCot.iTipoDestino, gobjGeracaoPedCompraCot.lFornCliDestino, gobjGeracaoPedCompraCot.iFilialDestino, objItemConcorrencia.lFornecedor, objItemConcorrencia.iFilial)
        If lErro <> SUCESSO And lErro <> 63822 Then gError 62793
        
    Else
    
        For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
            objCotItemConc.dQuantidadeComprar = objItemConcorrencia.dQuantidade
        Next
    
    End If
        
    Call Escolher_Cotacoes(objItemConcorrencia)
    
    If lNumIntItem > 0 Then
        Call Localiza_ItemReqCompra(gcolRequisicaoCompra, lNumIntItem, objItemRC, objReqCompra)
        
        If objItemConcorrencia.colCotacaoItemConc.Count > 0 Then
            For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
                objCotItemConc.iTipoTributacao = objItemRC.iTipoTributacao
            Next
        End If
    End If
    
    ItemConcorrencia_Atualiza_Cotacoes = SUCESSO

    Exit Function
    
Erro_ItemConcorrencia_Atualiza_Cotacoes:

    ItemConcorrencia_Atualiza_Cotacoes = gErr
    
    Select Case gErr
    
        Case 62791, 62793
        
        Case 62792
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161170)
    
    End Select
    
    Exit Function

End Function

Private Sub Localiza_ItemConcorrencia(colItemConcorrencia As Collection, objItemConcorrencia As ClassItemConcorrencia, iItem As Integer, objItemRC As ClassItemReqCompras)
'Devolve os dados do item de concorrecia ligado ao ItemRc passado

Dim objItemConcAux As ClassItemConcorrencia
Dim lForn As Long, iFilForn As Integer
Dim iIndice As Integer, bAchou As Boolean

    iItem = 0
    iIndice = 0
    bAchou = False

    'Busca nos itens de concorrencia
    For Each objItemConcAux In colItemConcorrencia
        iIndice = iIndice + 1
        'Se for exclusivo
        If objItemRC.iExclusivo = MARCADO Then
            lForn = objItemRC.lFornecedor
            iFilForn = objItemRC.iFilial
        Else
            lForn = 0
            iFilForn = 0
        End If
        If objItemConcAux.sProduto = objItemRC.sProduto And objItemConcAux.lFornecedor = lForn And objItemConcAux.iFilial = iFilForn Then
           Set objItemConcorrencia = objItemConcAux
           'encontrou
           bAchou = True
           Exit For
        End If
    Next

    If bAchou Then iItem = iIndice

    Exit Sub

End Sub
Private Sub Adiciona_Codigo(colIndices As Collection, iItem As Integer)
'se o código passado não estiver na coleção ele é adiconado
Dim iIndice As Integer

    For iIndice = 1 To colIndices.Count
        If colIndices(iIndice) = iItem Then Exit Sub
    Next
        
    colIndices.Add iItem

    Exit Sub
    
End Sub

Private Function GridProdutos2_Preenche() As Long
'Preenche o grid de produtos 2

Dim objItemConc As ClassItemConcorrencia
Dim objQuantSupl As ClassQuantSuplementar
Dim iLinha2 As Integer, lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iLinha1 As Integer
Dim colFilEmp As New AdmCollCodigoNome
Dim colFilForn As New Collection
Dim colForn As New AdmCollCodigoNome
Dim objCodNome As AdmlCodigoNome
Dim objFilEmp As New AdmFiliais
Dim iPosicao As Integer

On Error GoTo Erro_GridProdutos2_Preenche
    
    'Limpa o grid de produtos2
    Call Grid_Limpa(objGridProdutos2)
    
    iLinha1 = 0
    iLinha2 = 0
    
    'Para cada item de conc
    For Each objItemConc In gcolItemConcorrencia
        iLinha1 = iLinha1 + 1
        If objItemConc.iEscolhido = MARCADO Then
            
            'Para cada quant supl
            For Each objQuantSupl In objItemConc.colQuantSuplementar
            
                iLinha2 = iLinha2 + 1
                'Preenche com os dados do item de conorrência
                GridProdutos2.TextMatrix(iLinha2, iGrid_Produto2_Col) = GridProdutos1.TextMatrix(iLinha1, iGrid_Produto1_Col)
                GridProdutos2.TextMatrix(iLinha2, iGrid_DescProduto2_Col) = objItemConc.sDescricao
                GridProdutos2.TextMatrix(iLinha2, iGrid_UnidadeMed2_Col) = objItemConc.sUM
                GridProdutos2.TextMatrix(iLinha2, iGrid_Quantidade2_Col) = Formata_Estoque(objQuantSupl.dQuantidade)
                  
                If objQuantSupl.iTipoDestino = TIPO_DESTINO_EMPRESA Then
                    
                    Call Busca_Na_Colecao(colFilEmp, objQuantSupl.iFilialDestino, iPosicao)
                    
                    If iPosicao = 0 Then
                    
                        objFilEmp.lCodEmpresa = glEmpresa
                        objFilEmp.iCodFilial = objQuantSupl.iFilialDestino
                                                                
                        lErro = CF("FilialEmpresa_Le", objFilEmp, True)
                        If lErro <> SUCESSO And lErro <> 27378 Then gError 62788
                        If lErro <> SUCESSO Then gError 62789
                        
                        Set objCodNome = New AdmlCodigoNome
                        
                        objCodNome.sNome = objFilEmp.sNome
                        objCodNome.lCodigo = objFilEmp.iCodFilial
                        
                        colFilEmp.Add objCodNome.lCodigo, objCodNome.sNome
                    
                    Else
                        Set objCodNome = colFilEmp(iPosicao)
                    End If
                    'Preenche os dados do destino
                    GridProdutos2.TextMatrix(iLinha2, iGrid_TipoDestino_Col) = "Empresa"
                    GridProdutos2.TextMatrix(iLinha2, iGrid_Destino_Col) = ""
                  
                    GridProdutos2.TextMatrix(iLinha2, iGrid_FilialDestino_Col) = objCodNome.lCodigo & SEPARADOR & objCodNome.sNome
                  
                ElseIf objQuantSupl.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then
                    
                    GridProdutos2.TextMatrix(iLinha2, iGrid_TipoDestino_Col) = "Fornecedor"
                                          
                    Call Busca_Na_Colecao(colForn, objQuantSupl.lFornCliDestino, iPosicao)
                                        
                    If iPosicao = 0 Then
                        objFornecedor.lCodigo = objQuantSupl.lFornCliDestino
                        
                        'Lê o fornecedor
                        lErro = CF("Fornecedor_Le", objFornecedor)
                        If lErro <> SUCESSO And lErro <> 12729 Then gError 62790
                        If lErro <> SUCESSO Then gError 62791
                                            
                        Set objCodNome = New AdmlCodigoNome
                        
                        objCodNome.lCodigo = objFornecedor.lCodigo
                        objCodNome.sNome = objFornecedor.sNomeReduzido
                    
                        colForn.Add objCodNome.lCodigo, objCodNome.sNome
                    Else
                        Set objCodNome = colForn(iPosicao)
                    End If
                    
                    GridProdutos2.TextMatrix(iLinha2, iGrid_Destino_Col) = objCodNome.sNome
                      
                    Call Busca_FilialForn(colFilForn, objQuantSupl.lFornCliDestino, objQuantSupl.iFilialDestino, iPosicao)
                    
                    If iPosicao = 0 Then
                        Set objFilialFornecedor = New ClassFilialFornecedor
                        
                        objFilialFornecedor.lCodFornecedor = objQuantSupl.lFornCliDestino
                        objFilialFornecedor.iCodFilial = objQuantSupl.iFilialDestino
                    
                        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                        If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
                    
                        'Se não encontrou==>Erro
                        If lErro = 12929 Then gError 63990
                                         
                        colFilForn.Add objFilialFornecedor
                    Else
                        Set objFilialFornecedor = colFilForn(iPosicao)
                    End If
                    'Preenche os dados do destino
                    GridProdutos2.TextMatrix(iLinha2, iGrid_FilialDestino_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
                  
                End If
                
                GridProdutos2.TextMatrix(iLinha2, iGrid_Fornecedor2_Col) = GridProdutos1.TextMatrix(iLinha1, iGrid_Fornecedor1_Col)
                GridProdutos2.TextMatrix(iLinha2, iGrid_FilialForn2_Col) = GridProdutos1.TextMatrix(iLinha1, iGrid_FilialForn1_Col)
                
            Next
        End If
    Next
    
    objGridProdutos2.iLinhasExistentes = iLinha2

    Call Grid_Refresh_Checkbox(objGridProdutos2)
    
    GridProdutos2_Preenche = SUCESSO
    
    Exit Function
    
Erro_GridProdutos2_Preenche:

    GridProdutos2_Preenche = gErr
    
    Select Case gErr
    
        Case 62788, 62790, 63989

        Case 62789
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilEmp.iCodFilial)
        
        Case 62791
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63990
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFornecedor.lCodigo)

    End Select
    
    Exit Function
    
End Function

Function ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia As ClassItemConcorrencia, iItem As Integer, Optional objReqCompra As ClassRequisicaoCompras, Optional objItemRC As ClassItemReqCompras, Optional dQuantidade As Double = 0)

Dim iIndice As Integer
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objQtSupl As ClassQuantSuplementar
Dim lErro As Long
Dim bExclui As Boolean
Dim objProduto As New ClassProduto
Dim dFator As Double
    
On Error GoTo Erro_ItemConcorrencia_Exclui_QuantComprar
    
    'Se a quantidade não foi passada
    If dQuantidade = 0 Then
        
        objProduto.sCodigo = objItemRC.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 23080 Then gError 62785
        If lErro <> SUCESSO Then gError 62786
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 62787
        
        'A quantidade a exclui é a do ItemRC passado
        dQuantidade = objItemRC.dQuantComprar * dFator
        'Exclui a ligação do item RC com o item conc
        bExclui = True
    End If

    'diminui a quantidade a comprar do item de concorrencia vinculado
    objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade - dQuantidade
    
       
    iIndice = 0
    
    'Se algum item de req foi passado
    If Not (objItemRC Is Nothing) Then
        'Exclui o vinculo entre o item de requisicao e o item de concorrencia
        For Each objItemRCItemConc In objItemConcorrencia.colItemRCItemConcorrencia
            iIndice = iIndice + 1
            'BUsca o vinculo do ItemRc e ItemConc
            If objItemRCItemConc.lItemReqCompra = objItemRC.lNumIntDoc Then
                'Se a quant do item foi toda excluída
                If bExclui Then
                    'exclui o link entre o item RC e o item conc
                    objItemConcorrencia.colItemRCItemConcorrencia.Remove iIndice
                'senão
                Else
                    'Diminui a quantidade excluída
                    objItemRCItemConc.dQuantidade = objItemRCItemConc.dQuantidade - dQuantidade
                End If
                            
                Exit For
            End If
        Next

        iIndice = 0
        'Diminui a quantidade a comprar do correspondente em quant suplementares
        For Each objQtSupl In objItemConcorrencia.colQuantSuplementar
            iIndice = iIndice + 1
            If objQtSupl.iTipoDestino = objReqCompra.iTipoDestino And objQtSupl.iFilialDestino = objReqCompra.iFilialDestino And objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino Then
                objQtSupl.dQuantidade = objQtSupl.dQuantidade - dQuantidade
                objQtSupl.dQuantRequisitada = objQtSupl.dQuantRequisitada - dQuantidade
                If objQtSupl.dQuantidade <= 0 Then objItemConcorrencia.colQuantSuplementar.Remove iIndice
                Exit For
            End If
        Next
        If objReqCompra.lUrgente = MARCADO Then objItemConcorrencia.dQuantUrgente = objItemConcorrencia.dQuantUrgente - dQuantidade
    End If
        
    'Se o item de concorrencia não está vinculado a nenum outro itemRC
    If (objItemConcorrencia.colItemRCItemConcorrencia.Count = 0) And gcolRequisicaoCompra.Count > 0 Then
        'Exclui o item de concorrência
        gcolItemConcorrencia.Remove iItem
    Else
        
        If iItem = 0 Then
            'Altera os dados de compra dos itens de concorr6encia
            '(inclusive cotações, se necessário)
            lErro = ItensConcorrencia_Cria_Altera(objItemRC)
            If lErro <> SUCESSO Then gError 62739
        End If
                
    End If
    
    ItemConcorrencia_Exclui_QuantComprar = SUCESSO
    
    Exit Function

Erro_ItemConcorrencia_Exclui_QuantComprar:

    ItemConcorrencia_Exclui_QuantComprar = gErr
    
    Select Case gErr
    
        Case 62739, 62785, 62787
        
        Case 62786
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161171)
            
    End Select
    
    Exit Function

End Function

Function ItemConcorrencia_Inclui_QuantComprar(objItemConcorrencia As ClassItemConcorrencia, iItem As Integer, Optional objReqCompra As ClassRequisicaoCompras, Optional objItemRC As ClassItemReqCompras, Optional dQuantidade As Double)

Dim iIndice As Integer
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objQtSupl As ClassQuantSuplementar
Dim lErro As Long
Dim bAchou As Boolean

On Error GoTo Erro_ItemConcorrencia_Inclui_QuantComprar

    'Se o item já foi passado atualizado
    If iItem > 0 Then

        'diminui a quantidade a comprar do item de concorrencia vinculado
        objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade + dQuantidade
        
        If Not (objItemRC Is Nothing) Then
            iIndice = 0
            'Atualiza o vinculo entre o item de requisicao e o item de concorrencia
            For Each objItemRCItemConc In objItemConcorrencia.colItemRCItemConcorrencia
                iIndice = iIndice + 1
                If objItemRCItemConc.lItemReqCompra = objItemRC.lNumIntDoc Then
                    objItemRCItemConc.dQuantidade = objItemRCItemConc.dQuantidade + dQuantidade
                    Exit For
                End If
            Next
            
            iIndice = 0
            'Aumenta a quantidade a comprar do correspondente em quant suplementares
            For Each objQtSupl In objItemConcorrencia.colQuantSuplementar
                iIndice = iIndice + 1
                If objQtSupl.iTipoDestino = objReqCompra.iTipoDestino And objQtSupl.iFilialDestino = objReqCompra.iFilialDestino And objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino Then
                    bAchou = True
                    objQtSupl.dQuantidade = objQtSupl.dQuantidade + dQuantidade
                    objQtSupl.dQuantRequisitada = objQtSupl.dQuantRequisitada + dQuantidade
                    Exit For
                End If
            Next
            'Se não há quant suplementar p\ esse destino
            If Not bAchou Then
                'Cria um registro de quant siplementar novo
                Set objQtSupl = New ClassQuantSuplementar
                
                objQtSupl.dQuantidade = dQuantidade
                objQtSupl.dQuantRequisitada = dQuantidade
                objQtSupl.iFilialDestino = objReqCompra.iFilialDestino
                objQtSupl.iTipoDestino = objReqCompra.iTipoDestino
                objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino
                
                If objReqCompra.lUrgente = MARCADO Then objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade + dQuantidade
            
                objItemConcorrencia.colQuantSuplementar.Add objQtSupl
            End If
        
        End If
    
    Else
        
        lErro = ItensConcorrencia_Cria_Altera(objItemRC)
        If lErro <> SUCESSO Then gError 62739
    End If
                    
    ItemConcorrencia_Inclui_QuantComprar = SUCESSO
    
    Exit Function
    
Erro_ItemConcorrencia_Inclui_QuantComprar:

    ItemConcorrencia_Inclui_QuantComprar = gErr
    
    Select Case gErr
        
        Case 62739
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161172)
        
    End Select

    Exit Function

End Function

Function Atualiza_QuantSupl(objItemConcorrencia As ClassItemConcorrencia, dQuantDiferenca As Double, iLinhaProd2 As Integer)
'Atualiza a coleçao de quantidades suplementares

Dim lErro As Long
Dim objQuantSupl As ClassQuantSuplementar
Dim lForn As Long
Dim iFilial As Integer
Dim iTipo As Integer
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Atualiza_QuantSupl

    lForn = 0
    iFilial = Codigo_Extrai(GridProdutos2.TextMatrix(iLinhaProd2, iGrid_FilialDestino_Col))
    
    'Recolhe o tipo de destino
    If GridProdutos2.TextMatrix(iLinhaProd2, iGrid_TipoDestino_Col) = "Empresa" Then
        iTipo = TIPO_DESTINO_EMPRESA
    Else
        iTipo = TIPO_DESTINO_FORNECEDOR
        
        Set objFornecedor = New ClassFornecedor
        
        objFornecedor.sNomeReduzido = GridProdutos2.TextMatrix(iLinhaProd2, iGrid_Destino_Col)
        'Lê o fornecdor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 62773
        If lErro <> SUCESSO Then gError 62774
        
        lForn = objFornecedor.lCodigo
    End If
    
    'Localiza o registro de quant supl correspondente
    For Each objQuantSupl In objItemConcorrencia.colQuantSuplementar
        
        If objQuantSupl.iFilialDestino = iFilial And objQuantSupl.lFornCliDestino = lForn And objQuantSupl.iTipoDestino = iTipo Then
            'Atualiza a quantidade
            If (objQuantSupl.dQuantidade + dQuantDiferenca) < objQuantSupl.dQuantRequisitada Then gError 62772
            objQuantSupl.dQuantidade = objQuantSupl.dQuantidade + dQuantDiferenca
        End If
    Next

    Atualiza_QuantSupl = SUCESSO

    Exit Function

Erro_Atualiza_QuantSupl:

    Atualiza_QuantSupl = gErr

    Select Case gErr
        
        Case 62772
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_MENOR_QUANTCOMPRAR_RC", gErr, (objQuantSupl.dQuantidade + dQuantDiferenca), objQuantSupl.dQuantRequisitada)
        
        Case 62773
        
        Case 62774
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161173)
    
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
    
    'Recolhe os campos que amarram  uma cotação na tela
    sMotivo = GridCotacoes.TextMatrix(iLinha, iGrid_MotivoEscolhaCot_Col)
    sProduto = GridCotacoes.TextMatrix(iLinha, iGrid_ProdutoCot_Col)
    sCondPagto = GridCotacoes.TextMatrix(iLinha, iGrid_CondPagtoCot_Col)
    sFornecedor = GridCotacoes.TextMatrix(iLinha, iGrid_FornecedorCot_Col)
    sFilial = GridCotacoes.TextMatrix(iLinha, iGrid_FilialFornCot_Col)
    
    For iIndice = 0 To Moeda.ListCount - 1
        If Moeda.List(iIndice) = GridCotacoes.TextMatrix(iLinha, iGrid_Moeda_Col) Then
            iMoeda = Moeda.ItemData(iIndice)
            Exit For
        End If
    Next
    
    'Se for exclusivo
    If sMotivo = MOTIVO_EXCLUSIVO_DESCRICAO Then
        
        'Para cada item de concorrencia
        For iIndice = 1 To objGridProdutos1.iLinhasExistentes
            
            'Busca o item com forn e filial amarrados
            If GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) = sProduto And _
               GridProdutos1.TextMatrix(iIndice, iGrid_Fornecedor1_Col) = sFornecedor And _
               GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col) = sFilial Then
                
                iItemConc = iIndice
        
            End If
        
        Next
        
    Else
        
        For iIndice = 1 To objGridProdutos1.iLinhasExistentes
            'Busca o item de concorrência ligado a cotação
            If GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) = sProduto And _
                Len(Trim(GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col))) = 0 Then
                iItemConc = iIndice
            End If
        Next
    End If
    
    'Seleciona o item de concorrência
    Set objItemConcorrencia = gcolItemConcorrencia(iItemConc)
    
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

Private Sub Calcula_TotalItens()
'Calcula o valor total dos itens selecionados

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
        objCotItemConc.iSelecionada = MARCADO
        If objCotItemConc.iEscolhido = MARCADO Then
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161174)

    End Select

    Exit Function
    
End Function

Private Sub Inicializa_QuantAssocia_ItenRC(colRequisicao As Collection)
'Zera o campo QuantNoPedido dos Itens de Requisição

Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras

    For Each objReqCompra In colRequisicao
        For Each objItemRC In objReqCompra.colItens
            objItemRC.dQuantNoPedido = 0
        Next
    Next
    
    Exit Sub

End Sub

Private Sub Transfere_Dados_Cotacoes(colCotacaoItemConc As Collection, colCotItemConcAux As Collection)

Dim objCotItemConc As ClassCotacaoItemConc
Dim objCotItemConcAux As ClassCotacaoItemConc

    Set colCotItemConcAux = New Collection
    
    For Each objCotItemConc In colCotacaoItemConc
        
        If objCotItemConc.iEscolhido = MARCADO Then
        
            Set objCotItemConcAux = New ClassCotacaoItemConc
            
            objCotItemConcAux.dAliquotaICMS = objCotItemConc.dAliquotaICMS
            objCotItemConcAux.dAliquotaIPI = objCotItemConc.dAliquotaIPI
            objCotItemConcAux.dCreditoICMS = objCotItemConc.dCreditoICMS
            objCotItemConcAux.dCreditoIPI = objCotItemConc.dCreditoIPI
            objCotItemConcAux.dPrecoAjustado = objCotItemConc.dPrecoAjustado
            objCotItemConcAux.dPrecoUnitario = objCotItemConc.dPrecoUnitario
            objCotItemConcAux.dPreferencia = objCotItemConc.dPreferencia
            objCotItemConcAux.dQuantCotada = objCotItemConc.dQuantCotada
            objCotItemConcAux.dQuantEntrega = objCotItemConc.dQuantEntrega
            objCotItemConcAux.dQuantidadeComprar = objCotItemConc.dQuantidadeComprar
            objCotItemConcAux.dtDataEntrega = objCotItemConc.dtDataEntrega
            objCotItemConcAux.dtDataValidade = objCotItemConc.dtDataValidade
            objCotItemConcAux.dValorPresente = objCotItemConc.dValorPresente
            objCotItemConcAux.iEscolhido = objCotItemConc.iEscolhido
            objCotItemConcAux.iPrazoEntrega = objCotItemConc.iPrazoEntrega
            objCotItemConcAux.iSelecionada = objCotItemConc.iSelecionada
            objCotItemConcAux.lItemCotacao = objCotItemConc.lItemCotacao
            objCotItemConcAux.lNumIntDoc = objCotItemConc.lNumIntDoc
            objCotItemConcAux.lPedCotacao = objCotItemConc.lPedCotacao
            objCotItemConcAux.sCondPagto = objCotItemConc.sCondPagto
            objCotItemConcAux.sFilial = objCotItemConc.sFilial
            objCotItemConcAux.sFornecedor = objCotItemConc.sFornecedor
            objCotItemConcAux.sMotivoEscolha = objCotItemConc.sMotivoEscolha
            objCotItemConcAux.sUMCompra = objCotItemConc.sUMCompra
            objCotItemConcAux.iMoeda = objCotItemConc.iMoeda
            objCotItemConcAux.dTaxa = objCotItemConc.dTaxa
            
            colCotItemConcAux.Add objCotItemConcAux
        End If
    Next

    Exit Sub
    
End Sub

Function Inclui_Quant_ItemReqCompra(objItemPC As ClassItemPedCompra, objItemConcorrencia As ClassItemConcorrencia, objQuantSupl As ClassQuantSuplementar, colRequisicao As Collection, colProdutos As Collection)

Dim lErro As Long
Dim dQuantidade As Double
Dim objItemReqCompra As ClassItemReqCompras
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim dDiferenca As Double
Dim objItemRC As ClassItemReqCompras
Dim objReqCompra As New ClassRequisicaoCompras
Dim objLocItemPC As ClassLocalizacaoItemPC
Dim bAchou As Boolean, dFatorCOM As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Inclui_Quant_ItemReqCompra

    Call Busca_Produto(objItemPC.sProduto, colProdutos, objProduto, bAchou)

    If Not bAchou Then
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objItemPC.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 86147
        If lErro <> SUCESSO Then gError 86149
    
        colProdutos.Add objProduto
    
    End If
    
    dQuantidade = objItemPC.dQuantidade

    'Para cada item de req que gerou esse item de concorrência
    For Each objItemRCItemConc In objItemConcorrencia.colItemRCItemConcorrencia

        'Busca os dados do item
        Call Localiza_ItemReqCompra(colRequisicao, objItemRCItemConc.lItemReqCompra, objItemReqCompra, objReqCompra)

        'Se o item acessado é do mesmo tipo de destino do PC
        If objReqCompra.iTipoDestino = objQuantSupl.iTipoDestino And objReqCompra.lFornCliDestino = objQuantSupl.lFornCliDestino And objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino And (objItemReqCompra.dQuantComprar - objItemReqCompra.dQuantNoPedido > 0) Then
                    
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemReqCompra.sUM, objItemPC.sUM, dFatorCOM)
            If lErro <> SUCESSO Then gError 86148
            
            'Calcula a diferença entre a qt do ItemPC e a qt ñ associada do ItemRC
            dDiferenca = dQuantidade - ((objItemReqCompra.dQuantComprar - objItemReqCompra.dQuantNoPedido) * dFatorCOM)

            'Cria um objItemRC
            Set objItemRC = New ClassItemReqCompras

            'recolhe alguns dados
            objItemRC.lNumIntDoc = objItemReqCompra.lNumIntDoc
            objItemRC.iAlmoxarifado = objItemReqCompra.iAlmoxarifado
            objItemRC.sProduto = objItemReqCompra.sProduto
            objItemRC.sUM = objItemReqCompra.sUM
            objItemRC.sCcl = objItemReqCompra.sCcl
            objItemRC.sDescProduto = objItemReqCompra.sDescProduto
            objItemRC.sContaContabil = objItemReqCompra.sContaContabil

            'se a diferença for positiva
            If dDiferenca >= 0 Then
                'A quantidade do item q não está associada a ItemPC será utilizada
                objItemRC.dQuantComprar = objItemReqCompra.dQuantComprar - objItemReqCompra.dQuantNoPedido
                objItemReqCompra.dQuantNoPedido = objItemReqCompra.dQuantComprar
            'se for negativa
            Else
                'Parte da quantidade do item q não está associada a ItemPC será utilizada
                objItemRC.dQuantComprar = dQuantidade / dFatorCOM
                objItemReqCompra.dQuantNoPedido = objItemReqCompra.dQuantNoPedido + (dQuantidade / dFatorCOM)
            End If

            If objItemRC.iAlmoxarifado > 0 Then

                bAchou = False
                For Each objLocItemPC In objItemPC.colLocalizacao
                    If objLocItemPC.iAlmoxarifado = objItemRC.iAlmoxarifado Then
                        bAchou = True
                        objLocItemPC.dQuantidade = objLocItemPC.dQuantidade + (objItemRC.dQuantComprar * dFatorCOM)
                    End If
                Next

                If Not bAchou Then
                    Set objLocItemPC = New ClassLocalizacaoItemPC

                    objLocItemPC.dQuantidade = (objItemRC.dQuantComprar * dFatorCOM)
                    objLocItemPC.iAlmoxarifado = objItemRC.iAlmoxarifado
                    objLocItemPC.sCcl = objItemRC.sCcl
                    objLocItemPC.sContaContabil = objItemRC.sContaContabil

                    objItemPC.colLocalizacao.Add objLocItemPC
                End If
            End If

            objItemPC.colItemReqCompras.Add objItemRC
            'Atualiza a quantidade que falta associar a ItemPC
            dQuantidade = dQuantidade - (objItemRC.dQuantComprar * dFatorCOM)

            'Se já associou toda a quantidade, sai
            If dQuantidade = 0 Then Exit Function

        End If

    Next

    Inclui_Quant_ItemReqCompra = SUCESSO

    Exit Function

Erro_Inclui_Quant_ItemReqCompra:

    Inclui_Quant_ItemReqCompra = gErr

    Select Case gErr
    
        Case 86147, 86148
        
        Case 86149
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161175)

    End Select

    Exit Function

End Function

Function colItensCotacao_Adiciona(lItemCotacao As Long, colItensCotacao As Collection) As Long
'Se o Item de cotação não existe na coleção ele é lido e incluído

Dim objItemCotacao As ClassItemCotacao
Dim bAchou As Boolean
Dim lErro As Long

On Error GoTo Erro_colItensCotacao_Adiciona

    bAchou = False
    'Busca o Item de cotação
    For Each objItemCotacao In colItensCotacao
        If objItemCotacao.lNumIntDoc = lItemCotacao Then
            bAchou = True
            Exit For
        End If
    Next
    
    If Not bAchou Then
        Set objItemCotacao = New ClassItemCotacao
        
        objItemCotacao.lNumIntDoc = lItemCotacao
        'Lê o Item cotação
        lErro = CF("ItemCotacao_Le", objItemCotacao)
        If lErro <> SUCESSO Then gError 62725
        
        'Adiciona na coleção
        colItensCotacao.Add objItemCotacao, CStr(objItemCotacao.lNumIntDoc)

    End If
    
    colItensCotacao_Adiciona = SUCESSO
    
    Exit Function

Erro_colItensCotacao_Adiciona:

    colItensCotacao_Adiciona = gErr
    
    Select Case gErr
    
        Case 62725
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161176)
    
    End Select

    Exit Function

End Function


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
                If objPCExclu.iTipoDestino = objPCGeral.iTipoDestino And objPCExclu.iFilialDestino = objPCGeral.iFilialDestino And objPCExclu.lFornCliDestino = objPCGeral.lFornCliDestino And objPCExclu.lFornecedor = objPCGeral.lFornecedor And objPCExclu.iFilial = objPCGeral.iFilial And objPCExclu.iCondicaoPagto = objPCGeral.iCondicaoPagto Then
                
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161177)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161178)
            
    End Select
    
    Exit Function

End Function

Private Function Busca_QuantComprar_ItemReq(lReqCompra As Long, iFilialReq As Integer, iItem As Integer, dQuantComprar As Double)

Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim lErro As Long

On Error GoTo Erro_Busca_QuantComprar_ItemReq

    dQuantComprar = 0

    'Para cada Requisição da tela
    For Each objReqCompra In gcolRequisicaoCompra
        'se for a req passada
        If objReqCompra.lCodigo = lReqCompra And objReqCompra.iFilialEmpresa = iFilialReq Then
            'Localiza o item procurado
            For Each objItemRC In objReqCompra.colItens
                If objItemRC.iItem = iItem Then
                    
                    objProduto.sCodigo = objItemRC.sProduto
                    'Lê o produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 62796
                    If lErro <> SUCESSO Then gError 62797
                    
                    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
                    If lErro <> SUCESSO Then gError 62798
                    
                    'COnverte para a UM compra
                    dQuantComprar = objItemRC.dQuantComprar * dFator
                    Exit For
                End If
            Next
        End If
        
    Next
    
    Busca_QuantComprar_ItemReq = SUCESSO

    Exit Function

Erro_Busca_QuantComprar_ItemReq:

    Busca_QuantComprar_ItemReq = gErr
    
    Select Case gErr
    
        Case 62796, 62798
        
        Case 62797
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161179)
            
    End Select

    Exit Function

End Function

Sub Monta_Colecao_Campos_Cotacao(colCampos As Collection, iOrdenacao As Integer)
'monta a coleção de campos para a ordenação

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

Private Sub Busca_Na_Colecao(collCodigoNome As AdmCollCodigoNome, lCodigo As Long, iPosicao As Integer)
'Busca a chave lCodigo na coleção

Dim objlCodigoNome As AdmlCodigoNome
Dim iIndice As Integer

    iPosicao = 0
    iIndice = 0
    
    'Para cada item da coleção
    For Each objlCodigoNome In collCodigoNome
        
        iIndice = iIndice + 1
        
        'Busca o item com a chave passada
        If objlCodigoNome.lCodigo = lCodigo Then
            
            iPosicao = iIndice
            Exit For
        
        End If
    
    Next
    
    Exit Sub

End Sub


Private Sub Busca_FilialForn(colFilialForn As Collection, lFornecedor As Long, iFilial As Integer, iPosicao As Integer)

Dim objFilialFornecedor As ClassFilialFornecedor
Dim iIndice As Integer

    iPosicao = 0
    
    For iIndice = 1 To colFilialForn.Count
        
        Set objFilialFornecedor = colFilialForn(iIndice)
        If objFilialFornecedor.lCodFornecedor = lFornecedor And objFilialFornecedor.iCodFilial = iFilial Then
            iPosicao = iIndice
            Exit Sub
        End If
    Next
        
    Exit Sub
    
End Sub

Private Function Requisicoes_Atualiza() As Long
    
Dim objRequisicao As New ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim lErro As Long
    
On Error GoTo Erro_Requisicoes_Atualiza
    
    'Se a Requisição foi selecionada
    If objGridRequisicoes.objGrid.Col = iGrid_EscolhidoReq_Col And objGridRequisicoes.iLinhasExistentes > 0 Then
               
        Set objRequisicao = gcolRequisicaoCompra(GridRequisicoes.Row)
        
        'Atualiza o campo selecionado na requisicao
        objRequisicao.iSelecionado = GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_EscolhidoReq_Col)
        
        'Para cada Item
        For Each objItemRC In objRequisicao.colItens
        
            If objRequisicao.iSelecionado = MARCADO Then
                If objItemRC.iSelecionado = DESMARCADO Then
                    objItemRC.iSelecionado = MARCADO
                    objItemRC.dQuantComprar = objItemRC.dQuantidade - objItemRC.dQuantCancelada - objItemRC.dQuantPedida - objItemRC.dQuantRecebida
                End If
            End If
            
            'Atualiza os dados do item de concorrência vinculado ao ItemRC
            lErro = ItensConcorrencia_Atualiza(objRequisicao, objItemRC)
            If lErro <> SUCESSO Then gError 62750
        
        Next
        
        'Preenche o grid de itens de requisição
        lErro = GridItensReq_Preenche()
        If lErro <> SUCESSO Then gError 62751
        
        'Preenche o grid de produtos e cotações
        lErro = Grids_Produto_Preenche()
        If lErro <> SUCESSO Then gError 62742

    End If
    
    Requisicoes_Atualiza = SUCESSO
    
    Exit Function
    
Erro_Requisicoes_Atualiza:

    Requisicoes_Atualiza = gErr
    
    Select Case gErr
    
        Case 62742, 62750, 62751
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161180)
    
    End Select

    Exit Function

End Function


Function Atualiza_ItensReq() As Long

Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim objItemRC As ClassItemReqCompras
Dim objReqCompra As ClassRequisicaoCompras
Dim lErro As Long, bAchou As Boolean

On Error GoTo Erro_Atualiza_ItensReq

    'Busca o ItemRc e a Requisição correspondente a linha clicada
    For iIndice1 = 1 To gcolRequisicaoCompra.Count
        
        Set objReqCompra = gcolRequisicaoCompra(iIndice1)
        For iIndice2 = 1 To objReqCompra.colItens.Count
            
            Set objItemRC = objReqCompra.colItens(iIndice2)
            
            If objItemRC.iItem = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_Item_Col)) And _
               objReqCompra.lCodigo = StrParaLong(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_CodigoReqItem_Col)) And _
               objReqCompra.iFilialEmpresa = Codigo_Extrai(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_FilialReqItem_Col)) Then
                'Encontrou
                bAchou = True
                Exit For
            End If
        Next
        'Se já achou sai
        If bAchou Then Exit For
    Next
    
    If objItemRC.iSelecionado = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_EscolhidoItem_Col)) Then Exit Function
    
    'Atualiza a seleção do Item RC
    objItemRC.iSelecionado = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_EscolhidoItem_Col))

    'Atualiza os dados do item de concorrência ligado ao item RC
    lErro = ItensConcorrencia_Atualiza(objReqCompra, objItemRC)
    If lErro <> SUCESSO Then gError 62743
    
    'Preenche o grid de produtos
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62742

    Atualiza_ItensReq = SUCESSO
    
    Exit Function

Erro_Atualiza_ItensReq:

    Atualiza_ItensReq = gErr
    
    Select Case gErr

        Case 62742, 62743

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161181)

    End Select

    Exit Function

End Function


Private Sub Obtem_ReqCompra(colRequisicao As Collection, lNumIntReq As Long, objReqCompra As ClassRequisicaoCompras)
'Devolve os dados da Requisição de compras do Item de Requisição de compras passado

Dim objRequisicao As ClassRequisicaoCompras

    'Busca a Requisicao de compras
    For Each objRequisicao In colRequisicao
        'Se é a Requisição procurada
        If objRequisicao.lNumIntDoc = lNumIntReq Then
            'Guarda a requisição
            Set objReqCompra = objRequisicao
            'Sai da função
            Exit For
        End If
    Next

    Exit Sub

End Sub

Private Sub Escolher_Cotacoes(objItemConcorrencia As ClassItemConcorrencia)
'recebe a coleção de Itens de cotação lida do BD e Escolhe para
'o usuário aquelas que possuem melhor preço ,ou melhor preco + prazo entrega
'como defaut
Dim dMelhorPreco As Double
Dim objCotItemConcMelhor As ClassCotacaoItemConc
Dim objCotItemConc As ClassCotacaoItemConc
Dim dValorPresente As Double
Dim lErro As Long
Dim dTaxa As Double
Dim dValorPresenteReal As Double
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim iIndice As Integer
Dim objCondicaoPagto As ClassCondicaoPagto, dDias As Double

On Error GoTo Erro_Escolher_Cotacoes
    
    dMelhorPreco = 0
      
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
    
    If objCondicaoPagto.colParcelas.Count > 0 Then objCotItemConcMelhor.dtDataVencPriParc = objCondicaoPagto.colParcelas.Item(1).dtVencimento
    dDias = 0
    For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
        dDias = dDias + DateDiff("d", gdtDataAtual, objCondicaoPagto.colParcelas.Item(iIndice).dtVencimento)
    Next
    objCotItemConcMelhor.dPrazoMedio = dDias / objCondicaoPagto.iNumeroParcelas
    
    'Para cada cotação do item
    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
        
        Set objCondicaoPagto = New ClassCondicaoPagto
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConc.sCondPagto)
        
        'Recalcula o Valor Presente
        lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
        If lErro <> SUCESSO Then gError 62733
        
        If objCondicaoPagto.colParcelas.Count > 0 Then objCotItemConc.dtDataVencPriParc = objCondicaoPagto.colParcelas.Item(1).dtVencimento
        dDias = 0
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
            dDias = dDias + DateDiff("d", gdtDataAtual, objCondicaoPagto.colParcelas.Item(iIndice).dtVencimento)
        Next
        objCotItemConc.dPrazoMedio = dDias / objCondicaoPagto.iNumeroParcelas

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
                        
                        'dMelhorPreco = objCotItemConc.dValorPresente
                        Set objCotItemConcMelhor = objCotItemConc
                        objCotItemConcMelhor.sMotivoEscolha = MOTIVO_PRECO_PRAZO_DESCRICAO
                        objCotItemConcMelhor.iEscolhido = MARCADO
                        objCotItemConcMelhor.iSelecionada = MARCADO
                    End If
                End If
                
                If objCotItemConc.iPrazoEntrega = objCotItemConcMelhor.iPrazoEntrega Then
                    
                    'If objCotItemConc.dtDataVencPriParc > objCotItemConcMelhor.dtDataVencPriParc Then
                    If objCotItemConc.dPrazoMedio < objCotItemConcMelhor.dPrazoMedio Then
                    
                        objCotItemConcMelhor.sMotivoEscolha = ""
                        objCotItemConcMelhor.iEscolhido = DESMARCADO
                        objCotItemConcMelhor.iSelecionada = DESMARCADO
                        
                        'Guarda essa cotação como a de melhor preço
                        dMelhorPreco = dValorPresenteReal
                        
                        Set objCotItemConcMelhor = objCotItemConc
                        
                        objCotItemConcMelhor.sMotivoEscolha = MOTIVO_MELHORPRECO_DESCRICAO
                        objCotItemConcMelhor.iEscolhido = MARCADO
                        objCotItemConcMelhor.iSelecionada = MARCADO
                    
                    End If
                    
                End If
                
                
            End If
        End If
    Next
    
    Exit Sub
    
Erro_Escolher_Cotacoes:

    Select Case gErr
    
        Case 62733
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161182)
            
    End Select
        
    Exit Sub
        
End Sub

Private Sub Localiza_ItemReqCompra(colRequisicao As Collection, lItemReqCompra As Long, objItemReqCompra As ClassItemReqCompras, objReqCompra As ClassRequisicaoCompras)
'Localiza o Item de Requisicao com o numero interno passado

Dim iIndice As Integer
Dim objItemRC As ClassItemReqCompras
    
    'Para cada Requsiicao
    For iIndice = 1 To colRequisicao.Count
        Set objReqCompra = colRequisicao(iIndice)
        'Para cada item
        For Each objItemRC In objReqCompra.colItens
            'Se for o item procurado
            If objItemRC.lNumIntDoc = lItemReqCompra Then
                'Devolve o item encontrado
                Set objItemReqCompra = objItemRC
                'Sai a funcao
                Exit Sub
            End If
        Next
    Next

    Exit Sub

End Sub

Private Function Traz_Cotacao_Tela() As Long

Dim lErro As Long
Dim objCotacao As New ClassCotacao
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemRCCot As ClassItemReqCompras
Dim objItemReq As ClassItemReqCompras
Dim objReqCompras As ClassRequisicaoCompras
Dim bAchou As Boolean

On Error GoTo Erro_Traz_Cotacao_Tela

    Call Grid_Limpa(objGridRequisicoes)
    Call Grid_Limpa(objGridItensRequisicoes)
    Call Grid_Limpa(objGridProdutos1)
    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)
    
    Set gcolRequisicaoCompra = New Collection
    Set gcolItemConcorrencia = New Collection

    If gobjGeracaoPedCompraCot.iCotacaoSel > 0 Then

        Set objCotacao = gobjGeracaoPedCompraCot.colCotacao(gobjGeracaoPedCompraCot.iCotacaoSel)

        lErro = CF("Cotacao_Le_Tudo", objCotacao)
        If lErro <> SUCESSO Then gError 62803

        lErro = CF("Requisicoes_Le_Cotacao", objCotacao, gcolRequisicaoCompra)
        If lErro <> SUCESSO And lErro <> 70441 Then gError 62804
        
    End If
        
    If gcolRequisicaoCompra.Count > 0 Then


        For Each objCotacaoProduto In objCotacao.colCotacaoProduto

            For Each objItemRCCot In objCotacaoProduto.colItemReqCompras
                bAchou = False
                For Each objReqCompras In gcolRequisicaoCompra
                    objReqCompras.iSelecionado = MARCADO

                    For Each objItemReq In objReqCompras.colItens

                        If objItemReq.lNumIntDoc = objItemRCCot.lNumIntDoc Then
                            objItemReq.iSelecionado = MARCADO

                            lErro = ItensConcorrencia_Cria_Altera(objItemReq)
                            If lErro <> SUCESSO Then gError 62805

                            bAchou = True
                            Exit For
                        End If
                    Next
                    If bAchou Then Exit For
                Next
            Next
        Next

        lErro = Traz_Requisicoes_Tela(gobjGeracaoPedCompraCot)
        If lErro <> SUCESSO Then gError 62802
    Else
        lErro = CotacaoProduto_Cria_ItensConcorrencia(objCotacao)
        If lErro <> SUCESSO Then gError 62817

    End If
    
    'Preenche os grids de produto correspondentes aos itens de concorrência
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62748
              
    Traz_Cotacao_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Cotacao_Tela:

    Traz_Cotacao_Tela = gErr
    
    Select Case gErr
    
        Case 62748, 62802 To 62805, 62817
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161183)
            
    End Select
    
    Exit Function
    
End Function

Function CotacaoProduto_Cria_ItensConcorrencia(objCotacao As ClassCotacao) As Long
'Cria itens de concorrência através das cotações do produto

Dim lErro As Long
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objQtSupl As ClassQuantSuplementar
Dim iItem As Integer
Dim objProduto As New ClassProduto
Dim colProdutos As New Collection
Dim colFiliaisFornec As New Collection
Dim objFilialFornec As ClassFilialFornecedor

On Error GoTo Erro_CotacaoProduto_Cria_ItensConcorrencia

    iItem = 0

    'Para cada cotação do produto
    For Each objCotacaoProduto In objCotacao.colCotacaoProduto
    
        iItem = iItem + 1
        'Cria um novo item de concorrência
        Set objItemConcorrencia = New ClassItemConcorrencia
        
        objProduto.sCodigo = objCotacaoProduto.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> AD_SQL_SUCESSO And lErro <> 23080 Then gError 86117
        If lErro <> AD_SQL_SUCESSO Then gError 86118
        
        'Preenche os dados os item de concorrência
        objItemConcorrencia.sProduto = objCotacaoProduto.sProduto
        objItemConcorrencia.sDescricao = objProduto.sDescricao
        objItemConcorrencia.sUM = objCotacaoProduto.sUM
        objItemConcorrencia.dQuantidade = objCotacaoProduto.dQuantidade
        objItemConcorrencia.dtDataNecessidade = DATA_NULA
        objItemConcorrencia.iEscolhido = MARCADO
        objItemConcorrencia.iFilial = objCotacaoProduto.iFilial
        objItemConcorrencia.lFornecedor = objCotacaoProduto.lFornecedor
        
        'Adiciona o item criado na coleção de itens de concorrência
        gcolItemConcorrencia.Add objItemConcorrencia
        
        'Se tiver tipo de destino amarrado
        If objCotacao.iTipoDestino <> TIPO_DESTINO_AUSENTE Then
            
            'Inicializa as quantidades suplementares
            Set objQtSupl = New ClassQuantSuplementar
            
            objQtSupl.dQuantidade = objCotacaoProduto.dQuantidade
            objQtSupl.dQuantRequisitada = 0
            objQtSupl.iFilialDestino = objCotacao.iFilialDestino
            objQtSupl.iTipoDestino = objCotacao.iTipoDestino
            objQtSupl.lFornCliDestino = objCotacao.lFornCliDestino
            
            'Adiciona na coleção de quantidades suplementares do item
            objItemConcorrencia.colQuantSuplementar.Add objQtSupl
        
        End If
            
        'Atualiza as cotações dos itensa de concorrência
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iItem)
        If lErro <> SUCESSO Then gError 62821
    Next

    CotacaoProduto_Cria_ItensConcorrencia = SUCESSO
    
    Exit Function
    
Erro_CotacaoProduto_Cria_ItensConcorrencia:

    CotacaoProduto_Cria_ItensConcorrencia = gErr

    Select Case gErr

        Case 62821, 86117
        
        Case 86118
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161184)
    
    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)
        
    Select Case objControl.Name
    
        Case QuantComprarItemRC.Name
        
            If StrParaInt(GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col)) = DESMARCADO Then
                QuantComprarItemRC.Enabled = False
            Else
                QuantComprarItemRC.Enabled = True
            End If
    
        'MotivoEscolha
        Case MotivoEscolhaCot.Name

            If objControl.Name = MotivoEscolhaCot.Name And _
               GridCotacoes.TextMatrix(iLinha, iGrid_MotivoEscolhaCot_Col) = MOTIVO_EXCLUSIVO_DESCRICAO Then
               objControl.Enabled = False
            Else
               objControl.Enabled = True
            End If
                
        Case Quantidade2.Name
            'Se o usuário puder aumentar a quantidade requisitada
            If gcolRequisicaoCompra.Count > 0 Then
                If giPodeAumentarQuant = MARCADO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            Else
                objControl.Enabled = True
            End If
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodosItensRC_Click()
'Marca todas CheckBox do GridItensRequisicoes

Dim lErro As Long
Dim iItem As Integer
Dim iIndice As Integer
Dim objItemRC As ClassItemReqCompras
Dim colIndices As New Collection
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemConc As New ClassItemConcorrencia

On Error GoTo Erro_BotaoMarcarTodosItensRC_Click
    
    If gcolRequisicaoCompra.Count = 0 Then Exit Sub
    
    'Para cada Req selecionada
    For Each objReqCompra In gcolRequisicaoCompra
        'se a req está selecionada
        If objReqCompra.iSelecionado = MARCADO Then
            'marca os itens de requisicao
            For Each objItemRC In objReqCompra.colItens
                If objItemRC.iSelecionado = DESMARCADO Then
                    objItemRC.iSelecionado = MARCADO
                    
                    'Cria ou Altera os itens de concorrencia existentes
                    lErro = ItensConcorrencia_Cria_Altera(objItemRC)
                    If lErro <> SUCESSO Then gError 62757
                         
                    Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConc, iItem, objItemRC)
                    
                    Call Adiciona_Codigo(colIndices, iItem)
                    
                End If
            Next
        End If
    Next
    
    'Atualiza as cotações
    For iIndice = 1 To colIndices.Count
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, colIndices(iIndice))
        If lErro <> SUCESSO Then gError 62766
    Next
    
    'seleciona no grid
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes
        GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col) = MARCADO
    Next
    
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)
    
    'Prenche o grid de produtos
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62758
    
    Exit Sub

Erro_BotaoMarcarTodosItensRC_Click:

    Select Case gErr

        Case 62766, 62757, 62758

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161185)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava a Concorrencia

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    'Recolhe os dados da tela e armazena em objConcorrencia
    lErro = Move_Concorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63761

    'Insere ou Altera uma concorrencia no BD
    lErro = CF("Concorrencia_Grava", objConcorrencia)
    If lErro <> SUCESSO Then gError 63672

    Call Rotina_Aviso(vbOKOnly, "AVISO_CONCORRENCIA_GRAVADA", objConcorrencia.lCodigo)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
    
    Select Case gErr

        Case 63756

        Case 63761, 63672
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161186)

    End Select

    Exit Function

End Function
Private Sub Busca_Produto(sProduto As String, colProdutos As Collection, objProduto As ClassProduto, bAchou As Boolean)

Dim objProdAux As ClassProduto

    bAchou = False
    
    For Each objProdAux In colProdutos
        
        If objProdAux.sCodigo = sProduto Then
            bAchou = True
            Set objProduto = objProdAux
            Exit For
        End If
    
    Next

    Exit Sub

End Sub

Public Sub BotaoEditarProduto_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_BotaoEditarProduto_Click

    'Se está editando um produto do GridProdutos1
    If FrameProdutos(1).Visible = True Then

        'Verifica se tem alguma linha selecionada no GridProdutos1
        If GridProdutos1.Row = 0 Then gError 66760

        'Verifica se o Produto está preenchido
        If Len(Trim(GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_Produto1_Col))) > 0 Then
            lErro = CF("Produto_Formata", GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_Produto1_Col), sProduto, iPreenchido)
            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        End If

    'Se está editando um produto do GridProdutos2
    Else

        'Verifica se tem alguma linha selecionada no GridProdutos1
        If GridProdutos2.Row = 0 Then gError 66943

        'Verifica se o Produto está preenchido
        If Len(Trim(GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col))) > 0 Then
            lErro = CF("Produto_Formata", GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col), sProduto, iPreenchido)
            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        End If

    End If

    objProduto.sCodigo = sProduto

    'Chama a Tela de Produto
    Call Chama_Tela("Produto", objProduto)

    Exit Sub

Erro_BotaoEditarProduto_Click:

    Select Case gErr

        Case 66943, 66760
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161187)

    End Select

    Exit Sub

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161188)
    
    End Select

End Function

Private Sub Indica_Melhores()
'Indica as melhores opcoes

Dim dMenorPreco As Double
Dim objItemCotItemConc As ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia
Dim objItemCotItemConcAux As ClassCotacaoItemConc

On Error GoTo Erro_Indica_Melhores

    Call Grid_Refresh_Checkbox_Limpa(objGridCotacoes)
    
    For Each objItemConcorrencia In gcolItemConcorrencia
        
        dMenorPreco = 0
        
        Set objItemCotItemConcAux = New ClassCotacaoItemConc
        
        'Para cada produto da colecao ...
         For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
            
            'Se for para aparecer no grid ...
            If objItemCotItemConc.iSelecionada = MARCADO Then
            
                'Desmarca.
                objItemCotItemConc.iEscolhido = DESMARCADO
                
                'Caso ainda nao tenhamos um menor preco => Menor = $$ do Primeiro item
                If dMenorPreco = 0 Then
                    
                    dMenorPreco = objItemCotItemConc.dPrecoAjustado
                    
                    Set objItemCotItemConcAux = New ClassCotacaoItemConc
                    Set objItemCotItemConcAux = objItemCotItemConc
                    
                End If
                
                'Se o preco for menor do que o menor preco ja encontrado
                If objItemCotItemConc.dPrecoAjustado < dMenorPreco Then
                    
                    'Guarda o menor preco
                    dMenorPreco = objItemCotItemConc.dPrecoAjustado
                    
                    'Coloca o preco anterior como desmarcado
                    objItemCotItemConcAux.iEscolhido = DESMARCADO
                    
                    'Aponta para o novo candidato
                    Set objItemCotItemConcAux = New ClassCotacaoItemConc
                    Set objItemCotItemConcAux = objItemCotItemConc
                    
                End If
            
            End If
            
        Next
        
        'Seleciona o Menor
        objItemCotItemConcAux.iEscolhido = MARCADO
        
    Next
    
    Call Grid_Refresh_Checkbox(objGridCotacoes)

    Exit Sub

Erro_Indica_Melhores:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161189)
    
    End Select

End Sub

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitario.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoUnitarioReal.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

