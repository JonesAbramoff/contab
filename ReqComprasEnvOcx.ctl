VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ReqComprasEnvOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Itens"
      Height          =   8280
      Index           =   2
      Left            =   165
      TabIndex        =   36
      Top             =   675
      Visible         =   0   'False
      Width           =   16650
      Begin MSMask.MaskEdBox TipoTribItem 
         Height          =   255
         Left            =   555
         TabIndex        =   47
         Top             =   3750
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   12
         Mask            =   "TipoTribItem"
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoPedidos 
         Caption         =   "Pedidos"
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
         Left            =   60
         TabIndex        =   53
         Top             =   285
         Width           =   1515
      End
      Begin VB.TextBox ContaContabil 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   48
         Top             =   4080
         Width           =   1455
      End
      Begin VB.ComboBox Exclusivo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ReqComprasEnvOcx.ctx":0000
         Left            =   5715
         List            =   "ReqComprasEnvOcx.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   4200
         Width           =   1305
      End
      Begin VB.ComboBox FilialForn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4335
         TabIndex        =   50
         Top             =   4065
         Width           =   1770
      End
      Begin VB.TextBox Observacao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6255
         MaxLength       =   255
         TabIndex        =   52
         Top             =   4020
         Width           =   2355
      End
      Begin VB.ComboBox UM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2775
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   930
         Width           =   1065
      End
      Begin VB.TextBox DescProduto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   39
         Top             =   1020
         Width           =   4000
      End
      Begin MSMask.MaskEdBox QuantCancelada 
         Height          =   225
         Left            =   6405
         TabIndex        =   44
         Top             =   1020
         Width           =   870
         _ExtentX        =   1535
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
      Begin MSMask.MaskEdBox CentroCusto 
         Height          =   225
         Left            =   105
         TabIndex        =   46
         Top             =   4050
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
         Left            =   5655
         TabIndex        =   43
         Top             =   975
         Width           =   780
         _ExtentX        =   1376
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
         Left            =   4800
         TabIndex        =   42
         Top             =   1020
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   7215
         TabIndex        =   45
         Top             =   1005
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
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   3855
         TabIndex        =   41
         Top             =   1005
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   195
         TabIndex        =   38
         Top             =   960
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fornec 
         Height          =   225
         Left            =   2610
         TabIndex        =   49
         Top             =   4110
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   6855
         Left            =   15
         TabIndex        =   37
         Top             =   1110
         Width           =   16545
         _ExtentX        =   29184
         _ExtentY        =   12091
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   8250
      Index           =   3
      Left            =   120
      TabIndex        =   60
      Top             =   675
      Visible         =   0   'False
      Width           =   16680
      Begin VB.TextBox UMPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   5325
         MaxLength       =   50
         TabIndex        =   66
         Top             =   2115
         Width           =   1080
      End
      Begin VB.TextBox DescProdPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   63
         Top             =   2130
         Width           =   4000
      End
      Begin MSMask.MaskEdBox QuantReqPV 
         Height          =   225
         Left            =   5850
         TabIndex        =   62
         Top             =   2130
         Width           =   1365
         _ExtentX        =   2408
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
      Begin MSMask.MaskEdBox CodPV 
         Height          =   225
         Left            =   4410
         TabIndex        =   64
         Top             =   1665
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ItemReqPV 
         Height          =   225
         Left            =   285
         TabIndex        =   65
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox QuantPV 
         Height          =   225
         Left            =   6780
         TabIndex        =   67
         Top             =   2145
         Width           =   1365
         _ExtentX        =   2408
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
      Begin MSMask.MaskEdBox ProdutoPV 
         Height          =   225
         Left            =   990
         TabIndex        =   68
         Top             =   2130
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridPV 
         Height          =   8010
         Left            =   105
         TabIndex        =   61
         Top             =   135
         Width           =   16530
         _ExtentX        =   29157
         _ExtentY        =   14129
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
      Caption         =   "Cabeçalho"
      Height          =   8310
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   675
      Width           =   16695
      Begin VB.Frame Frame5 
         Caption         =   "Cabeçalho"
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   8775
         Begin VB.Frame Frame6 
            Caption         =   "Local de Entrega"
            Height          =   960
            Left            =   225
            TabIndex        =   17
            Top             =   2475
            Width           =   8355
            Begin VB.Frame Frame2 
               Caption         =   "Tipo"
               Height          =   585
               Left            =   195
               TabIndex        =   18
               Top             =   210
               Width           =   3945
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
                  Left            =   2295
                  TabIndex        =   20
                  Top             =   255
                  Width           =   1335
               End
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
                  Left            =   435
                  TabIndex        =   19
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1515
               End
            End
            Begin VB.Frame FrameTipo 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   675
               Index           =   0
               Left            =   4545
               TabIndex        =   21
               Top             =   195
               Width           =   3495
               Begin VB.Label Label37 
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
                  Left            =   540
                  TabIndex        =   22
                  Top             =   195
                  Width           =   465
               End
               Begin VB.Label FilialEmpresa 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1065
                  TabIndex        =   23
                  Top             =   165
                  Width           =   2145
               End
            End
            Begin VB.Frame FrameTipo 
               BorderStyle     =   0  'None
               Height          =   675
               Index           =   1
               Left            =   4665
               TabIndex        =   24
               Top             =   195
               Visible         =   0   'False
               Width           =   3450
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
                  Index           =   1
                  Left            =   90
                  TabIndex        =   25
                  Top             =   90
                  Width           =   1035
               End
               Begin VB.Label FornecDestino 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1185
                  TabIndex        =   26
                  Top             =   15
                  Width           =   2145
               End
               Begin VB.Label Label32 
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
                  TabIndex        =   27
                  Top             =   360
                  Width           =   465
               End
               Begin VB.Label FilialFornec 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1185
                  TabIndex        =   28
                  Top             =   360
                  Width           =   2145
               End
            End
         End
         Begin VB.CheckBox Urgente 
            Caption         =   "Urgente"
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
            Left            =   6075
            TabIndex        =   14
            Top             =   1215
            Width           =   1005
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1440
            TabIndex        =   59
            Top             =   330
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
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
            Left            =   5040
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   70
            Top             =   1710
            Width           =   960
         End
         Begin VB.Label CodigoPV 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6045
            TabIndex        =   69
            Top             =   1650
            Width           =   1020
         End
         Begin VB.Label Label21 
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
            Height          =   195
            Left            =   390
            TabIndex        =   12
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label CodigoOP 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   1620
            Width           =   915
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
            Height          =   195
            Left            =   675
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   3
            Top             =   375
            Width           =   690
         End
         Begin VB.Label Requisitante 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6060
            TabIndex        =   5
            Top             =   315
            Width           =   2565
         End
         Begin VB.Label Label5 
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
            Height          =   195
            Left            =   4875
            TabIndex        =   4
            Top             =   375
            Width           =   1140
         End
         Begin VB.Label FilialCompra 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1440
            TabIndex        =   7
            Top             =   750
            Width           =   2565
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
            Height          =   195
            Left            =   210
            TabIndex        =   6
            Top             =   780
            Width           =   1155
         End
         Begin VB.Label Ccl 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   1200
            Width           =   915
         End
         Begin VB.Label CclPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "C. Custo:"
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
            Left            =   570
            TabIndex        =   10
            Top             =   1260
            Width           =   795
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
            Index           =   1
            Left            =   4320
            TabIndex        =   8
            Top             =   795
            Width           =   1695
         End
         Begin VB.Label TipoTributacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6060
            TabIndex        =   9
            Top             =   735
            Width           =   2520
         End
         Begin VB.Label Observ 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   2085
            Width           =   7260
         End
         Begin VB.Label Label4 
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
            Index           =   0
            Left            =   270
            TabIndex        =   15
            Top             =   2145
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Datas"
         Height          =   795
         Left            =   120
         TabIndex        =   29
         Top             =   4440
         Width           =   8790
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1170
            TabIndex        =   31
            Top             =   300
            Width           =   1095
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
            Height          =   195
            Left            =   660
            TabIndex        =   30
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Data Envio:"
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
            Left            =   3060
            TabIndex        =   32
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label DataEnvio 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4110
            TabIndex        =   33
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label DataLimite 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7065
            TabIndex        =   35
            Top             =   300
            Width           =   1095
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
            Left            =   6000
            TabIndex        =   34
            Top             =   360
            Width           =   1035
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   14760
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   595
         Picture         =   "ReqComprasEnvOcx.ctx":0027
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoBaixar 
         Height          =   360
         Left            =   1100
         Picture         =   "ReqComprasEnvOcx.ctx":0181
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Baixar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ReqComprasEnvOcx.ctx":0343
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   90
         Picture         =   "ReqComprasEnvOcx.ctx":04C1
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8730
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   16875
      _ExtentX        =   29766
      _ExtentY        =   15399
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
Attribute VB_Name = "ReqComprasEnvOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim iAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim gColItensReqCompras As Collection
Dim iChamaTela As Integer

Dim objGridItens As AdmGrid

Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_QuantPedida_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_QuantCancelada_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_CCL_Col As Integer
Dim iGrid_ContaContabil_Col As Integer
Dim iGrid_TipoTributacao_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialFornecedor_Col As Integer
Dim iGrid_Exclusivo_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iFrameAtual As Integer

'GridPV
Dim objGridPV As AdmGrid
Dim iGrid_ItemReqPV_Col As Integer
Dim iGrid_ProdutoPV_Col As Integer
Dim iGrid_DescProdPV_Col As Integer
Dim iGrid_CodPV_Col As Integer
Dim iGrid_UMPV_Col As Integer
Dim iGrid_QuantReqPV_Col As Integer
Dim iGrid_QuantPV_Col As Integer

Private WithEvents objEventoPedidoCompras As AdmEvento
Attribute objEventoPedidoCompras.VB_VarHelpID = -1
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoBotaoRequisicoes As AdmEvento
Attribute objEventoBotaoRequisicoes.VB_VarHelpID = -1

Private Sub BotaoBaixar_Click()

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoBaixar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se existe alguma Requisição de Compras
    If Len(Trim(Codigo.Text)) = 0 Then gError 54182

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BAIXA_REQCOMPRAS", Codigo.Text)
    If vbMsgRes = vbNo Then gError 62645

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 54183
    
    lErro = CF("RequisicaoCompras_Le_Numero", objRequisicaoCompras)
    If lErro <> SUCESSO And lErro <> 54439 Then gError 86101
    If lErro <> SUCESSO Then gError 86102
    
    If objRequisicaoCompras.dtDataEnvio = DATA_NULA Then gError 86103
    
    'Baixa a Requisição de Compras
    lErro = CF("RequisicaoCompraEnv_Baixar", objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 54186

    'Limpa a tela
    Call Limpa_Tela_ReqComprasEnv

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixar_Click:

    Select Case gErr

        Case 54182
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_CARREGADA", gErr)

        Case 54183, 54186, 62645, 86101
        
        Case 86102
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_CADASTRADA", gErr, objRequisicaoCompras.lCodigo, objRequisicaoCompras.iFilialEmpresa)
        
        Case 86103
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_ENVIADA", gErr, objRequisicaoCompras.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173831)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 54333

    'Limpa a tela
    Call Limpa_Tela_ReqComprasEnv

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 54333

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173832)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim objReqCompras As New ClassRequisicaoCompras

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 54334

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 54335

    objReqCompras.lCodigo = objRequisicaoCompras.lCodigo
    objReqCompras.iFilialEmpresa = objRequisicaoCompras.iFilialEmpresa

    lErro = CF("RequisicaoCompras_Le_Numero", objReqCompras)
    If lErro <> SUCESSO And lErro <> 54439 Then gError 86104
    
    'Se a Requisicao não existe ==> erro
    If lErro = 54439 Then gError 86105
    If objReqCompras.dtDataEnvio = DATA_NULA Then gError 86106
    
    'Chama Rotina para iniciar a gravação
    lErro = CF("RequisicaoCompraEnv_Grava", objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 54337

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 54334
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_CARREGADA", gErr)

        Case 54335, 54337, 86104
        
        Case 86105
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objReqCompras.lCodigo)
            
        Case 86106
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_ENVIADA", gErr, objReqCompras.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173833)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Function Move_GridItens_Memoria(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Recolhe do Grid os dados do itens para a memória

Dim lErro As Long
Dim sProduto As String
Dim sCclFormatada As String
Dim sContaFormatada As String
Dim objItemReqCompras As New ClassItemReqCompras
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim iCont As Integer
Dim iCclPreenchida As Integer
Dim iContaPreenchida As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemReqCompras = New ClassItemReqCompras

        'Verifica se o Produto está preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then

            'Formata o produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then Error 54338

            'Armazena o Produto
            objItemReqCompras.sProduto = sProduto

        End If

        'Armazena a descrição
        objItemReqCompras.sDescProduto = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        'Armazena os restante dos dados do item
        objItemReqCompras.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemReqCompras.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemReqCompras.dQuantPedida = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
        objItemReqCompras.dQuantRecebida = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantRecebida_Col))
        objItemReqCompras.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancelada_Col))

        If objItemReqCompras.iAlmoxarifado > 0 Then
            objAlmoxarifado.iCodigo = objItemReqCompras.iAlmoxarifado
            
            'Lê o almoxarifado com o código passado por objAlmoxarifado
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then Error 54417
            'Se não encontrou --> Erro
            If lErro = 25056 Then Error 54418
            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
        End If

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CCL_Col))) > 0 Then
            
            sCclFormatada = String(STRING_CCL, 0)
        
            lErro = CF("Ccl_Formata", GridItens.TextMatrix(iIndice, iGrid_CCL_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 54419
        
            objItemReqCompras.sCcl = sCclFormatada
        End If
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col))) > 0 Then
            sContaFormatada = String(STRING_CONTA, 0)
            
            lErro = CF("Conta_Formata", GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col), sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 54420

            objItemReqCompras.sContaContabil = sContaFormatada
        End If
        objItemReqCompras.iTipoTributacao = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col))
        
        If Len(Trim(objFornecedor.sNomeReduzido)) > 0 Then
            objFornecedor.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col)
        
            'Lê os dados do Fornecedor a partir do seu Nome Reduzido
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 54423
            If lErro = 6681 Then Error 54424
        
            objItemReqCompras.lFornecedor = objFornecedor.lCodigo
        
        End If

        objItemReqCompras.iFilial = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col))

        For iCont = 0 To Exclusivo.ListCount - 1
            If Exclusivo.List(iCont) = GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) Then
                objItemReqCompras.iExclusivo = Exclusivo.ItemData(iCont)
                Exit For
            End If
        Next
        
        objItemReqCompras.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)

        objItemReqCompras.lNumIntDoc = gColItensReqCompras.Item(iIndice)

        'Adiciona o item na colecao de itens do pedido de compra
        objRequisicaoCompras.colItens.Add objItemReqCompras

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = Err

    Select Case Err

        Case 54338, 54417, 54419, 54420, 54423

        Case 54418
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)
        
        Case 54424
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173834)

    End Select

    Exit Function

End Function

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objReqCompras As New ClassRequisicaoCompras
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 76027
    
    objReqCompras.lCodigo = StrParaLong(Codigo.Text)
    objReqCompras.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se a Requisicao de Compra informada existe
    lErro = CF("RequisicaoCompras_Le_Numero", objReqCompras)
    If lErro <> SUCESSO And lErro <> 54439 Then gError 76030
    
    'Se a Requisicao não existe ==> erro
    If lErro = 54439 Then gError 76031
    If objReqCompras.dtDataEnvio = DATA_NULA Then gError 86107
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Requisições de Compra Enviada", "REQCOMTO.NumIntDoc = @NREQCOM", 1, "REQCOM", "NREQCOM", objReqCompras.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76032
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 76027
            Call Rotina_Erro(vbOKOnly, "ERRO_REQCOMPRAS_IMPRESSAO", gErr)
            
        Case 76030, 76032
        
        Case 76031
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objReqCompras.lCodigo)
        
        Case 86107
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_ENVIADA", gErr, objReqCompras.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173835)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoPedidos_Click()

Dim lErro As Long
Dim iTela As Integer
Dim colSelecao As New Collection
Dim objPedidoCompras As New ClassPedidoCompras
Dim objItemReqCompras As New ClassItemReqCompras

On Error GoTo Erro_BotaoPedidos_Click

    'Se não tiver nennhum item selecionado no grid de itens, sai da rotina
    If GridItens.Row = 0 Then gError 89442

    objItemReqCompras.lNumIntDoc = gColItensReqCompras.Item(GridItens.Row)
    
    'Passa a posicao do item no grid
    iTela = GridItens.Row
    
    'Verifica que tela de Pedido de Compras chamar: Pedidos Associados, Enviado, novo, Baixado
    lErro = CF("ItemReqCompra_Busca_PedidoCompra", iTela, objPedidoCompras, objItemReqCompras)
    If lErro <> SUCESSO And lErro <> 54403 Then gError 54411

    'Se não há Pedidos vinculados ao ItemRC
    If lErro = 54403 Then gError 75474
    
    Select Case iTela

        Case 0, 1
            'Chama a tela de Pedido Compras Enviado
            Call Chama_Tela("PedComprasCons", objPedidoCompras)

        Case 2
            'Chama tela de Pedido Compras Novo
            Call Chama_Tela("PedComprasGerado", objPedidoCompras)

        Case 3
            'Chama a tela de Pedidos Associados
            colSelecao.Add objItemReqCompras.lNumIntDoc
      
            Call Chama_Tela("PedComprasItemReqLista", colSelecao, objPedidoCompras, objEventoPedidoCompras)
    
    End Select

    Exit Sub

Erro_BotaoPedidos_Click:

    Select Case gErr

        Case 54411

        Case 75474
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMRC_DESVINCULADO_ITEMPC", gErr, GridItens.Row)

        Case 89442
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173836)

    End Select

    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'Inicialização
Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New Collection

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    'Inicializa os ObjEventos
    Set objEventoCodigo = New AdmEvento
    Set objEventoBotaoRequisicoes = New AdmEvento
    Set objEventoPedidoCompras = New AdmEvento
    
    'Atualiza a global e seta um tipo Padrao
    iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True

    'Inicializa o GridItens
    Set objGridItens = New AdmGrid

    Set gColItensReqCompras = New Collection

    'Faz a inicialização do Grid
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then Error 54358

    Set objGridPV = New AdmGrid

    lErro = Inicializa_GridPV(objGridPV)
    If lErro <> SUCESSO Then gError 178665

    'Inicializa mascara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 54409
            
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_LIGHT Then
        
        FilialCompra.left = POSICAO_FORA_TELA
        Label3.left = POSICAO_FORA_TELA
        Label3.Visible = False
        FilialEmpresa.left = POSICAO_FORA_TELA
        Label37.left = POSICAO_FORA_TELA
        Label37.Visible = False
        
    End If
            
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 54358, 54409, 178665

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173837)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid

Dim lErro As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Em Pedido")
    objGridInt.colColuna.Add ("Recebida")
    objGridInt.colColuna.Add ("Cancelada")
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
    objGridInt.colCampo.Add (QuantPedida.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (QuantCancelada.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)
    objGridInt.colCampo.Add (TipoTribItem.Name)
    objGridInt.colCampo.Add (Fornec.Name)
    objGridInt.colCampo.Add (FilialForn.Name)
    objGridInt.colCampo.Add (Exclusivo.Name)
    objGridInt.colCampo.Add (Observacao.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_QuantPedida_Col = 5
    iGrid_QuantRecebida_Col = 6
    iGrid_QuantCancelada_Col = 7
    iGrid_Almoxarifado_Col = 8
    iGrid_CCL_Col = 9
    iGrid_ContaContabil_Col = 10
    iGrid_TipoTributacao_Col = 11
    iGrid_Fornecedor_Col = 12
    iGrid_FilialFornecedor_Col = 13
    iGrid_Exclusivo_Col = 14
    iGrid_Observacao_Col = 15

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 18
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura manual para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'GridItens.Width = 8880

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridItens.Col

            Case iGrid_QuantCancelada_Col

                'Quantidade Cancelada
                lErro = Saida_Celula_QuantCancelada(objGridInt)
                If lErro <> SUCESSO Then Error 54406

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 54405

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 54405
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54406

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173838)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantCancelada(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_QuantCancelada

    Set objGridInt.objControle = QuantCancelada

    'Se quantidade estiver preenchida
    If Len(Trim(QuantCancelada.ClipText)) > 0 Then
        
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantCancelada.Text)
        If lErro <> SUCESSO Then Error 54404

        dQuantidade = StrParaDbl(QuantCancelada.Text)

        If dQuantidade > StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col)) - StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantRecebida_Col)) - StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col)) Then Error 54435
        
        'Coloca o valor Formatado na tela
        QuantCancelada.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 54407

    Saida_Celula_QuantCancelada = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantCancelada:

    Saida_Celula_QuantCancelada = Err

    Select Case Err

        Case 54404, 54407
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54435
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCANCELADA_MAIOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173839)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objRequisicaoCompras As ClassRequisicaoCompras) As Long

Dim lErro As Long
Dim bAchou As Boolean

On Error GoTo Erro_Trata_Parametros
    
    bAchou = False

    'Verifica se alguma Requisição foi passada por parâmetro
    If Not (objRequisicaoCompras Is Nothing) Then

        'Se o número interno estiver preenchido
        If objRequisicaoCompras.lNumIntDoc > 0 Then

            'Le a Requisição de Compras
            lErro = CF("RequisicaoCompras_Le", objRequisicaoCompras)
            If lErro <> SUCESSO And lErro <> 54397 Then Error 54369
            If lErro = SUCESSO Then bAchou = True
        End If
        'Se o número interno não estiver preenchido mas o código estiver
        If objRequisicaoCompras.lCodigo > 0 Then
            
            'Lê a Requisição de Compras pelo código e Filial Empresa
            lErro = CF("RequisicaoCompras_Le_Numero", objRequisicaoCompras)
            If lErro <> SUCESSO And lErro <> 54439 Then Error 54433
            If lErro = SUCESSO Then bAchou = True
           
        End If
        
        If bAchou Then
        
             'Se a data de envio está preenchida
            If objRequisicaoCompras.dtDataEnvio = DATA_NULA Then Error 54372
    
            'Traz os dados da Requisição para a tela
            lErro = Traz_RequisicaoEnv_Tela(objRequisicaoCompras)
            If lErro <> SUCESSO Then Error 54371
        End If
    'Se não foi passado nenhuma Requisição como parâmetro
    Else
         
        'Limpa a tela ReqComprasEnv
        Call Limpa_Tela_ReqComprasEnv
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 54369, 54370, 54371, 54433
        
        Case 54434
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA2", Err)

        Case 54372
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_ENVIADA", Err, objRequisicaoCompras.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173840)

    End Select

    Exit Function

End Function

'*****************************************************************************************************************************
''SISTEMA DE SETAS
'*****************************************************************************************************************************
'Extrai os campos da tela que correspondem aos campos no Banco de Dados
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim sNomeRed As String

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ReqCompras_Requisitante"
    
    ' Move todos os dados Presentes na Tela em objRequisaoModelo
    lErro = Move_Tela_Memoria(objRequisicaoCompras)
    If lErro <> SUCESSO Then Error 54373

    sNomeRed = Requisitante.Caption
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objRequisicaoCompras.lCodigo, 0, "Codigo"
    colCampoValor.Add "Urgente", objRequisicaoCompras.lUrgente, 0, "Urgente"
    colCampoValor.Add "Requisitante", objRequisicaoCompras.lRequisitante, 0, "Requisitante"
    colCampoValor.Add "Ccl", objRequisicaoCompras.sCcl, STRING_BUFFER_MAX_TEXTO, "Ccl"
    colCampoValor.Add "FilialCompra", objRequisicaoCompras.iFilialCompra, 0, "FilialCompra"
    colCampoValor.Add "Observacao", objRequisicaoCompras.lObservacao, 0, "Observacao"
    colCampoValor.Add "Data", objRequisicaoCompras.dtData, 0, "Data"
    colCampoValor.Add "DataEnvio", objRequisicaoCompras.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataLimite", objRequisicaoCompras.dtDataLimite, 0, "DataLimite"
    colCampoValor.Add "TipoDestino", objRequisicaoCompras.iTipoDestino, 0, "TipoDestino"
    colCampoValor.Add "FornCliDestino", objRequisicaoCompras.lFornCliDestino, 0, "FornCliDestino"
    colCampoValor.Add "FilialDestino", objRequisicaoCompras.iFilialDestino, 0, "FilialDestino"
    colCampoValor.Add "NumIntDoc", objRequisicaoCompras.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoTributacao", objRequisicaoCompras.iTipoTributacao, 0, "TipoTributacao"
    colCampoValor.Add "OPCodigo", objRequisicaoCompras.sOPCodigo, STRING_BUFFER_MAX_TEXTO, "OPCodigo"
    colCampoValor.Add "NomeReduzido", sNomeRed, STRING_BUFFER_MAX_TEXTO, "NomeReduzido"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "DataEnvio", OP_DIFERENTE, DATA_NULA

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 54373

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173841)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do Banco de Dados
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objRequisicaoCompras.lCodigo = colCampoValor.Item("Codigo").vValor
    objRequisicaoCompras.iFilialEmpresa = giFilialEmpresa
    
    If objRequisicaoCompras.lCodigo <> 0 Then
        objRequisicaoCompras.lUrgente = colCampoValor.Item("Urgente").vValor
        objRequisicaoCompras.lRequisitante = colCampoValor.Item("Requisitante").vValor
        objRequisicaoCompras.sCcl = colCampoValor.Item("Ccl").vValor
        objRequisicaoCompras.iFilialCompra = colCampoValor.Item("FilialCompra").vValor
        objRequisicaoCompras.lObservacao = colCampoValor.Item("Observacao").vValor
        objRequisicaoCompras.dtData = colCampoValor.Item("Data").vValor
        objRequisicaoCompras.dtDataEnvio = colCampoValor.Item("DataEnvio").vValor
        objRequisicaoCompras.dtDataLimite = colCampoValor.Item("DataLimite").vValor
        objRequisicaoCompras.iTipoDestino = colCampoValor.Item("TipoDestino").vValor
        objRequisicaoCompras.lFornCliDestino = colCampoValor.Item("FornCliDestino").vValor
        objRequisicaoCompras.iFilialDestino = colCampoValor.Item("FilialDestino").vValor
        objRequisicaoCompras.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
        objRequisicaoCompras.iTipoTributacao = colCampoValor.Item("TipoTributacao").vValor
        objRequisicaoCompras.sOPCodigo = colCampoValor.Item("OPCodigo").vValor

        'Traz os dados da Requisição para a tela
        lErro = Traz_RequisicaoEnv_Tela(objRequisicaoCompras)
        If lErro <> SUCESSO Then Error 54374

    End If
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 54374

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173842)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'Funções que retiram e colocam os dados na tela
'""""""""""""""""""""""""""""""""""""""""""""""
Function Move_Tela_Memoria(objRequisicaoCompras As ClassRequisicaoCompras) As Long
' Move os dados da tela para o objRequisicaoCompras

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Move_Tela_Memoria

    'Move o código para a memória
    objRequisicaoCompras.lCodigo = StrParaLong(Codigo.Text)

    'Move a FilialEmpresa
    objRequisicaoCompras.iFilialEmpresa = giFilialEmpresa
    
    objRequisicaoCompras.sDigitador = gsUsuario
    objRequisicaoCompras.dtDataBaixa = DATA_NULA
    
    'Requisitante
    If Len(Trim(Requisitante.Caption)) > 0 Then
        
        objRequisitante.sNomeReduzido = Requisitante.Caption
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then gError 75471
        
        'Requisitante não cadastrado
        If lErro = 51152 Then gError 75472
        
        objRequisicaoCompras.lRequisitante = objRequisitante.lCodigo
    
    End If
    
    'Move o dados do Grid para a memória
    lErro = Move_GridItens_Memoria(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 54363

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr

        Case 54363, 75471
        
        Case 75472
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO1", gErr, Requisitante.Caption)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173843)

    End Select

    Exit Function

End Function

Function Traz_RequisicaoEnv_Tela(objRequisicaoCompras As ClassRequisicaoCompras) As Long

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim objFilial As New AdmFiliais
Dim objObservacao As New ClassObservacao
Dim sDescricao As String

On Error GoTo Erro_Traz_RequisicaoEnv_Tela

    'Limpa tela Requisicao Compras
    Call Limpa_Tela_ReqComprasEnv

    'Lê os itens da Requisição Compras Enviadas
    lErro = CF("ItensReqCompras_Le", objRequisicaoCompras)
    If lErro <> SUCESSO And lErro <> 54377 Then gError 54378
    
    'Coloca os dados na tela

    'Cabeçalho
    
    Codigo.Text = CStr(objRequisicaoCompras.lCodigo)
    Urgente.Value = objRequisicaoCompras.lUrgente
    CodigoOP.Caption = objRequisicaoCompras.sOPCodigo
    
    objRequisitante.lCodigo = objRequisicaoCompras.lRequisitante

    'Le o requisitante para colocar o NomeReduzido na tela
    lErro = CF("Requisitante_Le", objRequisitante)
    If lErro <> SUCESSO And lErro <> 49084 Then gError 54379
    If lErro = 49084 Then gError 54366

    Requisitante.Caption = objRequisitante.sNomeReduzido

    If Len(Trim(objRequisicaoCompras.sCcl)) > 0 Then
    
        'Preenche a CCL
        sCclMascarado = String(STRING_CCL, 0)
    
        lErro = Mascara_MascararCcl(objRequisicaoCompras.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 54380
    
        Ccl.Caption = sCclMascarado
    
    End If
    
    'Tipo de Tributação
    If objRequisicaoCompras.iTipoTributacao <> 0 Then
        
        lErro = CF("TiposTributacao_Le", objRequisicaoCompras.iTipoTributacao, sDescricao)
        If lErro <> SUCESSO And lErro <> 54271 Then gError 54269
        
        'Tipo de Tributação não cadastrada
        If lErro = 54271 Then gError 54270
            
        TipoTributacao.Caption = objRequisicaoCompras.iTipoTributacao & SEPARADOR & sDescricao
    
    End If
    
    'Preenche o Restante dos campos
    
    'Filial Compra
    objFilial.iCodFilial = objRequisicaoCompras.iFilialCompra
    
    lErro = CF("FilialEmpresa_Le", objFilial)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 54431
    If lErro = 27378 Then gError 54432
    
    FilialCompra.Caption = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
    
    'Observação
    If objRequisicaoCompras.lObservacao <> 0 Then
    
        objObservacao.lNumInt = objRequisicaoCompras.lObservacao
        
        'Lê a Observação passada por objObservacao
        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError 54429
        If lErro = 53827 Then gError 54430
        
        Observ.Caption = objObservacao.sObservacao
    
    End If
    
    If objRequisicaoCompras.dtData <> DATA_NULA Then
        Data.Caption = Format(objRequisicaoCompras.dtData, "dd/mm/yyyy")
    Else
        Data.Caption = ""
    End If
    
    If objRequisicaoCompras.dtDataEnvio <> DATA_NULA Then
        DataEnvio.Caption = Format(objRequisicaoCompras.dtDataEnvio, "dd/mm/yyyy")
    Else
        DataEnvio.Caption = ""
    End If
    
    If objRequisicaoCompras.dtDataLimite <> DATA_NULA Then
        DataLimite.Caption = Format(objRequisicaoCompras.dtDataLimite, "dd/mm/yyyy")
    Else
        DataLimite.Caption = ""
    End If
    
    'Preenche TipoDestino e suas Caracteristicas
    lErro = Preenche_TipoDestino(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 54381

    Set gColItensReqCompras = New Collection

    'Preenche o grid com os Itens da requicao
    lErro = Preenche_GridItens(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 54383

    'Preenche o grid com os Itens da requicao
    lErro = Preenche_GridPV(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 178680

    iAlterado = 0

    Traz_RequisicaoEnv_Tela = SUCESSO

    Exit Function

Erro_Traz_RequisicaoEnv_Tela:

    Traz_RequisicaoEnv_Tela = gErr

    Select Case gErr

        Case 54269, 54378, 54379, 54380, 54381, 54383, 54429, 54431, 178680
        
        Case 54270
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIBUTACAO_NAO_CADASTRADA", gErr, objRequisicaoCompras.iTipoTributacao)
            
        Case 54366
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)
        
        Case 54430
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)
        
        Case 54432
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilial.iCodFilial)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173844)

    End Select

    Exit Function

End Function

Function Preenche_TipoDestino(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Preenche o tipo destino e sua caracteristicas

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Preenche_TipoDestino

    TipoDestino.Item(objRequisicaoCompras.iTipoDestino).Value = True

    Select Case objRequisicaoCompras.iTipoDestino

        Case TIPO_DESTINO_EMPRESA

            objFilialEmpresa.iCodFilial = objRequisicaoCompras.iFilialDestino

            'Lê a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then Error 54376
            If lErro = 27378 Then Error 54375

            'Coloca a FilialEmpresa na tela
            FilialEmpresa.Caption = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

        Case TIPO_DESTINO_FORNECEDOR

            objFornecedor.lCodigo = objRequisicaoCompras.lFornCliDestino

            'Lê o fornecedor, seu nome reduzido
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then Error 54384

            'Se não encontrou --> Erro
            If lErro = 12729 Then Error 54368

            'Coloca o NomeReduzido do Fornecedor na tela
            FornecDestino.Caption = objFornecedor.sNomeReduzido

            'Passa o CodFornecedor e o CodFilial para o objfilialfornecedor
            objFilialFornecedor.lCodFornecedor = objRequisicaoCompras.lFornCliDestino
            objFilialFornecedor.iCodFilial = objRequisicaoCompras.iFilialDestino

            'Lê o filialforncedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then Error 54391

            'Se nao encontrou --> Erro
            If lErro = 18272 Then Error 54365

            'Coloca a filial na tela
            FilialFornec.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

    End Select

    Preenche_TipoDestino = SUCESSO

    Exit Function

Erro_Preenche_TipoDestino:

    Preenche_TipoDestino = Err

    Select Case Err

        Case 54364
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 54365
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", Err, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 54367
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case 54368
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)

        Case 54375
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)

        Case 54376, 54384, 54385, 54391, 54412

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173845)

    End Select

    Exit Function

End Function

Function Preenche_GridItens(objRequisicaoCompras As ClassRequisicaoCompras)
'Preenche o Grid com o ItemReqCompras
Dim lErro As Long
Dim objItemReqCompras As ClassItemReqCompras
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objObservacao As New ClassObservacao
Dim objFornecedor As New ClassFornecedor
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iIndice As Integer
Dim iCont As Integer
Dim sProdutoEnxuto As String
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sDescricao As String

On Error GoTo Erro_Preenche_GridItens

   iIndice = 0

    'Preenche GridItens
    For Each objItemReqCompras In objRequisicaoCompras.colItens

        iIndice = iIndice + 1

        sProdutoEnxuto = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(objItemReqCompras.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 54386

        'Coloca o produto na tela
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Preenche o Grid com o ItemReqCompras
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemReqCompras.sDescProduto
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemReqCompras.sUM
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemReqCompras.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col) = Formata_Estoque(objItemReqCompras.dQuantPedida)
        GridItens.TextMatrix(iIndice, iGrid_QuantRecebida_Col) = Formata_Estoque(objItemReqCompras.dQuantRecebida)
        GridItens.TextMatrix(iIndice, iGrid_QuantCancelada_Col) = Formata_Estoque(objItemReqCompras.dQuantCancelada)
        
        'Tipo de Tributação
        If objItemReqCompras.iTipoTributacao <> 0 Then
        
            lErro = CF("TiposTributacao_Le", objItemReqCompras.iTipoTributacao, sDescricao)
            If lErro <> SUCESSO And lErro <> 54271 Then Error 61678
        
            'Tipo de Tributação não cadastrada
            If lErro = 54271 Then Error 61679
            
            GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col) = objItemReqCompras.iTipoTributacao & SEPARADOR & sDescricao
        
        End If
        
        'Se ContaContábil está preenchida
        If Len(Trim(objItemReqCompras.sContaContabil)) > 0 Then

            sContaMascarada = String(STRING_CONTA, 0)
            
            'Formata a ContaContábil
            lErro = Mascara_MascararConta(objItemReqCompras.sContaContabil, sContaMascarada)
            If lErro <> SUCESSO Then Error 54387

            'Coloca ContaContábil na tela
            GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col) = sContaMascarada
        End If

        'Se Fornecedor está preenchido
        If objItemReqCompras.lFornecedor > 0 Then

            objFornecedor.lCodigo = objItemReqCompras.lFornecedor

            'Busca dados do Fornecedor em questão
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then Error 54388
            If lErro = 12729 Then Error 54413
            
            'Coloca o Fornecedor na tela
            GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
        End If

        If objItemReqCompras.iFilial > 0 Then
            objFilialFornecedor.iCodFilial = objItemReqCompras.iFilial
        
            'Lê os dados da Filial do Fornecedor passado por objFornecedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then Error 54421
            If lErro = 18272 Then Error 54422
        
            GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        End If
            
        If objItemReqCompras.lFornecedor <> 0 Then
        
            'Preenche a Combo Exclusivo
            For iCont = 0 To Exclusivo.ListCount - 1
                If Exclusivo.ItemData(iCont) = objItemReqCompras.iExclusivo Then
                    GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) = Exclusivo.List(iCont)
                    Exit For
                End If
            Next
                
        End If
        If objItemReqCompras.lObservacao > 0 Then
            
            objObservacao.lNumInt = objItemReqCompras.lObservacao
            
            'Lê a Observação passada por objObservacao
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then Error 54427
            If lErro = 53827 Then Error 54428
            
            GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao
        End If
        
        If objItemReqCompras.iAlmoxarifado > 0 Then
            objAlmoxarifado.iCodigo = objItemReqCompras.iAlmoxarifado

            'Busca o Almoxarifado em questão
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then Error 54390
            If lErro = 25056 Then Error 54414
        
            'Coloca o Almoxarifado na tela
            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
        End If
        'Se o Centro de Custo está preenchido
        If Len(Trim(objItemReqCompras.sCcl)) > 0 Then
            
            sCclMascarado = String(STRING_CCL, 0)
            'Formata o Centro de Custo
            lErro = Mascara_MascararCcl(objItemReqCompras.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 54392

            'Coloca o Centro de Custo formatado na tela
            GridItens.TextMatrix(iIndice, iGrid_CCL_Col) = sCclMascarado
        End If
        
        'Armazena os números internos dos itens
        gColItensReqCompras.Add objItemReqCompras.lNumIntDoc

    Next

    Call Grid_Refresh_Checkbox(objGridItens)
        
    objGridItens.iLinhasExistentes = gColItensReqCompras.Count

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = Err

    Select Case Err

        Case 54386, 54387, 54388, 54390, 54392, 54421, 54427, 61678

        Case 54413
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)
        
        Case 54414
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)
            
        Case 54422
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
        
        Case 54428
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", Err, objObservacao.lNumInt)
        
        Case 61679
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIBUTACAO_NAO_CADASTRADA", Err, objItemReqCompras.iTipoTributacao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173846)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Finaliza os objEventos
    Set objEventoCodigo = Nothing
    Set objEventoBotaoRequisicoes = Nothing
    Set objEventoPedidoCompras = Nothing
    
    'Libera o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

    Set objGridItens = Nothing
    Set gColItensReqCompras = Nothing
    Set objGridPV = Nothing

End Sub

Private Sub Codigo_LostFocus()
    If Len(Trim(Codigo.Text)) = 0 Then Call Limpa_Tela_ReqComprasEnv
End Sub

Private Sub TipoDestino_Click(Index As Integer)
'Guarda o Tipodestino que foi Escolhido

    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame correspondente a Index visivel
    FrameTipo(Index).Visible = True

    'Torna Frame atual invisivel
    FrameTipo(iFrameTipoDestinoAtual).Visible = False

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

End Sub

''""""""""""""""""""""""""""""""""""""""""""""""
''"  ROTINAS RELACIONADAS AO SISTEMA DE BROWSE "
''""""""""""""""""""""""""""""""""""""""""""""""

Private Sub CodigoLabel_Click()

Dim colSelecao As New Collection
Dim objRequisicaoCompras As New ClassRequisicaoCompras

    'Verifica se o código foi preenchido
    objRequisicaoCompras.lCodigo = StrParaLong(Codigo.Text)

    Call Chama_Tela("ReqComprasEnvLista", colSelecao, objRequisicaoCompras, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRequisicaoCompras = obj1

    'Traz os dados da Requisição para a tela
    lErro = Traz_RequisicaoEnv_Tela(objRequisicaoCompras)
    If lErro <> SUCESSO Then Error 54393

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 54393

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173847)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRequisicoes_Click()

Dim colSelecao As New Collection
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim objRequisitante As New ClassRequisitante
Dim lErro As Long

On Error GoTo Erro_BotaoRequisicoes_Click
    
    'Se o Requisitante estiver preenchido
    If Len(Trim(Requisitante.Caption)) > 0 Then
    
        objRequisitante.sNomeReduzido = Requisitante.Caption
    
        'Lê o Requisitante a partir do seu Nome Reduzido
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then Error 54725
        'Se não encontrou --> Erro
        If lErro = 51152 Then Error 54415
        
        objRequisicaoCompras.lRequisitante = objRequisitante.lCodigo
    
    End If
    
    'Adiciona em colSelecao a giFilialEmpresa
    colSelecao.Add giFilialEmpresa

    'OK ??? Uma tela de lista de Requisições deve receber um objRequisicaoCompras
    'Chama o Browser que Lista os Requisitantes
    Call Chama_Tela("ReqComprasEnvLista", colSelecao, objRequisicaoCompras, objEventoBotaoRequisicoes)
    
    Exit Sub

Erro_BotaoRequisicoes_Click:

    Select Case Err
        
        Case 54415
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO1", Err, objRequisitante.sNomeReduzido)
        
        Case 54725
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173848)
                
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoBotaoRequisicoes_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_objEventoRequisitante_evSelecao

    Set objRequisicaoCompras = obj1

    'Traz os dados da Requisição para a tela
    lErro = Traz_RequisicaoEnv_Tela(objRequisicaoCompras)
    If lErro <> SUCESSO Then Error 54416
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoRequisitante_evSelecao:

    Select Case Err

        Case 54416
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173849)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ReqComprasEnv()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ReqComprasEnv

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
    Urgente.Value = vbUnchecked
    Requisitante.Caption = ""
    Ccl.Caption = ""
    FilialCompra.Caption = ""
    Observ.Caption = ""
    Data.Caption = ""
    DataEnvio.Caption = ""
    DataLimite.Caption = ""
    TipoTributacao.Caption = ""
    CodigoOP.Caption = ""
    
    'Local de Entrega
    FornecDestino.Caption = ""
    FilialFornec.Caption = ""

    FilialEmpresa.Caption = ""

    'Limpa os grids
    Call Grid_Limpa(objGridItens)

    Call Grid_Limpa(objGridPV)

    Set gColItensReqCompras = New Collection

    iAlterado = 0
    
    Exit Sub

Erro_Limpa_Tela_ReqComprasEnv:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173850)

    End Select

    Exit Sub

End Sub

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

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridItens_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173851)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Requisição de Compra Enviada"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ReqComprasEnv"
    
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
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
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
Private Sub FornecedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecedorLabel(Index), Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub FornecDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecDestino, Source, X, Y)
End Sub

Private Sub FornecDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecDestino, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub FilialFornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFornec, Source, X, Y)
End Sub

Private Sub FilialFornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFornec, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresa, Source, X, Y)
End Sub

Private Sub FilialEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresa, Button, Shift, X, Y)
End Sub
Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub DataLimite_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataLimite, Source, X, Y)
End Sub

Private Sub DataLimite_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataLimite, Button, Shift, X, Y)
End Sub

Private Sub FilialCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialCompra, Source, X, Y)
End Sub

Private Sub FilialCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialCompra, Button, Shift, X, Y)
End Sub

Private Sub Ccl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Ccl, Source, X, Y)
End Sub

Private Sub Ccl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Ccl, Button, Shift, X, Y)
End Sub

Private Sub Requisitante_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Requisitante, Source, X, Y)
End Sub

Private Sub Requisitante_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Requisitante, Button, Shift, X, Y)
End Sub

Private Sub Observ_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Observ, Source, X, Y)
End Sub

Private Sub Observ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Observ, Button, Shift, X, Y)
End Sub

Private Sub DataEnvio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEnvio, Source, X, Y)
End Sub

Private Sub DataEnvio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEnvio, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
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

Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
End Sub


Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub CodigoOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoOP, Source, X, Y)
End Sub

Private Sub CodigoOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoOP, Button, Shift, X, Y)
End Sub

Private Sub TipoTributacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoTributacao, Source, X, Y)
End Sub

Private Sub TipoTributacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoTributacao, Button, Shift, X, Y)
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

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
    objGridInt.iLinhasVisiveis = 30

    'Largura da primeira coluna
    GridPV.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridPV = SUCESSO

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
            If lErro <> SUCESSO And lErro <> 28030 Then gError 178706
            
            'Se não encontrou, erro
            If lErro = 28030 Then gError 178705
    
            Set colItensPV = New Collection
            
            lErro = CF("ItemPVItemRC_Le", objItemReqCompra.lNumIntDoc, colItensPV)
            If lErro <> SUCESSO Then gError 178707
            
            For Each objItemPV In colItensPV
            
                iIndice = iIndice + 1
    
                GridPV.TextMatrix(iIndice, iGrid_ItemReqPV_Col) = iItemReq
                GridPV.TextMatrix(iIndice, iGrid_ProdutoPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_Produto_Col)
                GridPV.TextMatrix(iIndice, iGrid_DescProdPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_DescProduto_Col)
                
                Set objItemPV1 = New ClassItemPedido
                
                objItemPV1.iFilialEmpresa = giFilialEmpresa
                objItemPV1.lNumIntDoc = objItemPV.lNumIntDoc
                
                lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV1)
                If lErro <> SUCESSO And lErro <> 178567 Then gError 178708
                
                GridPV.TextMatrix(iIndice, iGrid_CodPV_Col) = objItemPV1.lCodPedido
                GridPV.TextMatrix(iIndice, iGrid_UMPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_UnidadeMed_Col)
                GridPV.TextMatrix(iIndice, iGrid_QuantReqPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_Quantidade_Col)
                
                'Converte a Unidade de Medida do ItemNF para a UM do Item selecionado no GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPV1.sUnidadeMed, objItemReqCompra.sUM, dFator)
                If lErro <> SUCESSO Then gError 178709
                
                GridPV.TextMatrix(iIndice, iGrid_QuantPV_Col) = Formata_Estoque(objItemPV.dQuantSC * dFator)
                
            Next
    
        Next

    Else
    
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = objRequisicaoCompra.sOPCodigo
    
        lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30401 Then gError 178710

        If lErro <> SUCESSO Then
        
            lErro = CF("ItensOP_Baixada_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 178689 Then gError 178711
        
        End If
        
        If lErro = SUCESSO Then
        
            For Each objItemOP In objOrdemProducao.colItens
                
                If objItemOP.lCodPedido <> 0 Then
                    CodigoPV.Caption = objItemOP.lCodPedido
                    Exit For
                End If
                
                If objItemOP.lNumIntDocPai <> 0 Then
                
                    lErro = CF("ItensOP_Le_PV", objItemOP.lNumIntDocPai, lCodigoPV, iFilialPV)
                    If lErro <> SUCESSO And lErro <> 178696 And lErro <> 178697 Then gError 178712
            
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

        Case 178705
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 178706 To 178712

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178713)

    End Select

    Exit Function

End Function

