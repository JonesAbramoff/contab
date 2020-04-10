VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ReqComprasBaixadaOcx 
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
   LockControls    =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   9540
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8190
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   90
      Width           =   1140
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "ReqComprasBaixadaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "ReqComprasBaixadaOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cabeçalho"
      Height          =   2625
      Left            =   135
      TabIndex        =   36
      Top             =   645
      Width           =   9225
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
         Left            =   210
         TabIndex        =   55
         Top             =   690
         Width           =   1005
      End
      Begin VB.Frame Frame6 
         Caption         =   "Local de Entrega"
         Height          =   1725
         Left            =   4500
         TabIndex        =   37
         Top             =   780
         Width           =   4395
         Begin VB.Frame FrameEmpresa 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   675
            Left            =   300
            TabIndex        =   52
            Top             =   900
            Width           =   3645
            Begin VB.Label FilialEmpresa 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   960
               TabIndex        =   54
               Top             =   150
               Width           =   2145
            End
            Begin VB.Label Label26 
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
               Left            =   390
               TabIndex        =   53
               Top             =   195
               Width           =   465
            End
         End
         Begin VB.Frame FrameFornecedor 
            BorderStyle     =   0  'None
            Height          =   675
            Left            =   330
            TabIndex        =   47
            Top             =   900
            Width           =   3645
            Begin VB.Label FilialFornec 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1170
               TabIndex        =   51
               Top             =   360
               Width           =   2145
            End
            Begin VB.Label Label22 
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
               TabIndex        =   50
               Top             =   405
               Width           =   465
            End
            Begin VB.Label Fornecedor 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1170
               TabIndex        =   49
               Top             =   0
               Width           =   2145
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
               Left            =   90
               TabIndex        =   48
               Top             =   60
               Visible         =   0   'False
               Width           =   1035
            End
         End
         Begin VB.Frame FrameCliente 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   675
            Left            =   300
            TabIndex        =   42
            Top             =   900
            Width           =   3645
            Begin VB.Label FilialCli 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1230
               TabIndex        =   46
               Top             =   360
               Width           =   2145
            End
            Begin VB.Label Cliente 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1230
               TabIndex        =   45
               Top             =   0
               Width           =   2145
            End
            Begin VB.Label ClienteLabel 
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
               Height          =   195
               Left            =   150
               TabIndex        =   44
               Top             =   60
               Visible         =   0   'False
               Width           =   1035
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
               Left            =   660
               TabIndex        =   43
               Top             =   405
               Width           =   465
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Tipo"
            Height          =   585
            Left            =   180
            TabIndex        =   38
            Top             =   210
            Width           =   4065
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Empresa/Filial"
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
               Left            =   90
               TabIndex        =   41
               Top             =   240
               Value           =   -1  'True
               Width           =   1635
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
               Left            =   1710
               TabIndex        =   40
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Cliente"
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
               Index           =   2
               Left            =   3090
               TabIndex        =   39
               Top             =   240
               Width           =   915
            End
         End
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   71
         Top             =   1410
         Width           =   795
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   70
         Top             =   2250
         Width           =   1095
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
         TabIndex        =   69
         Top             =   1830
         Width           =   1155
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   68
         Top             =   1005
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
         Height          =   195
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   67
         Top             =   330
         Width           =   690
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
         Left            =   2100
         TabIndex        =   66
         Top             =   315
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
         Left            =   6630
         TabIndex        =   65
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Codigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   64
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2670
         TabIndex        =   63
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label DataLimite 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7770
         TabIndex        =   62
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label FilialCompra 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1410
         TabIndex        =   61
         Top             =   1770
         Width           =   2520
      End
      Begin VB.Label Ccl 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1410
         TabIndex        =   60
         Top             =   1350
         Width           =   915
      End
      Begin VB.Label Requisitante 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1410
         TabIndex        =   59
         Top             =   960
         Width           =   2580
      End
      Begin VB.Label Observ 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1410
         TabIndex        =   58
         Top             =   2190
         Width           =   2865
      End
      Begin VB.Label DataEnvio 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5250
         TabIndex        =   57
         Top             =   300
         Width           =   1095
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
         Left            =   4110
         TabIndex        =   56
         Top             =   330
         Width           =   1020
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Itens"
      Height          =   2235
      Left            =   135
      TabIndex        =   2
      Top             =   3375
      Width           =   9225
      Begin VB.TextBox ContaContabil 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1770
         Width           =   1455
      End
      Begin VB.ComboBox Exclusivo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ReqComprasBaixadaOcx.ctx":0280
         Left            =   6390
         List            =   "ReqComprasBaixadaOcx.ctx":028A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1710
         Width           =   1305
      End
      Begin VB.ComboBox FilialForn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4650
         TabIndex        =   10
         Top             =   1710
         Width           =   1770
      End
      Begin VB.TextBox Observacao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   7710
         MaxLength       =   255
         TabIndex        =   9
         Top             =   1740
         Width           =   2355
      End
      Begin VB.ComboBox UM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3390
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox DescProduto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   7
         Top             =   300
         Width           =   1425
      End
      Begin VB.CheckBox CreditaICMS 
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
         Left            =   960
         TabIndex        =   4
         Top             =   1380
         Width           =   945
      End
      Begin VB.CheckBox CreditaIPI 
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
         Left            =   60
         TabIndex        =   3
         Top             =   1410
         Width           =   765
      End
      Begin MSMask.MaskEdBox QuantCancelada 
         Height          =   225
         Left            =   7230
         TabIndex        =   5
         Top             =   300
         Width           =   870
         _ExtentX        =   1535
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
      Begin MSMask.MaskEdBox CentroCusto 
         Height          =   225
         Left            =   -210
         TabIndex        =   6
         Top             =   1800
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
      Begin MSMask.MaskEdBox QuantRececida 
         Height          =   225
         Left            =   6420
         TabIndex        =   13
         Top             =   300
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
         Left            =   5460
         TabIndex        =   14
         Top             =   330
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
         Left            =   8160
         TabIndex        =   15
         Top             =   300
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
         Left            =   4500
         TabIndex        =   16
         Top             =   300
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
         Left            =   570
         TabIndex        =   17
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fornec 
         Height          =   225
         Left            =   2670
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
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
         Height          =   1065
         Left            =   150
         TabIndex        =   19
         Top             =   360
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   1879
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Conta Contabil"
         Height          =   195
         Left            =   1110
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   195
         Left            =   8100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   2010
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Exclusividade"
         Height          =   195
         Left            =   6450
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   " Filial Fornecedor"
         Height          =   195
         Left            =   4830
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   2010
         Width           =   1200
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor         "
         Height          =   195
         Left            =   3390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   31
         Top             =   2010
         Width           =   945
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado"
         Height          =   195
         Left            =   8190
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Recebida"
         Height          =   195
         Left            =   6420
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   150
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Em Pedido"
         Height          =   225
         Left            =   5520
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   150
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Unidade Med"
         Height          =   195
         Left            =   3390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   90
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         Height          =   195
         Left            =   4530
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Produto         "
         Height          =   195
         Left            =   810
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   90
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   2100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Centro Custo"
         Height          =   195
         Left            =   60
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   1980
         Width           =   915
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Cancelada"
         Height          =   195
         Left            =   7260
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Credita IPI"
         Height          =   195
         Left            =   60
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   1650
         Width           =   735
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Credita ICMS"
         Height          =   195
         Left            =   930
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   1590
         Width           =   930
      End
   End
   Begin VB.CommandButton BotaoRequisicoes 
      Caption         =   "Requisições Baixadas"
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
      Left            =   105
      TabIndex        =   1
      Top             =   5865
      Width           =   2235
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
      Height          =   285
      Left            =   7965
      TabIndex        =   0
      Top             =   5865
      Width           =   1395
   End
End
Attribute VB_Name = "ReqComprasBaixadaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
'Private Sub BotaoPedidos_Click()
'    'Analisa o ítem da Requisição no Grid.
'    'Se houver 1 Pedido associado, chama ele (novo ou enviado ou baixado).
'    'Se houver + de 1 abre um Browse de Pedidos associados. (passa como parametro a Requisição).
'        'Selecionado o Pedido chama a tela de Pedido Compra (novo ou enviado ou baixado).
'End Sub




''**** inicio do trecho a ser copiado *****
'
'Public Function Form_Load_Ocx() As Object
'
'    Set Form_Load_Ocx = Me
'    Caption = "Requisição de Compras Baixada"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "ReqComprasBaixada"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property

'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
'
'   RaiseEvent Unload
'
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''**** fim do trecho a ser copiado *****

Private Sub FilialEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresa, Source, X, Y)
End Sub

Private Sub FilialEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresa, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub FilialFornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFornec, Source, X, Y)
End Sub

Private Sub FilialFornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFornec, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Fornecedor, Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub FilialCli_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialCli, Source, X, Y)
End Sub

Private Sub FilialCli_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialCli, Button, Shift, X, Y)
End Sub

Private Sub Cliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cliente, Source, X, Y)
End Sub

Private Sub Cliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cliente, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Codigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Codigo, Button, Shift, X, Y)
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

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
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

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

