VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ReqComprasConsOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   8220
      Index           =   3
      Left            =   120
      TabIndex        =   73
      Top             =   780
      Visible         =   0   'False
      Width           =   16620
      Begin VB.TextBox UMPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4515
         MaxLength       =   50
         TabIndex        =   79
         Top             =   1965
         Width           =   705
      End
      Begin VB.TextBox DescProdPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   76
         Top             =   2595
         Width           =   6000
      End
      Begin MSMask.MaskEdBox QuantReqPV 
         Height          =   225
         Left            =   7020
         TabIndex        =   75
         Top             =   2535
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
      Begin MSMask.MaskEdBox CodPV 
         Height          =   225
         Left            =   4650
         TabIndex        =   77
         Top             =   1470
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
         Left            =   525
         TabIndex        =   78
         Top             =   2445
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
         Left            =   7020
         TabIndex        =   80
         Top             =   1950
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
      Begin MSMask.MaskEdBox ProdutoPV 
         Height          =   225
         Left            =   1230
         TabIndex        =   81
         Top             =   1935
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
         Height          =   7695
         Left            =   180
         TabIndex        =   74
         Top             =   195
         Width           =   16395
         _ExtentX        =   28919
         _ExtentY        =   13573
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         Enabled         =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Itens"
      Height          =   8175
      Index           =   2
      Left            =   210
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   16590
      Begin VB.Frame Frame4 
         Caption         =   "Pedidos de Compra"
         Height          =   2970
         Left            =   60
         TabIndex        =   25
         Top             =   5175
         Width           =   7500
         Begin MSMask.MaskEdBox UMPedido 
            Height          =   225
            Left            =   2760
            TabIndex        =   31
            Top             =   1080
            Width           =   1990
            _ExtentX        =   3519
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
         Begin MSMask.MaskEdBox QuantPedido 
            Height          =   225
            Left            =   1560
            TabIndex        =   32
            Top             =   1035
            Width           =   1990
            _ExtentX        =   3519
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
         Begin MSMask.MaskEdBox CodPedido 
            Height          =   225
            Left            =   510
            TabIndex        =   33
            Top             =   990
            Width           =   1870
            _ExtentX        =   3307
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
         Begin MSFlexGridLib.MSFlexGrid GridPC 
            Height          =   1335
            Left            =   150
            TabIndex        =   26
            Top             =   225
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2355
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Notas Fiscais de Entrada"
         Height          =   2970
         Left            =   7920
         TabIndex        =   23
         Top             =   5175
         Width           =   8565
         Begin MSMask.MaskEdBox Serie 
            Height          =   225
            Left            =   435
            TabIndex        =   30
            Top             =   1080
            Width           =   1870
            _ExtentX        =   3307
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
         Begin MSMask.MaskEdBox UMNF 
            Height          =   225
            Left            =   3435
            TabIndex        =   27
            Top             =   1050
            Width           =   1990
            _ExtentX        =   3519
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
         Begin MSMask.MaskEdBox QuantNF 
            Height          =   225
            Left            =   2385
            TabIndex        =   28
            Top             =   1200
            Width           =   1990
            _ExtentX        =   3519
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
         Begin MSMask.MaskEdBox NFiscal 
            Height          =   225
            Left            =   1365
            TabIndex        =   29
            Top             =   1125
            Width           =   1900
            _ExtentX        =   3360
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridNFs 
            Height          =   1305
            Left            =   150
            TabIndex        =   24
            Top             =   240
            Width           =   8220
            _ExtentX        =   14499
            _ExtentY        =   2302
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Itens"
         Height          =   5025
         Left            =   75
         TabIndex        =   6
         Top             =   30
         Width           =   16395
         Begin VB.ComboBox FilialFornecGrid 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   4365
            TabIndex        =   22
            Top             =   2250
            Width           =   1770
         End
         Begin VB.ComboBox Exclusivo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "ReqComprasConsOcx.ctx":0000
            Left            =   5400
            List            =   "ReqComprasConsOcx.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2055
            Width           =   1305
         End
         Begin VB.TextBox ObservacaoGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6210
            MaxLength       =   255
            TabIndex        =   16
            Top             =   2070
            Width           =   2355
         End
         Begin MSMask.MaskEdBox QuantCancelada 
            Height          =   225
            Left            =   6240
            TabIndex        =   18
            Top             =   375
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
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   225
            Left            =   5505
            TabIndex        =   19
            Top             =   315
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
            Left            =   4650
            TabIndex        =   20
            Top             =   360
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
            Left            =   7065
            TabIndex        =   21
            Top             =   360
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
         Begin VB.TextBox ContaContabil 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1425
            MaxLength       =   50
            TabIndex        =   10
            Top             =   2205
            Width           =   1455
         End
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2745
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   375
            Width           =   1065
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   8
            Top             =   450
            Width           =   4000
         End
         Begin MSMask.MaskEdBox TipoTribItem 
            Height          =   255
            Left            =   615
            TabIndex        =   11
            Top             =   1875
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
         Begin MSMask.MaskEdBox CentroCusto 
            Height          =   225
            Left            =   165
            TabIndex        =   12
            Top             =   2175
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
            Left            =   3840
            TabIndex        =   13
            Top             =   435
            Width           =   930
            _ExtentX        =   1640
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
            Left            =   180
            TabIndex        =   14
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
         Begin MSMask.MaskEdBox FornecGrid 
            Height          =   225
            Left            =   2685
            TabIndex        =   15
            Top             =   2235
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
            Height          =   2535
            Left            =   105
            TabIndex        =   7
            Top             =   315
            Width           =   16080
            _ExtentX        =   28363
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Cabeçalho"
      Height          =   8145
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   780
      Width           =   16635
      Begin VB.Frame Frame7 
         Caption         =   "Datas"
         Height          =   795
         Left            =   90
         TabIndex        =   60
         Top             =   5145
         Width           =   10365
         Begin VB.Label Databaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7485
            TabIndex        =   70
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Data Baixa:"
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
            Left            =   6450
            TabIndex        =   69
            Top             =   360
            Width           =   1005
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
            Left            =   4170
            TabIndex        =   66
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label DataLimite 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5235
            TabIndex        =   65
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label DataEnvio 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2940
            TabIndex        =   64
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
            Left            =   1890
            TabIndex        =   63
            Top             =   360
            Width           =   1020
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
            Left            =   150
            TabIndex        =   62
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   660
            TabIndex        =   61
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cabeçalho"
         Height          =   4125
         Left            =   90
         TabIndex        =   34
         Top             =   510
         Width           =   10365
         Begin VB.Frame Frame6 
            Caption         =   "Local de Entrega"
            Height          =   960
            Left            =   240
            TabIndex        =   48
            Top             =   2670
            Width           =   8250
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   675
               Index           =   0
               Left            =   4500
               TabIndex        =   57
               Top             =   210
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
                  TabIndex        =   59
                  Top             =   195
                  Width           =   465
               End
               Begin VB.Label FilialEmpresa 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1050
                  TabIndex        =   58
                  Top             =   165
                  Width           =   2145
               End
            End
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Height          =   675
               Index           =   1
               Left            =   4665
               TabIndex        =   49
               Top             =   195
               Visible         =   0   'False
               Width           =   3495
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
                  TabIndex        =   53
                  Top             =   90
                  Width           =   1035
               End
               Begin VB.Label FornecDestino 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1185
                  TabIndex        =   52
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
                  TabIndex        =   51
                  Top             =   360
                  Width           =   465
               End
               Begin VB.Label FilialFornec 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1185
                  TabIndex        =   50
                  Top             =   360
                  Width           =   2145
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Tipo"
               Height          =   585
               Left            =   195
               TabIndex        =   54
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
                  TabIndex        =   56
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
                  TabIndex        =   55
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1515
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
            Left            =   7920
            TabIndex        =   43
            Top             =   1260
            Width           =   1005
         End
         Begin VB.Label Label8 
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
            Left            =   5430
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   83
            Top             =   1290
            Width           =   960
         End
         Begin VB.Label CodigoPV 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6465
            TabIndex        =   82
            Top             =   1230
            Width           =   1020
         End
         Begin VB.Label FilialRequisicao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3660
            TabIndex        =   72
            Top             =   375
            Width           =   1440
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Filial Empresa:"
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
            Left            =   2385
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   71
            Top             =   435
            Width           =   1245
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
            Left            =   2625
            TabIndex        =   68
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label CodigoOP 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3645
            TabIndex        =   67
            Top             =   1245
            Width           =   915
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
            Left            =   150
            TabIndex        =   47
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label Observ 
            BorderStyle     =   1  'Fixed Single
            Height          =   900
            Left            =   1305
            TabIndex        =   46
            Top             =   1740
            Width           =   7680
         End
         Begin VB.Label TipoTributacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6450
            TabIndex        =   45
            Top             =   855
            Width           =   2505
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
            Left            =   4680
            TabIndex        =   44
            Top             =   915
            Width           =   1695
         End
         Begin VB.Label CclPadraoLabel 
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
            TabIndex        =   42
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label Ccl 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1305
            TabIndex        =   41
            Top             =   1320
            Width           =   915
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
            Left            =   105
            TabIndex        =   40
            Top             =   915
            Width           =   1155
         End
         Begin VB.Label FilialCompra 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1305
            TabIndex        =   39
            Top             =   855
            Width           =   2520
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
            Left            =   5235
            TabIndex        =   38
            Top             =   435
            Width           =   1140
         End
         Begin VB.Label Requisitante 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6450
            TabIndex        =   37
            Top             =   375
            Width           =   2505
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
            Left            =   570
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   36
            Top             =   435
            Width           =   690
         End
         Begin VB.Label Codigo 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1305
            TabIndex        =   35
            Top             =   375
            Width           =   795
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   15750
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   1140
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "ReqComprasConsOcx.ctx":0027
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "ReqComprasConsOcx.ctx":0129
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8640
      Left            =   75
      TabIndex        =   4
      Top             =   420
      Width           =   16860
      _ExtentX        =   29739
      _ExtentY        =   15240
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
Attribute VB_Name = "ReqComprasConsOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim gColItensReqCompras As Collection
Dim gColItemPedCompras As Collection
Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim iChamaTela As Integer

'GridItens
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

'GridNF
Dim objGridNF As AdmGrid
Dim iGrid_Serie_Col As Integer
Dim iGrid_NumeroNF_Col As Integer
Dim iGrid_QuantNF_Col As Integer
Dim iGrid_UMNF_Col As Integer

'GridPedidos
Dim objGridPedidos As AdmGrid
Dim iGrid_NumeroPC_Col As Integer
Dim iGrid_QuantPC_Col As Integer
Dim iGrid_UMPC_Col As Integer

'GridPV
Dim objGridPV As AdmGrid
Dim iGrid_ItemReqPV_Col As Integer
Dim iGrid_ProdutoPV_Col As Integer
Dim iGrid_DescProdPV_Col As Integer
Dim iGrid_CodPV_Col As Integer
Dim iGrid_UMPV_Col As Integer
Dim iGrid_QuantReqPV_Col As Integer
Dim iGrid_QuantPV_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New Collection

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    'Inicializa os ObjEventos
    Set objEventoCodigo = New AdmEvento
    
    'Seta um tipo Padrao
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True

    'Inicializa os Grids
    Set objGridItens = New AdmGrid
    Set objGridNF = New AdmGrid
    Set objGridPedidos = New AdmGrid
    
    Set gColItensReqCompras = New Collection
    Set gColItemPedCompras = New Collection
    
    'Faz a inicialização do GridItens
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 66641

    'Faz a inicialização do GridNF
    lErro = Inicializa_GridNF(objGridNF)
    If lErro <> SUCESSO Then gError 66643

    'Faz a inicialização do GridPedidos
    lErro = Inicializa_GridPedidos(objGridPedidos)
    If lErro <> SUCESSO Then gError 66644

    Set objGridPV = New AdmGrid

    lErro = Inicializa_GridPV(objGridPV)
    If lErro <> SUCESSO Then gError 178664

    'Inicializa mascara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 66642

    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_LIGHT Then
        
        FilialCompra.left = POSICAO_FORA_TELA
        Label3.left = POSICAO_FORA_TELA
        Label3.Visible = False
        FilialEmpresa.left = POSICAO_FORA_TELA
        Label37.left = POSICAO_FORA_TELA
        Label37.Visible = False
        
    End If

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 66641, 66642, 66643, 66644, 178664

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173804)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

Dim colSelecao As New Collection
Dim objRequisicaoCompras As ClassRequisicaoCompras

    'Se no Trata_Parametros nenhuma Requisição de Compras foi passada
    If iChamaTela = 1 Then
        
        'Chama a tela ReqComprasTodasLista
        Call Chama_Tela("ReqComprasTodasLista", colSelecao, objRequisicaoCompras, objEventoCodigo)
        iChamaTela = 0
    
    End If

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long

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
    objGridInt.colCampo.Add (FornecGrid.Name)
    objGridInt.colCampo.Add (FilialFornecGrid.Name)
    objGridInt.colCampo.Add (Exclusivo.Name)
    objGridInt.colCampo.Add (ObservacaoGrid.Name)

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

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE
    objGridItens.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridItens.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridNF(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("U.M.")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Serie.Name)
    objGridInt.colCampo.Add (NFiscal.Name)
    objGridInt.colCampo.Add (QuantNF.Name)
    objGridInt.colCampo.Add (UMNF.Name)

    'Colunas do Grid
    iGrid_Serie_Col = 1
    iGrid_NumeroNF_Col = 2
    iGrid_QuantNF_Col = 3
    iGrid_UMNF_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridNFs
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_NFS_ITEMREQ + 1

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8
    
    'Proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    
    'Largura da primeira coluna
    GridNFs.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridNF = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridPedidos(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("U.M.")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (CodPedido.Name)
    objGridInt.colCampo.Add (QuantPedido.Name)
    objGridInt.colCampo.Add (UMPedido.Name)

    'Colunas do Grid
    iGrid_NumeroPC_Col = 1
    iGrid_QuantPC_Col = 2
    iGrid_UMPC_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridPC

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PEDIDOS_ITEMREQ + 1

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8
    
    'Proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridPedidos = SUCESSO

    Exit Function

End Function

Private Sub Limpa_Tela_ReqComprasCons()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ReqComprasCons

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
    Codigo.Caption = ""
    FilialRequisicao.Caption = ""
    Urgente.Value = vbUnchecked
    Requisitante.Caption = ""
    Ccl.Caption = ""
    FilialCompra.Caption = ""
    Observ.Caption = ""
    Data.Caption = ""
    DataEnvio.Caption = ""
    DataLimite.Caption = ""
    TipoTributacao.Caption = ""
    
    'Local de Entrega
    FornecDestino.Caption = ""
    FilialFornec.Caption = ""

    FilialEmpresa.Caption = ""

    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridNF)
    Call Grid_Limpa(objGridPedidos)
    Call Grid_Limpa(objGridPV)
    
    Set gColItensReqCompras = New Collection
    Set gColItemPedCompras = New Collection
        
    Exit Sub

Erro_Limpa_Tela_ReqComprasCons:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173805)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objRequisicaoCompras As ClassRequisicaoCompras) As Long
    
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma Requisição foi passada por parâmetro
    If Not (objRequisicaoCompras Is Nothing) Then
        
        'Se foi passado o NumIntDoc da Requisição
        If objRequisicaoCompras.lNumIntDoc > 0 Then
        
            'Lê a Requisição Compras a partir de seu número interno
            lErro = CF("RequisicaoComprasTodas_Le", objRequisicaoCompras)
            If lErro = SUCESSO Then

                'Traz os dados da Requisição para a tela
                lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompras)
                If lErro <> SUCESSO Then gError 66645
            
            End If
                    
        'Se foi passado o Código da Requisição
        ElseIf objRequisicaoCompras.lCodigo > 0 Then
            
            'Lê a Requisição Compras a partir do código
            lErro = CF("RequisicaoComprasTodas_Le_Numero", objRequisicaoCompras)
            If lErro <> SUCESSO And lErro <> 76062 Then gError 67642
            
            'Se não encontrou -> Erro
            If lErro = 76062 Then gError 66646
        
            'Traz os dados da Requisição para a tela
            lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompras)
            If lErro <> SUCESSO Then gError 67641

        End If
        
    'Se não foi passado nenhuma Requisição como parâmetro
    Else
                
        'Chama tela de Browse de todas as Requisições de Compras
        iChamaTela = 1
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 66645, 67641, 67642
        
        Case 66646
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objRequisicaoCompras.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173806)

    End Select

    Exit Function

End Function

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objReqCompras As New ClassRequisicaoCompras
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Caption)) = 0 Then gError 76021
    
    objReqCompras.lCodigo = StrParaLong(Codigo.Caption)
    objReqCompras.iFilialEmpresa = Codigo_Extrai(FilialRequisicao.Caption)
    
    'Verifica se a Requisicao de Compra informada existe
    lErro = CF("RequisicaoComprasTodas_Le_Numero", objReqCompras)
    If lErro <> SUCESSO And lErro <> 76062 Then gError 76024
    
    'Se a Requisicao não existe ==> erro
    If lErro = 76062 Then gError 76025
        
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Requisições de Compra Consulta", "REQCOMTO.NumIntDoc = @NREQCOM", 1, "REQCOM", "NREQCOM", objReqCompras.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76026
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 76021
            Call Rotina_Erro(vbOKOnly, "ERRO_REQCOMPRAS_IMPRESSAO", gErr)
            
        Case 76024, 76026
        
        Case 76025
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objReqCompras.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173807)

    End Select
    
    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim colSelecao As New Collection
Dim objRequisicaoCompras As New ClassRequisicaoCompras

    'Verifica se o código foi preenchido
    objRequisicaoCompras.lCodigo = StrParaLong(Codigo.Caption)
        
    'Chama a tela ReqComprasTodasLista
    Call Chama_Tela("ReqComprasTodasLista", colSelecao, objRequisicaoCompras, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRequisicaoCompras = obj1

    'Traz os dados da Requisição para a tela
    lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 66647

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 66647

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173808)

    End Select

    Exit Sub

End Sub

'*****************************************************************************************************************************
''SISTEMA DE SETAS
'*****************************************************************************************************************************

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim sNomeRed As String

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ReqComprasTodas_Requisitante"
    
    ' Move todos os dados Presentes na Tela em objRequisaoModelo
    lErro = Move_Tela_Memoria(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 66648

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
    colCampoValor.Add "DataBaixa", objRequisicaoCompras.dtDataBaixa, 0, "DataBaixa"
    colCampoValor.Add "TipoDestino", objRequisicaoCompras.iTipoDestino, 0, "TipoDestino"
    colCampoValor.Add "FornCliDestino", objRequisicaoCompras.lFornCliDestino, 0, "FornCliDestino"
    colCampoValor.Add "FilialDestino", objRequisicaoCompras.iFilialDestino, 0, "FilialDestino"
    colCampoValor.Add "NumIntDoc", objRequisicaoCompras.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoTributacao", objRequisicaoCompras.iTipoTributacao, 0, "TipoTributacao"
    colCampoValor.Add "OPCodigo", objRequisicaoCompras.sOPCodigo, STRING_BUFFER_MAX_TEXTO, "OPCodigo"
    colCampoValor.Add "NomeReduzido", sNomeRed, STRING_BUFFER_MAX_TEXTO, "NomeReduzido"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 66648

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173809)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objRequisicaoCompras
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
        objRequisicaoCompras.dtDataBaixa = colCampoValor("DataBaixa").vValor
        objRequisicaoCompras.iTipoDestino = colCampoValor.Item("TipoDestino").vValor
        objRequisicaoCompras.lFornCliDestino = colCampoValor.Item("FornCliDestino").vValor
        objRequisicaoCompras.iFilialDestino = colCampoValor.Item("FilialDestino").vValor
        objRequisicaoCompras.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
        objRequisicaoCompras.iTipoTributacao = colCampoValor.Item("TipoTributacao").vValor
        objRequisicaoCompras.sOPCodigo = colCampoValor.Item("OPCodigo").vValor
        
        'Traz os dados da Requisição para a tela
        lErro = Traz_RequisicaoCompra_Tela(objRequisicaoCompras)
        If lErro <> SUCESSO Then gError 66649

    End If
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 66649

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173810)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Move os dados da tela para o objRequisicaoCompra

Dim objRequisitante As New ClassRequisitante
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    'Move o codigo e a descricao
    objRequisicaoCompra.lCodigo = StrParaLong(Codigo.Caption)
    objRequisicaoCompra.iFilialEmpresa = Codigo_Extrai(FilialRequisicao.Caption)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173811)

    End Select

    Exit Function

End Function

Function Traz_RequisicaoCompra_Tela(objRequisicaoCompras As ClassRequisicaoCompras) As Long

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim objFilial As New AdmFiliais
Dim objObservacao As New ClassObservacao
Dim sDescricao As String

On Error GoTo Erro_Traz_RequisicaoCompra_Tela

    'Limpa tela Requisicao Compras
    Call Limpa_Tela_ReqComprasCons

    'Lê os itens da Requisição Compras Enviadas
    lErro = CF("ItensReqComprasTodos_Le", objRequisicaoCompras)
    If lErro <> SUCESSO And lErro <> 54377 Then gError 66650

    'Cabeçalho
    Codigo.Caption = CStr(objRequisicaoCompras.lCodigo)
    
    objFilial.iCodFilial = objRequisicaoCompras.iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilial, True)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 66657
    If lErro = 27378 Then gError 66661
    
    FilialRequisicao.Caption = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
    
    Urgente.Value = objRequisicaoCompras.lUrgente
    CodigoOP.Caption = objRequisicaoCompras.sOPCodigo
    
    If objRequisicaoCompras.lRequisitante <> 0 Then
    
        objRequisitante.lCodigo = objRequisicaoCompras.lRequisitante
    
        'Le o requisitante para colocar o NomeReduzido na tela
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 66652
        If lErro = 49084 Then gError 66659
    
        Requisitante.Caption = objRequisitante.sNomeReduzido
    
    End If
    
    If Len(Trim(objRequisicaoCompras.sCcl)) > 0 Then
    
        'Preenche a CCL
        sCclMascarado = String(STRING_CCL, 0)
    
        lErro = Mascara_MascararCcl(objRequisicaoCompras.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 66653
    
        Ccl.Caption = sCclMascarado
    
    End If
    
    'Tipo de Tributação
    If objRequisicaoCompras.iTipoTributacao <> 0 Then
        
        lErro = CF("TiposTributacao_Le", objRequisicaoCompras.iTipoTributacao, sDescricao)
        If lErro <> SUCESSO And lErro <> 54271 Then gError 66651
        
        'Tipo de Tributação não cadastrada
        If lErro = 54271 Then gError 66658
            
        TipoTributacao.Caption = objRequisicaoCompras.iTipoTributacao & SEPARADOR & sDescricao
    
    End If
    
    'Preenche o Restante dos campos
    
    'Filial Compra
    objFilial.iCodFilial = objRequisicaoCompras.iFilialCompra
    
    lErro = CF("FilialEmpresa_Le", objFilial, True)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 66657
    If lErro = 27378 Then gError 66661
    FilialCompra.Caption = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
    
    
    'Observação
    If objRequisicaoCompras.lObservacao <> 0 Then
    
        objObservacao.lNumInt = objRequisicaoCompras.lObservacao
        
        'Lê a Observação passada por objObservacao
        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError 66656
        If lErro = 53827 Then gError 66660
        
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
    
    If objRequisicaoCompras.dtDataBaixa <> DATA_NULA Then
        DataBaixa.Caption = Format(objRequisicaoCompras.dtDataBaixa, "dd/mm/yyyy")
    Else
        DataBaixa.Caption = ""
    End If
    
    'Preenche TipoDestino e suas Caracteristicas
    lErro = Preenche_TipoDestino(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 66654

    Set gColItensReqCompras = New Collection
    
    'Preenche o grid com os Itens da requicao
    lErro = Preenche_GridItens(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 66655
    
    lErro = Preenche_GridPV(objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 178682
    
    Traz_RequisicaoCompra_Tela = SUCESSO

    Exit Function

Erro_Traz_RequisicaoCompra_Tela:

    Traz_RequisicaoCompra_Tela = gErr

    Select Case gErr

        Case 66650, 66651, 66652, 66653, 66654, 66655, 66656, 66657, 178682
        
        Case 66658
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIBUTACAO_NAO_CADASTRADA", gErr, objRequisicaoCompras.iTipoTributacao)
            
        Case 66659
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)
        
        Case 66660
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)
        
        Case 66661
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilial.iCodFilial)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173812)

    End Select

    Exit Function

End Function

Private Sub TipoDestino_Click(Index As Integer)
'Guarda o Tipodestino que foi Escolhido

    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame correspondente a Index visivel
    FrameTipoDestino(Index).Visible = True

    'Torna Frame atual invisivel
    FrameTipoDestino(iFrameTipoDestinoAtual).Visible = False

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

End Sub

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
            If lErro <> SUCESSO And lErro <> 27378 Then gError 66665
            If lErro = 27378 Then gError 66664

            'Coloca a FilialEmpresa na tela
            FilialEmpresa.Caption = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

        Case TIPO_DESTINO_FORNECEDOR

            objFornecedor.lCodigo = objRequisicaoCompras.lFornCliDestino

            'Lê o fornecedor, seu nome reduzido
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 66666

            'Se não encontrou --> Erro
            If lErro = 12729 Then gError 66663

            'Coloca o NomeReduzido do Fornecedor na tela
            FornecDestino.Caption = objFornecedor.sNomeReduzido

            'Passa o CodFornecedor e o CodFilial para o objfilialfornecedor
            objFilialFornecedor.lCodFornecedor = objRequisicaoCompras.lFornCliDestino
            objFilialFornecedor.iCodFilial = objRequisicaoCompras.iFilialDestino

            'Lê o filialforncedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 66667

            'Se nao encontrou --> Erro
            If lErro = 18272 Then gError 66662

            'Coloca a filial na tela
            FilialFornec.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

    End Select

    Preenche_TipoDestino = SUCESSO

    Exit Function

Erro_Preenche_TipoDestino:

    Preenche_TipoDestino = gErr

    Select Case gErr

        Case 66662
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 66663
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 66664
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 66665, 66666, 66667

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173813)

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
        If lErro <> SUCESSO Then gError 66668

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
            If lErro <> SUCESSO And lErro <> 54271 Then gError 66675
        
            'Tipo de Tributação não cadastrada
            If lErro = 54271 Then gError 66680
            
            GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col) = objItemReqCompras.iTipoTributacao & SEPARADOR & sDescricao
        
        End If
        
        'Se ContaContábil está preenchida
        If Len(Trim(objItemReqCompras.sContaContabil)) > 0 Then

            sContaMascarada = String(STRING_CONTA, 0)
            
            'Formata a ContaContábil
            lErro = Mascara_MascararConta(objItemReqCompras.sContaContabil, sContaMascarada)
            If lErro <> SUCESSO Then gError 66669

            'Coloca ContaContábil na tela
            GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col) = sContaMascarada
        End If

        
        'Se Fornecedor está preenchido
        If objItemReqCompras.lFornecedor > 0 Then

            objFornecedor.lCodigo = objItemReqCompras.lFornecedor

            'Busca dados do Fornecedor em questão
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 66670
            If lErro = 12729 Then gError 66676
            
            'Coloca o Fornecedor na tela
            GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
        End If
        
        'Se a Filial Foi preenchida
        If objItemReqCompras.iFilial > 0 Then
            
            objFilialFornecedor.iCodFilial = objItemReqCompras.iFilial
        
            'Lê os dados da Filial do Fornecedor passado por objFornecedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 66673
            If lErro = 18272 Then gError 66678
        
            GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        End If
            
        'Se o Fornecedor foi preenchido
        If objItemReqCompras.lFornecedor <> 0 Then
        
            'Preenche a Combo Exclusivo
            For iCont = 0 To Exclusivo.ListCount - 1
                If Exclusivo.ItemData(iCont) = objItemReqCompras.iExclusivo Then
                    GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) = Exclusivo.List(iCont)
                    Exit For
                End If
            Next
                
        End If
                
        'Se a observação foi preenchida
        If objItemReqCompras.lObservacao > 0 Then
            objObservacao.lNumInt = objItemReqCompras.lObservacao
            
            'Lê a Observação passada por objObservacao
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError 66674
            If lErro = 53827 Then gError 66679
            
            GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao
        End If
        
        If objItemReqCompras.iAlmoxarifado > 0 Then
            objAlmoxarifado.iCodigo = objItemReqCompras.iAlmoxarifado

            'Busca o Almoxarifado em questão
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 66671
            If lErro = 25056 Then gError 66677
        
            'Coloca o Almoxarifado na tela
            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
        End If
        
        'Se o Centro de Custo está preenchido
        If Len(Trim(objItemReqCompras.sCcl)) > 0 Then
            
            sCclMascarado = String(STRING_CCL, 0)
            'Formata o Centro de Custo
            lErro = Mascara_MascararCcl(objItemReqCompras.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 66672

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

    Preenche_GridItens = gErr

    Select Case gErr

        Case 66668, 66669, 66670, 66671, 66672, 66673, 66674, 66675

        Case 66676
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
        
        Case 66677
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
            
        Case 66678
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
        
        Case 66679
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)
        
        Case 66680
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIBUTACAO_NAO_CADASTRADA", gErr, objItemReqCompras.iTipoTributacao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173814)

    End Select

    Exit Function

End Function

Function Preenche_GridPC(colItemPedido As Collection, iLinhaGridItens As Integer) As Long
'Preenche o GridItensPC

Dim objItemPCInfo As ClassItemPedCompraInfo
Dim iIndice As Integer

On Error GoTo Erro_Preenche_GridPC
    
    'Limpa coleção de Itens de Pedidos de Compras
    Set gColItemPedCompras = New Collection
            
    'Para cada item de Pedido de compras relacionado ao Item de Requisição
    For Each objItemPCInfo In colItemPedido
        
        iIndice = iIndice + 1
        
        'Código do Pedido de Compras
        GridPC.TextMatrix(iIndice, iGrid_NumeroPC_Col) = objItemPCInfo.lPedCompra
        
        'Unidade de Medida
        GridPC.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = GridItens.TextMatrix(iLinhaGridItens, iGrid_UnidadeMed_Col)
        
        'Quantidade
        GridPC.TextMatrix(iIndice, iGrid_QuantPC_Col) = Formata_Estoque(objItemPCInfo.dQuantReceber)
            
        'Adiciona o item na coleção de itensPC
        gColItemPedCompras.Add objItemPCInfo.lNumIntDoc
    
    Next
    
    objGridPedidos.iLinhasExistentes = gColItemPedCompras.Count
    
    Preenche_GridPC = SUCESSO

    Exit Function
    
Erro_Preenche_GridPC:

    Preenche_GridPC = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173815)
    
    End Select
    
    Exit Function
    
End Function

Function Preenche_GridNF(colItemNF As Collection, iLinhaGridPC As Integer) As Long
'Preenche o GridItensPC

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iLinha As Integer
Dim sUM As String

On Error GoTo Erro_Preenche_GridNF
                                        
    'Para cada item de Pedido de compras relacionado ao Item de Requisição
    For Each objItemNF In colItemNF
        
        iIndice = iIndice + 1

        'Número da Nota Fiscal
        GridNFs.TextMatrix(iIndice, iGrid_NumeroNF_Col) = objItemNF.lNumNFOrig

        'Série
        GridNFs.TextMatrix(iIndice, iGrid_Serie_Col) = objItemNF.sSerieNFOrig

        'Procura no GridItens o Produto do ItemNF
        For iLinha = 1 To objGridItens.iLinhasExistentes
        
            'Formata o Produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 66687
            
            'Se encontrou o Produto
            If sProdutoFormatado = objItemNF.sProduto Then
            
                'Guarda a UM do Produto
                sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
            
                Set objProduto = New ClassProduto

                'Lê o Produto
                objProduto.sCodigo = sProdutoFormatado
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 66685
                
                'Se não encontrou, erro
                If lErro = 28030 Then gError 66688
                            
                'Converte a Unidade de Medida do ItemNF para a UM do Item selecionado no GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, sUM, dFator)
                If lErro <> SUCESSO Then gError 66686
        
                'Unidade de Medida
                GridNFs.TextMatrix(iIndice, iGrid_UMNF_Col) = sUM
        
                'Quantidade
                GridNFs.TextMatrix(iIndice, iGrid_QuantNF_Col) = Formata_Estoque(objItemNF.dQuantidade * dFator)
                                           
                Exit For
            
            End If
        
        Next
                        
    Next
    
    objGridNF.iLinhasExistentes = iIndice
    
    Preenche_GridNF = SUCESSO

    Exit Function
    
Erro_Preenche_GridNF:

    Preenche_GridNF = gErr

    Select Case gErr
    
        Case 66685, 66686, 66687
        
        Case 66688
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173816)
    
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

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Finaliza os objEventos
    Set objEventoCodigo = Nothing
    
    'Libera o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gColItensReqCompras = Nothing
    Set gColItemPedCompras = Nothing

    Set objGridItens = Nothing
    Set objGridNF = Nothing
    Set objGridPedidos = Nothing
    Set objGridPV = Nothing

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Requisição de Compra Consulta"
    Call Form_Load

End Function
    
Public Function Name() As String

    Name = "ReqComprasCons"

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

Private Sub FilialEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresa, Source, X, Y)
End Sub

Private Sub FilialEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresa, Button, Shift, X, Y)
End Sub

Private Sub FilialFornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFornec, Source, X, Y)
End Sub

Private Sub FilialFornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFornec, Button, Shift, X, Y)
End Sub

Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

'Tratamento dos Grids
Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer
Dim colItemPedido As New Collection
Dim lErro As Long

On Error GoTo Erro_GridItens_Click

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

    'Se alguma linha foi selecionada
    If GridItens.Row <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then
                
        'Limpa os Grids de Pedidos e NF
        Call Grid_Limpa(objGridPedidos)
        Call Grid_Limpa(objGridNF)
        
        'Lê os Itens de Pedidos de compras associado ao Item de Requisição selecionado
        lErro = CF("ItemRCItemPC_Le", gColItensReqCompras(GridItens.Row), colItemPedido)
        If lErro <> SUCESSO Then gError 66689
                    
        'Preenche o Grid de Itens de Pedido de Compras
        lErro = Preenche_GridPC(colItemPedido, GridItens.Row)
        If lErro <> SUCESSO Then gError 66690
        
    End If

    Exit Sub
    
Erro_GridItens_Click:

    Select Case gErr
    
        Case 66689, 66690
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173817)
    
    End Select
    
    Exit Sub
    
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

'Private Sub GridItens_LeaveCell()
'
'    Call Saida_Celula(objGridItens)
'
'End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridItens_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173818)

    End Select

    Exit Sub

End Sub

'GridNF
Private Sub GridNFs_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If
    
End Sub

Private Sub GridNFs_EnterCell()

    Call Grid_Entrada_Celula(objGridNF, iAlterado)

End Sub

Private Sub GridNFs_GotFocus()

    Call Grid_Recebe_Foco(objGridNF)

End Sub

Private Sub GridNFs_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If

End Sub

'Private Sub GridNFs_LeaveCell()
'
'    Call Saida_Celula(objGridNF)
'
'End Sub

Private Sub GridNFs_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridNF)

End Sub

Private Sub GridNFs_RowColChange()

    Call Grid_RowColChange(objGridNF)

End Sub

Private Sub GridNFs_Scroll()

    Call Grid_Scroll(objGridNF)

End Sub

Private Sub GridNFs_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridNFs_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridNF)

    Exit Sub

Erro_GridNFs_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173819)

    End Select

    Exit Sub

End Sub

Private Sub GridPC_Click()

Dim iExecutaEntradaCelula As Integer
Dim colItemNF As New Collection
Dim lErro As Long

On Error GoTo Erro_GridPC_Click

    Call Grid_Click(objGridPedidos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If
    
    'Se alguma linha foi selecionada
    If GridPC.Row <> 0 And Len(Trim(GridPC.TextMatrix(GridPC.Row, iGrid_NumeroPC_Col))) > 0 Then
        
        'Lê os Itens de Nota Fiscal de compras associados ao Item de Pedido de compras
        lErro = CF("ItemNFItemPC_Le", gColItemPedCompras(GridPC.Row), colItemNF)
        If lErro <> SUCESSO Then gError 66691
            
        'Preenche o Grid de Itens de Nota Fiscal
        lErro = Preenche_GridNF(colItemNF, GridPC.Row)
        If lErro <> SUCESSO Then gError 66692
        
    End If

    Exit Sub
    
Erro_GridPC_Click:

    Select Case gErr
    
        Case 66691, 66692
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173820)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub GridPC_EnterCell()

    Call Grid_Entrada_Celula(objGridPedidos, iAlterado)

End Sub

Private Sub GridPC_GotFocus()

    Call Grid_Recebe_Foco(objGridPedidos)

End Sub

Private Sub GridPC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPedidos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If

End Sub

'Private Sub GridPC_LeaveCell()
'
'    Call Saida_Celula(objGridPedidos)
'
'End Sub

Private Sub GridPC_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridPedidos)

End Sub

Private Sub GridPC_RowColChange()

    Call Grid_RowColChange(objGridPedidos)

End Sub

Private Sub GridPC_Scroll()

    Call Grid_Scroll(objGridPedidos)

End Sub

Private Sub GridPC_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridPC_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridPedidos)

    Exit Sub

Erro_GridPC_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173821)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecedorLabel(Index), Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel(Index), Button, Shift, X, Y)
End Sub


Private Sub DataBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataBaixa, Source, X, Y)
End Sub

Private Sub DataBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
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
            If lErro <> SUCESSO And lErro <> 28030 Then gError 178666
            
            'Se não encontrou, erro
            If lErro = 28030 Then gError 178667
    
            Set colItensPV = New Collection
            
            lErro = CF("ItemPVItemRC_Le", objItemReqCompra.lNumIntDoc, colItensPV)
            If lErro <> SUCESSO Then gError 178668
            
            For Each objItemPV In colItensPV
            
                iIndice = iIndice + 1
    
                GridPV.TextMatrix(iIndice, iGrid_ItemReqPV_Col) = iItemReq
                GridPV.TextMatrix(iIndice, iGrid_ProdutoPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_Produto_Col)
                GridPV.TextMatrix(iIndice, iGrid_DescProdPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_DescProduto_Col)
                
                Set objItemPV1 = New ClassItemPedido
                
                objItemPV1.iFilialEmpresa = giFilialEmpresa
                objItemPV1.lNumIntDoc = objItemPV.lNumIntDoc
                
                lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV1)
                If lErro <> SUCESSO And lErro <> 178567 Then gError 178669
                
                GridPV.TextMatrix(iIndice, iGrid_CodPV_Col) = objItemPV1.lCodPedido
                GridPV.TextMatrix(iIndice, iGrid_UMPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_UnidadeMed_Col)
                GridPV.TextMatrix(iIndice, iGrid_QuantReqPV_Col) = GridItens.TextMatrix(iItemReq, iGrid_Quantidade_Col)
                
                'Converte a Unidade de Medida do ItemNF para a UM do Item selecionado no GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPV1.sUnidadeMed, objItemReqCompra.sUM, dFator)
                If lErro <> SUCESSO Then gError 178670
                
                GridPV.TextMatrix(iIndice, iGrid_QuantPV_Col) = Formata_Estoque(objItemPV.dQuantSC * dFator)
                
            Next
    
        Next

    Else
    
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = objRequisicaoCompra.sOPCodigo
    
        lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30401 Then gError 178700

        If lErro <> SUCESSO Then
        
            lErro = CF("ItensOP_Baixada_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 178689 Then gError 178701
        
        End If
        
        If lErro = SUCESSO Then
        
            For Each objItemOP In objOrdemProducao.colItens
                
                If objItemOP.lCodPedido <> 0 Then
                    CodigoPV.Caption = objItemOP.lCodPedido
                    Exit For
                End If
                
                If objItemOP.lNumIntDocPai <> 0 Then
                
                    lErro = CF("ItensOP_Le_PV", objItemOP.lNumIntDocPai, lCodigoPV, iFilialPV)
                    If lErro <> SUCESSO And lErro <> 178696 And lErro <> 178697 Then gError 178702
            
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

        Case 178666, 178668 To 178670, 178700 To 178702

        Case 178667
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178703)

    End Select

    Exit Function

End Function

