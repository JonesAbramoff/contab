VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RecebMaterialFCom 
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4470
      Index           =   1
      Left            =   195
      TabIndex        =   43
      Top             =   780
      Width           =   9204
      Begin VB.Frame Frame11 
         Caption         =   "Entrada"
         Height          =   885
         Left            =   3495
         TabIndex        =   67
         Top             =   90
         Width           =   5550
         Begin MSComCtl2.UpDown UpDownEntrada 
            Height          =   300
            Left            =   2745
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntrada 
            Height          =   300
            Left            =   1665
            TabIndex        =   9
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraEntrada 
            Height          =   300
            Left            =   4500
            TabIndex        =   10
            Top             =   315
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora Entrada:"
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
            Index           =   1
            Left            =   3255
            TabIndex        =   115
            Top             =   375
            Width           =   1200
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Data Entrada:"
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
            Left            =   405
            TabIndex        =   69
            Top             =   375
            Width           =   1200
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Identificação"
         Height          =   885
         Index           =   0
         Left            =   195
         TabIndex        =   63
         Top             =   90
         Width           =   3075
         Begin VB.CommandButton BotaoLimparRec 
            Height          =   315
            Left            =   2040
            Picture         =   "RecebMaterialFCom.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Limpar Código"
            Top             =   345
            Width           =   345
         End
         Begin VB.Label NumRecebimento 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1260
            TabIndex        =   66
            Top             =   345
            Width           =   765
         End
         Begin VB.Label LabelRecebimento 
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
            Left            =   495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   65
            Top             =   405
            Width           =   660
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Nota Fiscal"
         Height          =   1035
         Index           =   1
         Left            =   195
         TabIndex        =   57
         Top             =   1005
         Width           =   8865
         Begin VB.Frame FrameNFForn 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   3555
            TabIndex        =   82
            Top             =   240
            Width           =   4500
            Begin VB.ComboBox Serie 
               Height          =   315
               Left            =   1185
               TabIndex        =   1
               Top             =   225
               Width           =   765
            End
            Begin MSMask.MaskEdBox NFiscal 
               Height          =   300
               Left            =   3360
               TabIndex        =   2
               Top             =   217
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin VB.Label NFiscalLabel 
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
               Left            =   2580
               TabIndex        =   84
               Top             =   240
               Width           =   720
            End
            Begin VB.Label SerieLabel 
               AutoSize        =   -1  'True
               Caption         =   "Série:"
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
               TabIndex        =   83
               Top             =   270
               Width           =   510
            End
         End
         Begin VB.Frame FrameNFPropria 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   3630
            TabIndex        =   85
            Top             =   285
            Visible         =   0   'False
            Width           =   4500
         End
         Begin VB.OptionButton NFiscalPropria 
            Caption         =   "Nota Fiscal Própria"
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
            Left            =   945
            TabIndex        =   11
            Top             =   315
            Width           =   2025
         End
         Begin VB.OptionButton NFiscalForn 
            Caption         =   "Nota Fiscal do Fornecedor"
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
            Left            =   930
            TabIndex        =   0
            Top             =   705
            Width           =   2700
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Fornecedor"
         Height          =   855
         Left            =   165
         TabIndex        =   54
         Top             =   2085
         Width           =   8865
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   6090
            TabIndex        =   4
            Top             =   345
            Width           =   1860
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   2340
            TabIndex        =   3
            Top             =   330
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1215
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   56
            Top             =   375
            Width           =   1035
         End
         Begin VB.Label Label3 
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
            Left            =   5535
            TabIndex        =   55
            Top             =   375
            Width           =   465
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Pedidos de Compra"
         Height          =   1464
         Left            =   180
         TabIndex        =   51
         Top             =   2952
         Width           =   8865
         Begin VB.ListBox PedidosCompra 
            Height          =   735
            Left            =   4752
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   456
            Width           =   2052
         End
         Begin VB.ComboBox FilialCompra 
            Height          =   288
            ItemData        =   "RecebMaterialFCom.ctx":0532
            Left            =   1560
            List            =   "RecebMaterialFCom.ctx":0534
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   660
            Width           =   2604
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
            Height          =   564
            Left            =   6972
            Picture         =   "RecebMaterialFCom.ctx":0536
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   816
            Width           =   1812
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
            Height          =   564
            Left            =   6972
            Picture         =   "RecebMaterialFCom.ctx":1718
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   216
            Width           =   1812
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Pedidos de Compra"
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
            Height          =   192
            Left            =   4740
            TabIndex        =   53
            Top             =   204
            Width           =   1656
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   300
            TabIndex        =   52
            Top             =   690
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      Height          =   4455
      Index           =   3
      Left            =   180
      TabIndex        =   45
      Top             =   810
      Visible         =   0   'False
      Width           =   9156
      Begin VB.ComboBox Moeda 
         Enabled         =   0   'False
         Height          =   288
         Left            =   4728
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   144
         Width           =   1665
      End
      Begin MSMask.MaskEdBox Taxa 
         Height          =   312
         Left            =   7932
         TabIndex        =   118
         Top             =   144
         Width           =   1092
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox ComboPedidoCompras 
         Height          =   288
         Left            =   1884
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   144
         Width           =   1332
      End
      Begin VB.Frame Frame10 
         Caption         =   "Itens de Pedidos de Compra"
         Height          =   3300
         Index           =   0
         Left            =   60
         TabIndex        =   49
         Top             =   552
         Width           =   9072
         Begin MSMask.MaskEdBox TaxaGrid 
            Height          =   228
            Left            =   1980
            TabIndex        =   126
            Top             =   2412
            Visible         =   0   'False
            Width           =   1056
            _ExtentX        =   1852
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
         Begin VB.TextBox MoedaGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   684
            MaxLength       =   50
            TabIndex        =   125
            Top             =   2412
            Width           =   1200
         End
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   228
            Left            =   7092
            TabIndex        =   124
            Top             =   1692
            Visible         =   0   'False
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
         Begin MSMask.MaskEdBox ValorRecebido 
            Height          =   228
            Left            =   7812
            TabIndex        =   123
            Top             =   1980
            Visible         =   0   'False
            Width           =   1056
            _ExtentX        =   1852
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
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2664
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1980
            Width           =   1344
         End
         Begin MSMask.MaskEdBox ItemPC 
            Height          =   228
            Left            =   1188
            TabIndex        =   27
            Top             =   1980
            Width           =   468
            _ExtentX        =   820
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   228
            Left            =   6336
            TabIndex        =   32
            Top             =   1980
            Width           =   1056
            _ExtentX        =   1879
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
         Begin MSMask.MaskEdBox UM 
            Height          =   228
            Left            =   4032
            TabIndex        =   30
            Top             =   1980
            Width           =   1008
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
         Begin MSMask.MaskEdBox QuantAReceber 
            Height          =   228
            Left            =   5112
            TabIndex        =   31
            Top             =   1980
            Width           =   1056
            _ExtentX        =   1852
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
         Begin MSMask.MaskEdBox Prod 
            Height          =   228
            Left            =   1680
            TabIndex        =   28
            Top             =   1980
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoPC 
            Height          =   216
            Left            =   216
            TabIndex        =   26
            Top             =   1980
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   370
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItensPC 
            Height          =   2688
            Left            =   108
            TabIndex        =   120
            Top             =   264
            Width           =   8868
            _ExtentX        =   15637
            _ExtentY        =   4736
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoPedidoCompra 
         Caption         =   "Pedido de Compra"
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
         Left            =   6780
         TabIndex        =   50
         Top             =   3975
         Width           =   2265
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Moeda:"
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
         Height          =   192
         Left            =   4020
         TabIndex        =   122
         Top             =   204
         Width           =   648
      End
      Begin VB.Label LabelTaxa 
         AutoSize        =   -1  'True
         Caption         =   "Taxa:"
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
         Left            =   7392
         TabIndex        =   121
         Top             =   204
         Width           =   492
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pedido de Compras:"
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
         Left            =   132
         TabIndex        =   119
         Top             =   204
         Width           =   1716
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4365
      Index           =   4
      Left            =   180
      TabIndex        =   46
      Top             =   825
      Visible         =   0   'False
      Width           =   9204
      Begin VB.Frame Frame8 
         Caption         =   "Volumes"
         Height          =   885
         Left            =   150
         TabIndex        =   96
         Top             =   1380
         Width           =   8910
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5500
            TabIndex        =   100
            Top             =   338
            Width           =   1335
         End
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3300
            TabIndex        =   99
            Top             =   338
            Width           =   1335
         End
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7320
            MaxLength       =   20
            TabIndex        =   97
            Top             =   345
            Width           =   1440
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1275
            TabIndex        =   98
            Top             =   345
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label30 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   104
            Top             =   398
            Width           =   1050
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
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
            Left            =   2505
            TabIndex        =   103
            Top             =   398
            Width           =   750
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
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
            Left            =   4850
            TabIndex        =   102
            Top             =   398
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº :"
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
            Left            =   6960
            TabIndex        =   101
            Top             =   398
            Width           =   345
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados de Transporte"
         Height          =   1215
         Left            =   150
         TabIndex        =   86
         Top             =   120
         Width           =   8910
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   4920
            TabIndex        =   92
            Top             =   315
            Width           =   2205
         End
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   7965
            TabIndex        =   91
            Top             =   765
            Width           =   735
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   5835
            MaxLength       =   10
            TabIndex        =   90
            Top             =   765
            Width           =   1290
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frete por conta"
            Height          =   795
            Index           =   1
            Left            =   465
            TabIndex        =   87
            Top             =   270
            Width           =   2220
            Begin VB.OptionButton Destinatario 
               Caption         =   "Destinatário"
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
               Left            =   405
               TabIndex        =   89
               Top             =   495
               Width           =   1695
            End
            Begin VB.OptionButton Emitente 
               Caption         =   "Emitente"
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
               Left            =   405
               TabIndex        =   88
               Top             =   225
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "U.F. :"
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
            Left            =   7395
            TabIndex        =   95
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Placa Veículo:"
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
            Left            =   4485
            TabIndex        =   94
            Top             =   810
            Width           =   1275
         End
         Begin VB.Label TransportadoraLabel 
            AutoSize        =   -1  'True
            Caption         =   "Transportadora:"
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
            Left            =   3480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   93
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Complemento"
         Height          =   1995
         Left            =   128
         TabIndex        =   58
         Top             =   2340
         Width           =   8895
         Begin VB.TextBox Observacao 
            Height          =   300
            Left            =   2175
            MaxLength       =   40
            TabIndex        =   36
            Top             =   1380
            Width           =   4755
         End
         Begin VB.TextBox Mensagem 
            Height          =   300
            Left            =   2175
            MaxLength       =   250
            TabIndex        =   33
            Top             =   450
            Width           =   4755
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   300
            Left            =   2175
            TabIndex        =   34
            Top             =   960
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   300
            Left            =   5640
            TabIndex        =   35
            Top             =   960
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label16 
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
            Left            =   930
            TabIndex        =   62
            Top             =   1485
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem N.Fiscal:"
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
            TabIndex        =   61
            Top             =   450
            Width           =   1725
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto:"
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
            Left            =   4575
            TabIndex        =   60
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Peso Líquido:"
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
            Left            =   825
            TabIndex        =   59
            Top             =   1020
            Width           =   1200
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4365
      Index           =   2
      Left            =   180
      TabIndex        =   44
      Top             =   885
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   840
         Left            =   90
         TabIndex        =   71
         Top             =   3120
         Width           =   9090
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1635
            TabIndex        =   22
            Top             =   465
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   4572
            TabIndex        =   24
            Top             =   465
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   3093
            TabIndex        =   23
            Top             =   465
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   -20000
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   465
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IPIValor1 
            Height          =   285
            Left            =   135
            TabIndex        =   21
            Top             =   465
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
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
            Left            =   -20000
            TabIndex        =   81
            Top             =   255
            Width           =   825
         End
         Begin VB.Label Total 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7530
            TabIndex        =   80
            Top             =   465
            Width           =   1410
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Seguro"
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
            Left            =   3491
            TabIndex        =   79
            Top             =   255
            Width           =   615
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
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
            Left            =   2094
            TabIndex        =   78
            Top             =   240
            Width           =   450
         End
         Begin VB.Label SubTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6051
            TabIndex        =   77
            Top             =   465
            Width           =   1410
         End
         Begin VB.Label LabelTotais 
            AutoSize        =   -1  'True
            Caption         =   "Valor Produtos"
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
            Left            =   6126
            TabIndex        =   76
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   8010
            TabIndex        =   75
            Top             =   255
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
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
            Left            =   4857
            TabIndex        =   74
            Top             =   255
            Width           =   840
         End
         Begin VB.Label LabelIPIValor 
            AutoSize        =   -1  'True
            Caption         =   "IPI"
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
            Left            =   713
            TabIndex        =   73
            Top             =   255
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Itens"
         Height          =   3030
         Left            =   120
         TabIndex        =   47
         Top             =   60
         Width           =   9060
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1050
            Width           =   2250
         End
         Begin VB.ComboBox Produto 
            Height          =   315
            Left            =   330
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   885
            Width           =   1245
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   885
            Width           =   720
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   225
            Left            =   5775
            TabIndex        =   16
            Top             =   825
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   5310
            TabIndex        =   19
            Top             =   1710
            Width           =   1155
            _ExtentX        =   2037
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   5100
            TabIndex        =   18
            Top             =   1200
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorUnitario 
            Height          =   225
            Left            =   6855
            TabIndex        =   17
            Top             =   555
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   10
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   4620
            TabIndex        =   15
            Top             =   765
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   10
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
         Begin MSMask.MaskEdBox ValorTotal 
            Height          =   255
            Left            =   6900
            TabIndex        =   20
            Top             =   1365
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2595
            Left            =   150
            TabIndex        =   48
            Top             =   255
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   4577
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoCcls 
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
         Height          =   345
         Left            =   7305
         TabIndex        =   25
         Top             =   4005
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4365
      Index           =   5
      Left            =   180
      TabIndex        =   105
      Top             =   840
      Visible         =   0   'False
      Width           =   9240
      Begin VB.CommandButton BotaoLocalizacaoDist 
         Caption         =   "Estoque"
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
         Left            =   6960
         TabIndex        =   114
         Top             =   3885
         Width           =   1365
      End
      Begin VB.Frame Frame12 
         Caption         =   "Distribuição dos Produtos"
         Height          =   3465
         Left            =   300
         TabIndex        =   106
         Top             =   210
         Width           =   8370
         Begin MSMask.MaskEdBox UMDist 
            Height          =   225
            Left            =   4425
            TabIndex        =   107
            Top             =   120
            Visible         =   0   'False
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoAlmoxDist 
            Height          =   225
            Left            =   1740
            TabIndex        =   108
            Top             =   135
            Visible         =   0   'False
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
         Begin MSMask.MaskEdBox AlmoxDist 
            Height          =   225
            Left            =   3060
            TabIndex        =   109
            Top             =   135
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
         Begin MSMask.MaskEdBox QuantDist 
            Height          =   225
            Left            =   6540
            TabIndex        =   110
            Top             =   105
            Width           =   1470
            _ExtentX        =   2593
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
         Begin MSMask.MaskEdBox ItemNFDist 
            Height          =   225
            Left            =   1005
            TabIndex        =   111
            Top             =   105
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDist 
            Height          =   2910
            Left            =   360
            TabIndex        =   112
            Top             =   345
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   5133
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox QuantItemNFDist 
            Height          =   225
            Left            =   5025
            TabIndex        =   113
            Top             =   150
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
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
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6675
      ScaleHeight     =   495
      ScaleWidth      =   2610
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   75
      Width           =   2670
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2100
         Picture         =   "RecebMaterialFCom.ctx":2732
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1590
         Picture         =   "RecebMaterialFCom.ctx":28B0
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1080
         Picture         =   "RecebMaterialFCom.ctx":2DE2
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   585
         Picture         =   "RecebMaterialFCom.ctx":2F6C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   90
         Picture         =   "RecebMaterialFCom.ctx":30C6
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4860
      Left            =   135
      TabIndex        =   42
      Top             =   465
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8573
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos Compra"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuição"
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
Attribute VB_Name = "RecebMaterialFCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:

Dim gbLimpaTaxa As Boolean

Dim m_Caption As String
Event Unload()

'Declaração das Variáveis Globais
Public iAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iFrameAtual As Integer
Dim sFornecedorAnterior As String
Dim iFilialCompraAnterior As Integer
Dim iFilialAnterior As Integer
Dim gcolItemPedCompraInfo As Collection
Dim gcolPedidoCompra As Collection
Dim gColInfoBD As Collection
Dim gbCarregandoTela As Boolean

'GridItensPC
Dim objGridItensPC As AdmGrid
Dim iGrid_PedCompra_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_Prod_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_AReceber_Col As Integer
Dim iGrid_Recebido_Col As Integer
Dim iGrid_Recebido_RS_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
Dim iGrid_Moeda_Col As Integer
Dim iGrid_Taxa_Col As Integer

'distribuicao
Public gobjDistribuicao As Object

'GridItens
Dim objGrid As AdmGrid
Public objGridItens As AdmGrid

Public iGrid_Produto_Col As Integer
Public iGrid_Descricao_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
'distribuicao
'Public iGrid_Almoxarifado_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorTotal_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Ccl_Col As Integer

Dim WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Dim WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Dim WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Dim WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Dim WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Dim WithEvents objEventoRecebimento As AdmEvento
Attribute objEventoRecebimento.VB_VarHelpID = -1

Public Property Get GridItens() As Object
     Set GridItens = Me.Controls("GridItens")
End Property

Private Sub BotaoDesmarcarTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To PedidosCompra.ListCount - 1
        PedidosCompra.Selected(iIndice) = False
    Next
    
    ComboPedidoCompras.ListIndex = 0

End Sub

Private Sub BotaoLimparRec_Click()
    
    NumRecebimento.Caption = ""

End Sub

Private Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To PedidosCompra.ListCount - 1
        PedidosCompra.Selected(iIndice) = True
    Next
    
    ComboPedidoCompras.ListIndex = 0

End Sub

Private Sub BotaoPedidoCompra_Click()

Dim objPedidoCompra As New ClassPedidoCompras
Dim lErro As Long

On Error GoTo Erro_BotaoPedidoCompra_Click

    'Se nenhuma linha do Grid de Pedido de Compras foi selecionada, sai da rotina
    If GridItensPC.Row = 0 Then gError 89285

    objPedidoCompra.lCodigo = StrParaLong(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_PedCompra_Col))
    objPedidoCompra.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)

    'Chama a tela "PedComprasCons"
    Call Chama_Tela("PedComprasCons", objPedidoCompra)

    Exit Sub
    
Erro_BotaoPedidoCompra_Click:

    Select Case gErr
    
        Case 89285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166401)
    
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    'Se nenhuma linha foi selecionada -> Erro
    If GridItens.Row = 0 Then Error 54440

    'Se o produto da linha selecionad não foi preenchido --> Erro
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then Error 54441

    'Chama a tela "CclLista"
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub

Erro_BotaoCcls_Click:

    Select Case Err

        Case 54440
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 54441
             lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166402)

    End Select

    Exit Sub

End Sub

Private Sub ComboPedidoCompras_Click()

Dim lErro As Long

On Error GoTo Erro_ComboPedidoCompras_Click

    Moeda.ListIndex = -1
    
    gbLimpaTaxa = True
    
    'Preenche o grid
    lErro = Preenche_GridItensPC
    If lErro <> SUCESSO Then gError 108992
    
    Exit Sub
    
Erro_ComboPedidoCompras_Click:

    Select Case gErr
        
        Case 108992
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166403)
    
    End Select
    
End Sub

Private Sub DataEntrada_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEntrada, iAlterado)
    
End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

    'Verifica se alguma filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub
    
    If Len(Trim(Filial.Text)) > 0 Then
    
        If sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
        Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text) Then
    
            Call Atualiza_ListaPedidos
    
        End If
    
    End If
    
    Call Trata_FilialForn
    
End Sub

Private Function Trata_FilialForn() As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_Trata_FilialForn

    objFilialForn.iCodFilial = Codigo_Extrai(Filial.Text)

    If objFilialForn.iCodFilial <> 0 Then

        'Lê a Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialForn)
        If lErro <> SUCESSO Then gError 65589

    End If

    Trata_FilialForn = SUCESSO

    Exit Function

Erro_Trata_FilialForn:

    Trata_FilialForn = gErr

    Select Case gErr

        Case 65589

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166404)

    End Select

    Exit Function

End Function

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then
        Call Limpa_Tela_PC
        Exit Sub
    End If
    
    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 54487

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 54491

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 54488

        If lErro = 18272 Then gError 54489

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
    
    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 54490

    If sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
    Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text) Then

        Call Atualiza_ListaPedidos

    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 54487, 54488

        Case 54489
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
        
        Case 54490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 54491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166405)

    End Select

    Exit Sub

End Sub

Private Sub FilialCompra_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_GotFocus()

    iFilialCompraAnterior = Codigo_Extrai(FilialCompra.Text)

End Sub

Private Sub LabelRecebimento_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelRecebimento_Click

    'Se o Recebimento estiver preenchido
    If Len(Trim(NumRecebimento.Caption)) > 0 Then
        objNFiscal.lNumRecebimento = CLng(NumRecebimento.Caption)
    Else
        objNFiscal.lNumRecebimento = 0
    End If
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 89250

        'Se não achou o Fornecedor --> erro
        If lErro = 6681 Then gError 89251

        objNFiscal.lFornecedor = objFornecedor.lCodigo

    End If

    If Len(Trim(Filial.Text)) <> 0 Then objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    If NFiscalPropria.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRFPCO
    ElseIf NFiscalForn.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRFFCO
    Else
        objNFiscal.iTipoNFiscal = 0
    End If

    objNFiscal.sSerie = Serie.Text

    If Len(Trim(NFiscal.Text)) <> 0 Then objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)

    objNFiscal.dtDataEntrada = MaskedParaDate(DataEntrada)

    'Chama a tela de browse RecebMaterialFLista
    Call Chama_Tela("RecebMaterialFComLista", colSelecao, objNFiscal, objEventoRecebimento)

    Exit Sub

Erro_LabelRecebimento_Click:

    Select Case gErr
    
        Case 89250

        Case 89251
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166406)
          
    End Select

    Exit Sub

End Sub

Private Sub NFiscal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
    
End Sub

Private Sub NFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscal_Validate
    
    'Se o número foi preenchido
    If Len(Trim(NFiscal)) > 0 Then
    
        lErro = Valor_Positivo_Critica(NFiscal.Text)
        If lErro <> SUCESSO Then gError 67879
    
    End If
    
    Exit Sub
    
Erro_NFiscal_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 67879
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166407)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoRecebimento_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoRecebimento_evSelecao

    Set objNFiscal = obj1

    'Lê NFiscal no BD
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 89246
    
    If lErro = 31442 Then gError 89247

    'Coloca a Nota Fiscal na Tela1
    lErro = Traz_RecebMaterialFCom_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 66993

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Total_Calcula
    
    Me.Show

    Exit Sub

Erro_objEventoRecebimento_evSelecao:

    Select Case gErr

        Case 66993, 89246

        Case 89247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_NAO_CADASTRADO", gErr, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166408)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) <> 0 And GridItens.Row <> 0 Then

        lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 54442
        
        'Preenche o campo Ccl com o Ccl encontrado
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
        sCclMascarado = Ccl.Text

        'Coloca o valor do Ccl na coluna correspondente
        GridItens.TextMatrix(GridItens.Row, iGrid_Ccl_Col) = sCclMascarado

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 54442 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166409)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o  Numero do Recebimento não está preenchido, Erro
    If Len(Trim(NumRecebimento.Caption)) = 0 Then Error 54443

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Error 54444

    'Confirma exclusão?
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_RECEBIMENTO", CLng(NumRecebimento.Caption))

    'Se resposta for sim
    If vbMsg = vbYes Then

        'Chama a rotina de Exclusão
        lErro = CF("RecebMaterialFCom_Exclui", objNFiscal)
        If lErro <> SUCESSO Then Error 54445

        'Limpa a tela
        Call Limpa_Tela_RecebMaterialFCom

    End If

    GL_objMDIForm.MousePointer = vbDefault
        
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 54443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_PREENCHIDO", Err)
        
        Case 54444, 54445

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166410)

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

    ComboPedidoCompras.ListIndex = 0
    
    'Chama a função Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 54479

    'Limpa a tela
    Call Limpa_Tela_RecebMaterialFCom

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 54479

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166411)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 54480

    'Limpa a tela
    Call Limpa_Tela_RecebMaterialFCom
    
    gbCarregandoTela = False

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 54480

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166412)

    End Select

    Exit Sub

End Sub

Private Sub Filial_GotFocus()

    iFilialAnterior = Codigo_Extrai(Filial.Text)

End Sub

Private Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCompra_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialCompra.Text)) = 0 Then
        Call Limpa_Tela_PC
        Exit Sub
    End If
    
    'Verifica se é uma FilialEmpresa selecionada
    If FilialCompra.Text = FilialCompra.List(FilialCompra.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCompra, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 61908

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 61909

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 61910

        'coloca na tela
        FilialCompra.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome
    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 61911

    If sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> Codigo_Extrai(Filial.Text) Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text) Then
        
           Call Atualiza_ListaPedidos

    End If
        
    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case gErr

        Case 61909, 61908
            
        Case 61910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialCompra.Text)

        Case 61911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialCompra.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166413)

    End Select

    Exit Sub

End Sub

Public Sub Limpa_Tela_PC()
'Limpa da tela os Pedidos de Compras

    'Limpa os Grids
    Call Grid_Limpa(objGrid)
    Call Grid_Limpa(objGridItensPC)

    'Limpa a Lista de Pedidos de Compra
    PedidosCompra.Clear

    'Limpa as coleções
    Set gcolPedidoCompra = New Collection
    Set gcolItemPedCompraInfo = New Collection
    
    ComboPedidoCompras.ListIndex = 0
    
    PedidosCompra.Enabled = True
    
End Sub

Private Sub DataEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub NFiscalForn_Click()

    iAlterado = REGISTRO_ALTERADO
    
    FrameNFForn.Visible = True
    FrameNFPropria.Visible = False
    
End Sub

Private Sub NFiscalPropria_Click()

    iAlterado = REGISTRO_ALTERADO
    
    FrameNFPropria.Visible = True
    FrameNFForn.Visible = False
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim objFiliais As AdmFiliais
Dim colSerie As New colSerie
Dim colSiglasUF As New Collection
Dim colCodigoDescricao As New AdmColCodigoNome
Dim colFiliais As New Collection

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    gbLimpaTaxa = True

    Set objEventoTransportadora = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoSerie = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoRecebimento = New AdmEvento
    
    'distribuicao
    Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
    Set gobjDistribuicao.objTela = Me
    gobjDistribuicao.bTela = True
    
    'Verifica se módulo de Compras não está ativo
    If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then gError 74845
    
    If gcolModulo.Ativo(MODULO_ESTOQUE) <> MODULO_ATIVO Then gError 74846

    'Lê as séries correspondentes a FilialEmpresa = giFilialEmpresa
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 54495

    'Preenche a List da Combo Serie
    For iIndice = 1 To colSerie.Count
        Serie.AddItem colSerie(iIndice).sSerie
    Next

    'Lê as siglas dos Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colSiglasUF, STRING_ESTADO_SIGLA)
    If lErro <> SUCESSO Then gError 54498

    'Alimenta a Combo PlacaUF.
    For iIndice = 1 To colSiglasUF.Count
        PlacaUF.AddItem colSiglasUF(iIndice)
    Next

    'Lê o código e o Nome Reduzido da Transportadora
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 54499

    'Preenche a Combo Box Transportadora com código e Nome Reduzido
    For iIndice = 1 To colCodigoDescricao.Count
        Transportadora.AddItem colCodigoDescricao(iIndice).iCodigo & "-" & colCodigoDescricao(iIndice).sNome

        'Preenche ItemData com o Código
        Transportadora.ItemData(Transportadora.NewIndex) = colCodigoDescricao(iIndice).iCodigo
    Next

    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeEspecie
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
    If lErro <> SUCESSO Then gError 102438

    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeMarca
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    If lErro <> SUCESSO Then gError 102439

    'Carrega a combo de FiliaisCompra com as Filiais Empresa
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError 54492

    For Each objFiliais In colFiliais
        If objFiliais.iCodFilial <> 0 Then
            FilialCompra.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialCompra.ItemData(FilialCompra.NewIndex) = objFiliais.iCodFilial
        End If
    Next

    'Limpa a combo de pedidos de compra
    PedidosCompra.Clear

    Call CF("Filial_Seleciona", FilialCompra, giFilialEmpresa)

    'Coloca gdtDataAtual em DataEntrada
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    'Inicializa a Mascára de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Prod)
    If lErro <> SUCESSO Then gError 89266
    
    'Inicializa a Mascara de Ccl
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 54500

    'Formata a Quantidade para o Formato de Estoque
    Quantidade.Format = FORMATO_ESTOQUE

    Set objGrid = New AdmGrid
    Set objGridItens = objGrid
    Set objGridItensPC = New AdmGrid

    'Inicializa GridItens
    lErro = Inicializa_GridItens(objGrid)
    If lErro <> SUCESSO Then gError 54497

    'Inicializa GridItensPC
    lErro = Inicializa_GridItensPC(objGridItensPC)
    If lErro <> SUCESSO Then gError 54494

    'Inicializa o grid de Distribuicao
    lErro = gobjDistribuicao.Inicializa_GridDist()
    If lErro <> SUCESSO Then gError 89651
    
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError 108990
    
    'Adiciona a opcao TODOS na combo
    ComboPedidoCompras.Clear
    ComboPedidoCompras.AddItem "TODOS"

    NFiscalPropria.Value = True

    iFilialCompraAnterior = Codigo_Extrai(FilialCompra.Text)

    Set gcolItemPedCompraInfo = New Collection
    Set gColInfoBD = New Collection
    
    ComboPedidoCompras.ListIndex = 0
    
    gbCarregandoTela = False

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 54492, 54494, 54495, 54497, 54498, 54499, 54500, 89651, 108990, 102438, 102439

        Case 74845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_COMPRAS_INATIVO", gErr)
            
        Case 74846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_ESTOQUE_INATIVO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166414)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RecebFornCom"

    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Error 54501

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Fornecedor", objNFiscal.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "FilialForn", objNFiscal.iFilialForn, 0, "FilialForn"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_SERIE, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_NFISCAL_PLACA, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_NFISCAL_PLACA_UF, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, 0, "VolumeEspecie" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, 0, "VolumeMarca" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_NFISCAL_MENSAGEM, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "PesoLiq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "DataEntrada", objNFiscal.dtDataEntrada, 0, "DataEntrada"
'horaentrada
    colCampoValor.Add "HoraEntrada", CDbl(objNFiscal.dtHoraEntrada), 0, "HoraEntrada"
    colCampoValor.Add "FilialPedido", objNFiscal.iFilialPedido, 0, "FilialPedido"
    colCampoValor.Add "NumRecebimento", objNFiscal.lNumRecebimento, 0, "NumRecebimento"
    colCampoValor.Add "Observacao", objNFiscal.sObservacao, STRING_OBSERVACAO_OBSERVACAO, "Observacao"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_IGUAL, STATUS_LANCADO

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 54501

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166415)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colFiliais As New Collection
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objNFiscal.dtDataEntrada = colCampoValor.Item("DataEntrada").vValor
'horaentrada
    objNFiscal.dtHoraEntrada = colCampoValor.Item("HoraEntrada").vValor
    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objNFiscal.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    objNFiscal.iFilialForn = colCampoValor.Item("FilialForn").vValor
    objNFiscal.iTipoNFiscal = colCampoValor.Item("TipoNFiscal").vValor
    objNFiscal.sSerie = colCampoValor.Item("Serie").vValor
    objNFiscal.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
    objNFiscal.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
    objNFiscal.dValorFrete = colCampoValor.Item("ValorFrete").vValor
    objNFiscal.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
    objNFiscal.dValorOutrasDespesas = colCampoValor.Item("ValorOutrasDespesas").vValor
    objNFiscal.dValorDesconto = colCampoValor.Item("ValorDesconto").vValor
    objNFiscal.dValorTotal = colCampoValor.Item("ValorTotal").vValor
    objNFiscal.iCodTransportadora = colCampoValor.Item("CodTransportadora").vValor
    objNFiscal.sPlaca = colCampoValor.Item("Placa").vValor
    objNFiscal.sPlacaUF = colCampoValor.Item("PlacaUF").vValor
    objNFiscal.lVolumeQuant = colCampoValor.Item("VolumeQuant").vValor
    objNFiscal.lVolumeEspecie = colCampoValor.Item("VolumeEspecie").vValor
    objNFiscal.lVolumeMarca = colCampoValor.Item("VolumeMarca").vValor
    objNFiscal.sVolumeNumero = colCampoValor.Item("VolumeNumero").vValor
    objNFiscal.sMensagemNota = colCampoValor.Item("MensagemNota").vValor
    objNFiscal.dPesoLiq = colCampoValor.Item("PesoLiq").vValor
    objNFiscal.dPesoBruto = colCampoValor.Item("PesoBruto").vValor
    objNFiscal.iFilialPedido = colCampoValor.Item("FilialPedido").vValor
    objNFiscal.lNumRecebimento = colCampoValor.Item("NumRecebimento").vValor
    objNFiscal.sObservacao = colCampoValor.Item("Observacao").vValor
    objNFiscal.sVolumeNumero = colCampoValor.Item("VolumeNumero").vValor

    'Lê NFiscal no BD
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 89238

    If lErro = 31442 Then gError 89239

    lErro = Traz_RecebMaterialFCom_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 54502

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 54502, 89238

        Case 89239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_NAO_CADASTRADO", gErr, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166416)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoTransportadora = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoSerie = Nothing
    Set objEventoCcl = Nothing
    Set objEventoRecebimento = Nothing
    
    Set objGrid = Nothing
    Set objGridItensPC = Nothing

    'distribuicao
    Set gobjDistribuicao = Nothing

    Set gcolPedidoCompra = Nothing
    Set gcolItemPedCompraInfo = Nothing
    Set gColInfoBD = Nothing

    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub BotaoLocalizacaoDist_Click()
'distribuicao

    Call gobjDistribuicao.BotaoLocalizacaoDist_Click

End Sub

Public Sub ItemNFDist_Change()
'distribuicao

    Call gobjDistribuicao.ItemNFDist_Change

End Sub

Public Sub ItemNFDist_GotFocus()
'distribuicao

    Call gobjDistribuicao.ItemNFDist_GotFocus

End Sub

Public Sub ItemNFDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call gobjDistribuicao.ItemNFDist_KeyPress(KeyAscii)

End Sub

Public Sub ItemNFDist_Validate(Cancel As Boolean)
'distribuicao

    Call gobjDistribuicao.ItemNFDist_Validate(Cancel)

End Sub

Public Sub AlmoxDist_Change()
'distribuicao

    Call gobjDistribuicao.AlmoxDist_Change

End Sub

Public Sub AlmoxDist_GotFocus()
'distribuicao

    Call gobjDistribuicao.AlmoxDist_GotFocus

End Sub

Public Sub AlmoxDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call gobjDistribuicao.AlmoxDist_KeyPress(KeyAscii)

End Sub

Public Sub AlmoxDist_Validate(Cancel As Boolean)
'distribuicao

    Call gobjDistribuicao.AlmoxDist_Validate(Cancel)

End Sub

Public Sub QuantDist_Change()
'distribuicao

    Call gobjDistribuicao.QuantDist_Change

End Sub

Public Sub QuantDist_GotFocus()
'distribuicao

    Call gobjDistribuicao.QuantDist_GotFocus

End Sub

Public Sub QuantDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call gobjDistribuicao.QuantDist_KeyPress(KeyAscii)

End Sub

Public Sub QuantDist_Validate(Cancel As Boolean)
'distribuicao

    Call gobjDistribuicao.QuantDist_Validate(Cancel)

End Sub

Private Sub Fornecedor_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 54503

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then Error 54504

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

        'Se Fornecedor não foi preenchido limpa a combo de Filiais
        Filial.Clear
        ComboPedidoCompras.Clear
        ComboPedidoCompras.AddItem "TODOS"
        Moeda.ListIndex = -1
        Taxa.Text = ""
        Taxa.Enabled = False
        LabelTaxa.Enabled = False

    End If

    If sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
    Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text) Then

        Call Atualiza_ListaPedidos

    End If

    iFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case Err

        Case 54503, 54504

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166417)

    End Select

    Exit Sub

End Sub

Function Atualiza_ListaPedidos(Optional objNFiscal As ClassNFiscal) As Long
'Atualiza a Lista de Pedidos de Compra com os códigos desses pedidos

Dim lErro As Long
Dim iFilial As Integer
Dim iFilialCompra As Integer
Dim objPedidoCompras As ClassPedidoCompras
Dim objPedidoCompras1 As ClassPedidoCompras
Dim objFornecedor As New ClassFornecedor
Dim colPedidos As New Collection
Dim colPedidosRec As New Collection
Dim iEstaNaColecao As Integer

On Error GoTo Erro_Atualiza_ListaPedidos
    
    'Limpa os Grids
    Call Grid_Limpa(objGrid)
    Call Grid_Limpa(objGridItensPC)

    'Limpa a Lista de Pedidos de Compra
    PedidosCompra.Clear

    'Limpa as coleções
    Set gColInfoBD = New Collection
    Set gcolPedidoCompra = New Collection
    Set gcolItemPedCompraInfo = New Collection

    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.ClipText)) > 0 And Len(Trim(Filial.Text)) > 0 And Len(Trim(FilialCompra.Text)) > 0 Then

        sFornecedorAnterior = Fornecedor.Text
        iFilialAnterior = Codigo_Extrai(Filial.Text)
        iFilialCompraAnterior = Codigo_Extrai(FilialCompra.Text)

        objFornecedor.sNomeReduzido = Fornecedor.Text
        iFilial = Codigo_Extrai(Filial.Text)
        iFilialCompra = Codigo_Extrai(FilialCompra.Text)

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 54496
        If lErro = 6681 Then gError 54691

        'Le os pedidos de compra enviados para o fornecedor em questão com quantidade a receber
        lErro = CF("PedidoCompra_Le_EnvComQuantReceber", colPedidos, objFornecedor.lCodigo, iFilial, iFilialCompra)
        If lErro <> SUCESSO Then gError 66131

        If Not objNFiscal Is Nothing Then

            'Lê os Pedidos de Compra relacionados ao Recebimento
            lErro = CF("PedidoCompra_Le_Recebimento", objNFiscal, colPedidosRec)
            If lErro <> SUCESSO And lErro <> 66136 Then gError 67000
    
            lErro = CF("Aglutina_Pedidos", gcolPedidoCompra, colPedidos, colPedidosRec)
            If lErro <> SUCESSO Then gError 89272

        Else
        
            Set gcolPedidoCompra = colPedidos

        End If

        'Preenche a ListBox de pedidos com os pedidos lidos do BD
        For Each objPedidoCompras In gcolPedidoCompra
            PedidosCompra.AddItem objPedidoCompras.lCodigo
        Next

    End If

    Atualiza_ListaPedidos = SUCESSO

    Exit Function

Erro_Atualiza_ListaPedidos:

    Atualiza_ListaPedidos = gErr

    Select Case gErr

        Case 54496, 54506, 61744

        Case 54507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDIDOCOMPRAS", gErr, objFornecedor.sNomeReduzido, iFilial, iFilialCompra)

        Case 54691
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166418)

    End Select

    Exit Function

End Function

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche nomeReduzido com o fornecedor da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub Mensagem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscalLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NFiscalLabel_Click

'    'Verifica preenchimento de Fornecedor
'    If Len(Trim(Fornecedor.Text)) <> 0 Then
'
'        objFornecedor.sNomeReduzido = Fornecedor.Text
'        'Lê o Fornecedor
'        lErro = CF("Fornecedor_Le_NomeReduzido",objFornecedor)
'        If lErro <> SUCESSO And lErro <> 6681 Then Error 54511
'
'        'Se não achou o Fornecedor --> erro
'        If lErro = 6681 Then Error 54512
'
'        objNFiscal.lFornecedor = objFornecedor.lCodigo
'
'    End If
'
'    objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
'    objNFiscal.iFilialEmpresa = giFilialEmpresa
'
'    If NFiscalPropria.Value Then
'        objNFiscal.iTipoNFiscal = DOCINFO_NRFPCO
'    ElseIf NFiscalForn.Value Then
'        objNFiscal.iTipoNFiscal = DOCINFO_NRFFCO
'    Else
'        objNFiscal.iTipoNFiscal = 0
'    End If
'
'    objNFiscal.sSerie = Serie.Text
'
'    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
'    objNFiscal.dtDataEntrada = MaskedParaDate(DataEntrada)
'
'    Call Chama_Tela("RecebMaterialFComLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

Erro_NFiscalLabel_Click:

    Select Case Err

        Case 54511

        Case 54512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, Fornecedor.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166419)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Nome Reduzido na Tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Dispara o Validate de Fornecedor
    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    lErro = Traz_RecebMaterialFCom_Tela(objNFiscal)
    If lErro <> SUCESSO Then Error 54516

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case Err

        Case 54516

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166420)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    Serie.Text = objSerie.sSerie

    Me.Show

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Preenche o Text com Código e NomeReduzido
    Transportadora.Text = objTransportadora.iCodigo & "-" & objTransportadora.sNomeReduzido

    Me.Show

End Sub

Private Sub PesoBruto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoBruto_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoBruto.Text)) = 0 Then Exit Sub

    lErro = Valor_NaoNegativo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then Error 54517

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case Err

        Case 54517 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166421)

    End Select

    Exit Sub

End Sub

Private Sub PesoLiquido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoLiquido_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoLiquido.Text)) = 0 Then Exit Sub

    lErro = Valor_NaoNegativo_Critica(PesoLiquido.Text)
    If lErro <> SUCESSO Then Error 54518

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case Err

        Case 54518 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166422)

    End Select

    Exit Sub

End Sub

Private Sub Placa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PlacaUF_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PlacaUF_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'Verifica se foi preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO Then Error 54519

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case Err

        Case 54519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", Err, PlacaUF.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166423)

    End Select

    Exit Sub

End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()

Dim lErro As Long

On Error GoTo Erro_Serie_Click

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Serie_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166424)

    End Select

    Exit Sub

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    'Verifica se a Série está preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub

    'Verifica se é uma Série selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub

    'Verifica se NFiscalPropria

    If NFiscalPropria.Value Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(Serie)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 54521

        If lErro = 12253 Then Error 54522

    Else
        
        'Verifica se tamanho da série é maior do que o espaço no bd ==> erro
        If Len(Trim(Serie.Text)) > STRING_SERIE Then Error 54523

    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case Err

        Case 54521

        Case 54522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case 54523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166425)

    End Select

    Exit Sub

End Sub

Private Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As New Collection

    objSerie.sSerie = Serie.Text

    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

End Sub

Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not objNFiscal Is Nothing Then

        'Lê NFiscal no BD
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then Error 54524

        If lErro <> 31442 Then 'Se ela existe

            If objNFiscal.iTipoNFiscal <> DOCINFO_NRFFCO And objNFiscal.iTipoNFiscal <> DOCINFO_NRFPCO Then Error 54526

            lErro = Traz_RecebMaterialFCom_Tela(objNFiscal)
            If lErro <> SUCESSO Then Error 54525

        Else 'Se não existe
            Error 54527

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 54524, 54525

        Case 54526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_NAO_RECEBFORN", Err, objNFiscal.iTipoNFiscal)

        Case 54527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_NAO_CADASTRADO", Err, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166426)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    'Verifica o preenchimento da Data de Entrada
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Critica a Data
    lErro = Data_Critica(DataEntrada.Text)
    If lErro <> SUCESSO Then Error 54528

    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True

    Select Case Err

        Case 54528 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166427)

    End Select

    Exit Sub

End Sub

'horaentrada
Public Sub HoraEntrada_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HoraEntrada, iAlterado)

End Sub

'horaentrada
Public Sub HoraEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'horaentrada
Public Sub HoraEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraEntrada_Validate

    'Verifica se a hora de Entrada foi digitada
    If Len(Trim(HoraEntrada.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraEntrada.Text)
    If lErro <> SUCESSO Then gError 89816

    Exit Sub

Erro_HoraEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 89816

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166428)

    End Select

    Exit Sub

End Sub

Private Sub IPIValor1_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Opcao_Click()

   'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If

End Sub

Private Sub Transportadora_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection

    'Preenche o código da Transportadora
    If Len(Trim(Transportadora.Text)) <> 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)

    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

End Sub

Private Sub UpDownEntrada_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_DownClick

    'Verifica preenchimento da Data de Entrada
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 54529

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case Err

        Case 54529 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166429)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntrada_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_UpClick

    'Verifica preenchimneto da Data
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Aumanta a Data
    lErro = lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 54530

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case Err

        Case 54530 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166430)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim sUnidadeMed As String
Dim iIndice As Integer
Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim sProdutoMascarado As String
Dim iCont As Integer
Dim objPedidoCompra As New ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Se foi chamada do Saida de Celula, sai da rotina
    If iLocalChamada = 3 Then Exit Sub
    
    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name

        'Produto
        Case Produto.Name
            
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
                
                'Limpa a combo de Produtos
                Produto.Clear
                
                'Para cada Pedido de Compras da lista
                For iIndice = 0 To PedidosCompra.ListCount - 1

                    'Se o Pedido estiver selecionado
                    If PedidosCompra.Selected(iIndice) = True Then
                        
                        Set objPedidoCompra = gcolPedidoCompra.Item(iIndice + 1)
                    
                        For Each objItemPC In objPedidoCompra.colItens
    
                            'Mascara o Produto
                            lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
                            If lErro <> SUCESSO Then gError 54531
                            
                            Prod.PromptInclude = False
                            Prod.Text = sProdutoMascarado
                            Prod.PromptInclude = True
                    
                            sProdutoMascarado = Prod.Text

                            For iCont = 0 To Produto.ListCount - 1
                                If sProdutoMascarado = Produto.List(iCont) Then
                                    Exit For
                                End If
                            Next
                            
                            If iCont = Produto.ListCount Then
                                Produto.AddItem sProdutoMascarado
                            End If
                            
                        Next

                    End If

                Next

                'Remove da Combo os Produtos que já estão em GridItens
                For iIndice = Produto.ListCount - 1 To 0 Step -1
                    For iCont = 1 To objGrid.iLinhasExistentes
                        If GridItens.TextMatrix(iCont, iGrid_Produto_Col) = Produto.List(iIndice) Then
                            Produto.RemoveItem (iIndice)
                        End If
                    Next
                Next

                For iIndice = 0 To PedidosCompra.ListCount - 1
                    If PedidosCompra.Selected(iIndice) = True Then
                        Exit For
                    End If
                Next
            
                If PedidosCompra.ListCount = iIndice Then Produto.Clear
            
            End If
            
        'Unidade de Medida
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 54533

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado

                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 54534

                'Não achou o Produto
                If lErro = 28030 Then gError 61771

                objClasseUM.iClasse = objProduto.iClasseUM

                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 54535

                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next

                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM

                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 54670
                End If

            End If

        'Nas demais
'distribuicao
'        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, Quantidade.Name, Almoxarifado.Name, DescricaoItem.Name, Ccl.Name
        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, Quantidade.Name, DescricaoItem.Name, Ccl.Name

            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 54705

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
            'Se foi configurado para não aceitar valor unitário do Produto do ItemNF diferente do ItemPC
            If objControl.Name = ValorUnitario.Name And gobjCOM.iNFDiferentePC = NFISCAL_NAO_ACEITA_DIFERENCA_PC Then
                objControl.Enabled = False
            End If
        
    End Select

    'distribuicao
    lErro = gobjDistribuicao.Rotina_Grid_Enable_Dist(iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 89652

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 54531, 54533, 54534, 54535, 54670, 54705, 89652

        Case 61771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166431)

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

            'Se for o GridItens de Recebimento
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 54537

            'Se for o GridItens de Pedido de Compras
            Case GridItensPC.Name

                lErro = Saida_Celula_GridItensPC(objGridInt)
                If lErro <> SUCESSO Then gError 54538

            'distribuicao
            Case GridDist.Name

            lErro = gobjDistribuicao.Saida_Celula_Dist()
            If lErro <> SUCESSO Then gError 89653

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 54539

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 54537, 54538, 54539, 89653
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166432)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    Select Case objGridInt.objGrid.Col

        'distribuicao
'        'Almoxarifado
'        Case iGrid_Almoxarifado_Col
'            lErro = Saida_Celula_Almoxarifado(objGridInt)
'            If lErro <> SUCESSO Then Error 54542

        'Valor Unitário
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_ValorUnitario(objGridInt)
            If lErro <> SUCESSO Then Error 54543

        'Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then Error 54544

        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then Error 54545

        'Unidade de Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UnidadeMed(objGridInt)
            If lErro <> SUCESSO Then Error 54546

        'Percentagem de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then Error 54540

        'Desconto
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then Error 54541

        'Descrição
        Case iGrid_Descricao_Col
            lErro = Saida_Celula_DescricaoItem(objGridInt)
            If lErro <> SUCESSO Then Error 54548

        'Ccl
        Case iGrid_Ccl_Col
            lErro = Saida_Celula_Ccl(objGridInt)
            If lErro <> SUCESSO Then Error 54547

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = Err

    Select Case Err

        Case 54540, 54541, 54542, 54543, 54544, 54545, 54546, 54547, 54548
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166433)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItensPC(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensPC

    Select Case objGridInt.objGrid.Col

            'Quantidade Recebida
            Case iGrid_Recebido_Col
                lErro = Saida_Celula_QuantRecebida(objGridInt)
                If lErro <> SUCESSO Then Error 54549

    End Select

    Saida_Celula_GridItensPC = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensPC:

    Saida_Celula_GridItensPC = Err

    Select Case Err

        Case 54549
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166434)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantRecebida(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim objItemPCInfo As New ClassItemPedCompraInfo
Dim objItemPedCompraInfo As New ClassItemPedCompraInfo
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sProduto As String
Dim dQuantPosterior As Double
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantRecebidaAnterior As Double

On Error GoTo Erro_Saida_Celula_QuantRecebida

    Set objGridInt.objControle = QuantRecebida

    'Antes: dQuantRecebidaAnterior = StrParaDbl(QuantRecebida.Text)
    dQuantRecebidaAnterior = StrParaDbl(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_Recebido_Col))

    'Se quantidade recebida estiver preenchida
    If Len(Trim(QuantRecebida.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantRecebida.Text)
        If lErro <> SUCESSO Then gError 54550

        dQuantidade = StrParaDbl(QuantRecebida.Text)

        'Coloca o valor Formatado na tela
        QuantRecebida.Text = Formata_Estoque(dQuantidade)

        sProduto = GridItensPC.TextMatrix(GridItensPC.Row, iGrid_Prod_Col)
        
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
        If lErro <> SUCESSO Then gError 89290
        
        objProduto.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89291
        
        If lErro = 28030 Then gError 89292
        
        'Descobre o Indice do produto alterado no grid de itens
        For iIndice1 = 1 To objGrid.iLinhasExistentes
            If sProduto = GridItens.TextMatrix(iIndice1, iGrid_Produto_Col) Then Exit For
        Next
        
        'Para cada Item da coleção de ItensPC
        For Each objItemPedCompraInfo In gcolItemPedCompraInfo
            
            If sProdutoFormatado = objItemPedCompraInfo.sProduto Then
            
                If objItemPedCompraInfo.lPedCompra = StrParaLong(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_PedCompra_Col)) Then
                    
                    objItemPedCompraInfo.dQuantRecebida = dQuantidade
                    
                    'Se a quantidade recebida for maior que a quantidade a receber vezes a porcentagem a mais de recebimento
                    If StrParaDbl(QuantRecebida.Text) > StrParaDbl(objItemPedCompraInfo.dQuantReceber) + (objItemPedCompraInfo.dQuantReceber * objItemPedCompraInfo.dPercentMaisReceb) Then gError 54551
                    
                    'Converte a UM de GridItensPC para a UM do GridItens
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItensPC.TextMatrix(GridItensPC.Row, iGrid_UM_Col), GridItens.TextMatrix(iIndice1, iGrid_UnidadeMed_Col), dFator)
                    If lErro <> SUCESSO Then gError 89294
                    
                    dQuantPosterior = dQuantPosterior + dQuantidade * dFator
        
                Else
            
                    'Converte a UM de GridItensPC para a UM do GridItens
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPedCompraInfo.sUM, GridItens.TextMatrix(iIndice1, iGrid_UnidadeMed_Col), dFator)
                    If lErro <> SUCESSO Then gError 89293
                
                    dQuantPosterior = dQuantPosterior + objItemPedCompraInfo.dQuantRecebida * dFator
                    
                End If
                
            End If
            
        Next
        
        If dQuantPosterior = 0 Then
            GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col) = ""
        Else
            GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col) = Formata_Estoque(dQuantPosterior)
        End If
        
        Call Calcula_Valores(iIndice1)
                
    Else
        QuantRecebida.Text = Formata_Estoque(0)
    End If

    If dQuantidade <> dQuantRecebidaAnterior Then

        lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 89662

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30456

    Saida_Celula_QuantRecebida = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantRecebida:

    Saida_Celula_QuantRecebida = gErr

    Select Case gErr

        Case 54550, 89290, 89291, 89293, 89294, 89662
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTRECEBIDA_MAIOR_QUANTRECEBER", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166435)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se a Data esta preenchida
    If Len(Trim(Produto.Text)) > 0 Then

        lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 54552
        
        If lErro = 25041 Then gError 54690

        'Critica se o produto é comprável
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            'não pode ser um produto não comprável
            If objProduto.iCompras = PRODUTO_NAO_COMPRAVEL Then gError 54637
                        
        End If
        
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 54638
        
        Prod.PromptInclude = False
        Prod.Text = sProduto
        Prod.PromptInclude = True

        sProduto = Prod.Text
        
        For iIndice = 0 To Produto.ListCount - 1
            If Produto.List(iIndice) = sProduto Then
                Produto.ListIndex = iIndice
            End If
        Next
        
        'Coloca as demais características do produto na tela
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 54553
        
            Call Inclui_ItensPC(objProduto, GridItens.Row)
            
            'Atualiza as Quantidades de GridItensPC
            For iIndice = 1 To objGridItensPC.iLinhasExistentes
                If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = Produto.Text Then
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col)))
                End If
            Next
                    
        End If

    End If

    sProduto = Produto.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 54554

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = sProduto

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr
         
        Case 54552, 54553, 54554
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54638
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54690
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166436)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 54556

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = Err

    Select Case Err

        Case 54556
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166437)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iCont As Integer

On Error GoTo Erro_ProdutoLinha_Preenche

    'Preenche no Grid a Descrição do Produto e a Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
    GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao

'distribuicao
'    lErro = CF("AlmoxarifadoPadrao_Le",giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
'    If lErro <> SUCESSO And lErro <> 23796 Then gError 54557
'
'    If lErro = SUCESSO And iAlmoxarifado <> 0 Then
'        objAlmoxarifado.iCodigo = iAlmoxarifado
'
'        lErro = CF("Almoxarifado_Le",objAlmoxarifado)
'        If lErro <> SUCESSO And lErro <> 25056 Then gError 54558
'
'        'Se não achou o Almoxarifado --> erro
'        If lErro = 25056 Then gError 54559
'
'        'Coloca o Nome Reduzido na Coluna Almoxarifado
'        GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
'
'    End If

    'Preço Unitário
    For iCont = 1 To gcolItemPedCompraInfo.Count
        If objProduto.sCodigo = gcolItemPedCompraInfo(iCont).sProduto Then
            GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(gcolItemPedCompraInfo(iCont).dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
            Exit For
        End If
    Next
    
    'Se necessário cria uma nova linha no Grid
    If GridItens.Row - GridItens.FixedRows = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 54557, 54558

        Case 54559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166438)

    End Select

    Exit Function

End Function

Function InsereOrdenadaPC(objItemPCInfo As ClassItemPedCompraInfo) As Long
'Insere novo item de Pedido de compras na coleção global de ItensPC de forma ordenada

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_InsereOrdenadaPC

    iIndice = 1
    
    'Se a coleção está vazia
    If gcolItemPedCompraInfo.Count = 0 Then
        'Insere o ItemPC na coleção
        gcolItemPedCompraInfo.Add objItemPCInfo
    
    'Se não
    Else
    
        'Para cada Item da coleção de ItensPC
        Do While iIndice <= gcolItemPedCompraInfo.Count
            
            'Se o código do Pedido de Compra do Item passado for maior ou igual do que o que está na coleção
            If objItemPCInfo.lPedCompra >= gcolItemPedCompraInfo(iIndice).lPedCompra Then
                'Busca próximo Item da coleção
                iIndice = iIndice + 1
            'Se não
            Else
                'Adiciona o Item na coleção de ItensPC antes do Item da coleção que possui código PC maior
                gcolItemPedCompraInfo.Add objItemPCInfo, , iIndice
                Exit Do
            End If
            
        Loop
                    
        'Se o Pedido de Compra do Item é maior do que todos que estão na coleção
        If iIndice > gcolItemPedCompraInfo.Count Then
            'Adiciona o Item no final coleção
            gcolItemPedCompraInfo.Add objItemPCInfo
        End If
    
    End If
    
    'Limpa o GridItensPC
    Call Grid_Limpa(objGridItensPC)
    
    'Preenche o GridItensPC
    For iIndice = 1 To gcolItemPedCompraInfo.Count
                    
        lErro = PreencheLinha_ItensPC(gcolItemPedCompraInfo(iIndice))
        If lErro <> SUCESSO Then gError 61696
    
    Next
    
    InsereOrdenadaPC = SUCESSO
    
    Exit Function
    
Erro_InsereOrdenadaPC:

    InsereOrdenadaPC = gErr
    
    Select Case gErr
    
        Case 61696
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166439)
    
    End Select
    
    Exit Function
    
End Function

Function Inclui_ItensPC(objProduto As ClassProduto, iLinha As Integer) As Long
'Preenche Grid de Pedido de Compras

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objPedidoCompras As New ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim lErro As Long
Dim dFator As Double

On Error GoTo Erro_Inclui_ItensPC

    'Para cada pedido de compras
    For iIndice = 0 To PedidosCompra.ListCount - 1

        'Se o pedido estiver selecionado
        If PedidosCompra.Selected(iIndice) = True Then
            
            Set objPedidoCompras = gcolPedidoCompra.Item(iIndice + 1)

            iIndice2 = 0

            'Para cada item do pedido de compras
            For Each objItemPC In objPedidoCompras.colItens

                iIndice2 = iIndice2 + 1

                'Se o produto do item for igual ao do objProduto
                If objItemPC.sProduto = objProduto.sCodigo Then

                    Set objItemPCInfo = New ClassItemPedCompraInfo

                    objItemPCInfo.dPercentMaisReceb = objItemPC.dPercentMaisReceb
                    objItemPCInfo.dQuantReceber = objItemPC.dQuantidade
                    objItemPCInfo.dQuantRecebida = objItemPC.dQuantRecebida + objItemPC.dQuantRecebimento
                    objItemPCInfo.iItem = iIndice2
                    objItemPCInfo.iRecebForaFaixa = objItemPC.iRebebForaFaixa
                    objItemPCInfo.lNumIntDoc = objItemPC.lNumIntDoc
                    objItemPCInfo.lPedCompra = objPedidoCompras.lCodigo
                    objItemPCInfo.sDescProduto = objItemPC.sDescProduto
                    objItemPCInfo.sProduto = objItemPC.sProduto
                    objItemPCInfo.sUM = objItemPC.sUM

                    'Lê o Produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then Error 54707
                    If lErro = 28030 Then Error 54708

                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dFator)
                    If lErro <> SUCESSO Then Error 54709

                    objItemPCInfo.dQuantReceber = objItemPCInfo.dQuantReceber * dFator
                    objItemPCInfo.dQuantRecebida = objItemPCInfo.dQuantRecebida * dFator
                    objItemPCInfo.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
                    
                    'Adiciona o item em gColItemPedCompraInfo
                    Call InsereOrdenadaPC(objItemPCInfo)

                End If

            Next

        End If

    Next

    Inclui_ItensPC = SUCESSO

    Exit Function

Erro_Inclui_ItensPC:

    Inclui_ItensPC = Err

    Select Case Err

        Case 54561, 54709

        Case 54708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166440)

    End Select

    Exit Function

End Function

Sub Atualiza_QuantidadePC(dQuantAnterior As Double, dQuantPosterior As Double, iLinhaGridNF As Integer)
'Atualiza a quantidade Recebida do GridItensPC a partir da Quantidade que foi modificada do GridItens NF

Dim sProduto As String
Dim dQuantDiferenca As Double
Dim iIndice As Integer
Dim objItemPCInfo As ClassItemPedCompraInfo

    'Guarda o código do Produto do GridItens
    sProduto = GridItens.TextMatrix(iLinhaGridNF, iGrid_Produto_Col)
    
    'Guarda QuantDiferenca = Quantidade recebida total que estava no Grid - quantidade recebida total atual
    dQuantDiferenca = dQuantPosterior - dQuantAnterior
    
    'Se a quantidade aumentou
    If dQuantDiferenca > 0 Then

        For iIndice = 1 To objGridItensPC.iLinhasExistentes
            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto Then

                'Atualiza a Quantidade Recebida no GridItensPC
                If StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)) >= StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca Then
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca)
                    
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    
                    Exit For

                Else
                    dQuantDiferenca = dQuantDiferenca - (StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)) - StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)))
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)
                    
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                            
                    
                End If
            End If
        Next

    'Se a quantidade diminuiu
    ElseIf dQuantDiferenca < 0 Then

        For iIndice = objGridItensPC.iLinhasExistentes To 1 Step -1
            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto Then

                If StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca >= 0 Then
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(StrParaDbl(StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca))
                    
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    
                    Exit For

                Else
                    dQuantDiferenca = dQuantDiferenca + StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(0)
                    
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = 0
                    
                End If

            End If
        Next

    End If

End Sub

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objItemPCInfo As New ClassItemPedCompraInfo
Dim dQuantAnterior As Double
Dim dQuantPosterior As Double
Dim dQuantReceber As Double
Dim iFilialEmpresa As Integer
Dim dQuantTotalPC As Double
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantTotal As Double
'distribuicao
Dim dQuantidadeAnterior As Double
Dim dQuantidadeAtual As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'distribuicao
    dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    'fim  distribuicao

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 65748
    
        dQuantPosterior = StrParaDbl(Quantidade.Text)
        dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    
        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantPosterior)
            
        'distribuicao
        dQuantidadeAtual = StrParaDbl(Quantidade.Text)
        'fim  distribuicao
            
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 89305

        objProduto.sCodigo = sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89306
        
        If lErro = 28030 Then gError 89307

        dQuantTotal = 0
        
        'Acumula a quantidade
        For Each objItemPCInfo In gcolItemPedCompraInfo
        
            If objItemPCInfo.sProduto = objProduto.sCodigo Then
            
                'Converte a UM de GridItensPC para a UM do GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPCInfo.sUM, GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 89308
            
                dQuantTotal = dQuantTotal + (objItemPCInfo.dQuantReceber + (objItemPCInfo.dQuantPedida * objItemPCInfo.dPercentMaisReceb)) * dFator
                
            End If
        Next
            
        'SE A QUANTIDADE ULTRAPASSAR O TOTAL COM O % A MAIS PERMITIDO --> ERRO
        If (StrParaDbl(Quantidade.Text) - dQuantTotal) > QTDE_ESTOQUE_DELTA Then gError 65749
    
        'Atualiza a quantidade Recebida do GriditensPC
        Call Atualiza_QuantidadePC(dQuantAnterior, dQuantPosterior, GridItens.Row)

    End If
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65750

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 65751
       
    'Recalcula o total da nota fiscal
    Call Total_Calcula
    
    'inicio distribuicao
    If dQuantidadeAnterior <> dQuantidadeAtual Then
        
        lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 89629
        
    End If
    'fim distribuicao
        
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 65748, 65750, 65751, 65752, 83279, 89305, 89306, 89308, 89548, 89629
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_MAIOR_TOTALRECEBER", gErr, dQuantPosterior, dQuantTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 89307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166441)

    End Select

    Exit Function
    
End Function

Function Critica_Valores(objNFiscal As ClassNFiscal) As Long
'Verifica se os valores do Pedidos de Compras e dos ItensPC são
'iguais aos que foram colocados no frame de Itens

Dim lErro As Long
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim dValorTotalFrete As Double
Dim dValorTotalSeguro As Double
Dim dValorTotalDespesas As Double
Dim dValorTotalDescontos As Double
Dim iIndice As Integer
Dim iIndice3 As Integer
Dim objPedidoCompras As ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim sProdutoMascarado As String
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice2 As Integer
Dim iItem As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double, dPrecoUnitNF As Double, dPrecoUnitPC As Double
Dim bAchou As Boolean, bIgnora As Boolean

On Error GoTo Erro_Critica_Valores

    'Lê os valores dos Pedidos de Compras marcados
    lErro = CF("PedidoCompras_Le_Valores", gcolPedidoCompra)
    If lErro <> SUCESSO Then gError 66617
        
    'Para cada Pedido de Compras
    For Each objPedidoCompras In gcolPedidoCompra
        
        For iIndice2 = 0 To PedidosCompra.ListCount - 1
            
            If PedidosCompra.List(iIndice2) = CStr(objPedidoCompras.lCodigo) And PedidosCompra.Selected(iIndice2) = True Then
        
                'Acumula os Valores dos Pedidos marcados
                dValorTotalFrete = dValorTotalFrete + objPedidoCompras.dValorFrete
                dValorTotalSeguro = dValorTotalSeguro + objPedidoCompras.dValorSeguro
                dValorTotalDespesas = dValorTotalDespesas + objPedidoCompras.dOutrasDespesas
                dValorTotalDescontos = dValorTotalDescontos + objPedidoCompras.dValorDesconto
                
                Set objPedidoCompras.colItens = New Collection
                
                'Lê os Itens do Pedido
                lErro = CF("ItensPC_Le_Codigo", objPedidoCompras)
                If lErro <> SUCESSO Then gError 66622
            
                'Para cada Item do Pedido de Compras
                For Each objItemPC In objPedidoCompras.colItens
                    
                    objProduto.sCodigo = objItemPC.sProduto
            
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 89281
                    
                    If lErro = 28030 Then gError 89282
                    
                    'Mascara o Produto
                    lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
                    If lErro <> SUCESSO Then gError 66623
                    
                    Prod.PromptInclude = False
                    Prod.Text = sProdutoMascarado
                    Prod.PromptInclude = True
            
                    sProdutoMascarado = Prod.Text
                    
                    bAchou = False
                    bIgnora = True
                    For Each objItemPCInfo In gcolItemPedCompraInfo
                        If objItemPCInfo.lNumIntDoc = objItemPC.lNumIntDoc Then
                            bAchou = True
                            If objItemPCInfo.dQuantReceber > QTDE_ESTOQUE_DELTA Then bIgnora = False 'Tem algo sendo recebido desse item do pedido
                        End If
                    Next
                    
                    If Not bIgnora Then
                        
                        'Procura no GridItens o Produto igual ao do ItemPC
                        For iIndice = 1 To objGrid.iLinhasExistentes
                            
                            'Se encontrou
                            If sProdutoMascarado = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) Then
                                
                                'Converte a UM de GridItensPC para a UM do GridItens
                                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                                If lErro <> SUCESSO Then gError 89283
                                
                                'preco unitario do item da nf na unidade de medida do item pc
                                dPrecoUnitNF = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) * dFator
                                'preco unitario do item de pc em R$
                                dPrecoUnitPC = objItemPC.dPrecoUnitario * IIf(objItemPC.iMoeda <> MOEDA_REAL, objItemPC.dTaxa, 1)
                                
                                'Verifica se possui o mesmo unitário
                                If Abs(dPrecoUnitNF - dPrecoUnitPC) > DELTA_VALORMONETARIO Then
                                
                                    'Se aceita preços unitários de Itens da Nota Fiscal diferente dos Itens PC
                                    If gobjCOM.iNFDiferentePC = NFISCAL_ACEITA_DIFERENCA_PC Then
        
                                        'Exibe mensagem de aviso de o preço unitário for diferente
                                        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORUNITARIO_DIFERENTE_PC", dPrecoUnitNF, iIndice, dPrecoUnitPC)
                                        If vbMsgRes = vbNo Then gError 66624
        
                                    'Se não aceita
                                    Else
        
                                        'Exibe mensagem de erro
                                        gError 67452
        
                                    End If
                                                
                                End If
                                
                                'Se aceita preços unitários de Itens da Nota Fiscal diferente dos Itens PC
                                If gobjCOM.iNFDiferentePC = NFISCAL_ACEITA_DIFERENCA_PC Then
                                
                                    'Verifica se possui o mesmo desconto
                                    If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) <> objItemPC.dValorDesconto Then
                                        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESCONTOITEM_DIFERENTE_PC", StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)), iIndice, objItemPC.dValorDesconto)
                                        If vbMsgRes = vbNo Then gError 66625
                                    End If
                                
                                'Se não aceita
                                ElseIf StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) * dFator <> objItemPC.dPrecoUnitario * IIf(objItemPC.dTaxa = 0, 1, objItemPC.dTaxa) Then
    
                                    'Exibe mensagem de erro
                                    gError 67452
    
                                End If
                                
                            End If
                        
                        Next
                        
                    End If
                    
                Next
        
            End If
        
        Next
    
    Next
    
    'Se o Valor Frete da Nota Fiscal for diferente do Pedido de Compras
    If objNFiscal.dValorFrete <> dValorTotalFrete Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORFRETE_DIFERENTE_PC", objNFiscal.dValorFrete, dValorTotalFrete)
        If vbMsgRes = vbNo Then gError 66618
    End If
    
    'Se o Valor Seguro da Nota Fiscal for diferente do Pedido de Compras
    If objNFiscal.dValorSeguro <> dValorTotalSeguro Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORSEGURO_DIFERENTE_PC", objNFiscal.dValorSeguro, dValorTotalSeguro)
        If vbMsgRes = vbNo Then gError 66619
    End If
    
    'Se o Valor Despesas da Nota Fiscal for diferente do Pedido de Compras
    If objNFiscal.dValorOutrasDespesas <> dValorTotalDespesas Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORDESPESAS_DIFERENTE_PC", objNFiscal.dValorOutrasDespesas, dValorTotalDespesas)
        If vbMsgRes = vbNo Then gError 66620
    End If
    
    'Se o Valor Desconto da Nota Fiscal for diferente do Pedido de Compras
    If objNFiscal.dValorDesconto <> dValorTotalDescontos Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORDESCONTO_DIFERENTE_PC", objNFiscal.dValorDesconto, dValorTotalDescontos)
        If vbMsgRes = vbNo Then gError 66621
    End If
                
    Critica_Valores = SUCESSO
    
    Exit Function
    
Erro_Critica_Valores:

    Critica_Valores = gErr
    
    Select Case gErr
    
        Case 66617, 66618, 66619, 66620, 66621, 66622, 66623, 66624, 66625, 89281, 89283
        
        Case 67452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_PRECOUNITARIO_DIFERENTE", gErr, objItemPC.sProduto)
                    
        Case 89282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166442)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_Valores(iLinha As Integer) As Long

Dim sProduto As String
Dim lErro As Long
Dim lTamanho As Long
Dim dPercentDesc As Double
Dim dValorUnitario As Double
Dim dDesconto As Double
Dim dValorReal As Double
Dim dQuantidade As Double

On Error GoTo Erro_Calcula_Valores

    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dQuantidade = 0 Or Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))) = 0 Then

        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_ValorTotal_Col) = ""
        
    Else

        dValorUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
        dDesconto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
        
        lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))

        If lTamanho > 0 Then
            dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col))
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If


        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, dPercentDesc, dDesconto, dValorReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If

        
        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(dValorReal, "Standard")

    End If

    lErro = SubTotal_Calcula()
    If lErro <> SUCESSO Then gError 65754

    Calcula_Valores = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valores:

    Calcula_Valores = gErr
    
    Select Case gErr

        Case 65754

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166443)

    End Select

    Exit Function
End Function

Private Function SubTotal_Calcula() As Long
'Soma a coluna de Valor Total e acumula em SubTotal

Dim lErro As Long
Dim dSubTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_SubTotal_Calcula

    For iIndice = 1 To objGrid.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col))) <> 0 Then
            dSubTotal = dSubTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col))

        End If
    Next

    SubTotal.Caption = Format(CStr(dSubTotal), "Standard")

    lErro = Total_Calcula()
    If lErro <> SUCESSO Then Error 61654

    SubTotal_Calcula = SUCESSO

    Exit Function

Erro_SubTotal_Calcula:

    SubTotal_Calcula = Err

    Select Case Err

        Case 61654

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 166444)

    End Select

    Exit Function

End Function

Private Function Total_Calcula() As Long
'Calcula o Total

Dim dTotal As Double

    'Adiciona o SubTotal caso esteja preenchido
    If Len(Trim(SubTotal.Caption)) <> 0 And IsNumeric(SubTotal.Caption) Then dTotal = dTotal + CDbl(SubTotal.Caption)

    'Adiciona o Valor do Frete caso esteja preenchido
    If Len(Trim(ValorFrete.Text)) <> 0 And IsNumeric(ValorFrete.Text) Then dTotal = dTotal + CDbl(ValorFrete.Text)

    'Adiciona o Valor das Despesas caso esteja preenchido
    If Len(Trim(ValorDespesas.Text)) <> 0 And IsNumeric(ValorDespesas.Text) Then dTotal = dTotal + CDbl(ValorDespesas.Text)

    'Adiciona o Valor do Seguro caso esteja preenchido
    If Len(Trim(ValorSeguro.Text)) <> 0 And IsNumeric(ValorSeguro.Text) Then dTotal = dTotal + CDbl(ValorSeguro.Text)

    'Subtrai o Desconto caso esteja preenchido
    If Len(Trim(ValorDesconto.Text)) <> 0 And IsNumeric(ValorDesconto.Text) Then dTotal = dTotal - CDbl(ValorDesconto.Text)
    
    If Len(Trim(IPIValor1.Text)) > 0 And IsNumeric(IPIValor1.Text) Then dTotal = dTotal + CDbl(IPIValor1.Text)

    Total.Caption = Format(CStr(dTotal), "Standard")
    
    Total_Calcula = SUCESSO

End Function

'Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Almoxarifado do grid que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim iProdutoPreenchido As Integer
'Dim sProdutoFormatado As String
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim vbMsg As VbMsgBoxResult
'
'On Error GoTo Erro_Saida_Celula_Almoxarifado
'
'    Set objGridInt.objControle = Almoxarifado
'
'    If Len(Trim(Almoxarifado.Text)) <> 0 Then
'
'        lErro = CF("Produto_Formata",GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then Error 54569
'
'        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
'        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then Error 54567
'
'        If lErro = 25157 Then Error 54565
'
'        If lErro = 25162 Then Error 54566
'
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then Error 54568
'
'    Saida_Celula_Almoxarifado = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Almoxarifado:
'
'    Saida_Celula_Almoxarifado = Err
'
'    Select Case Err
'
'        Case 54565
'
'            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)
'
'            If vbMsg = vbYes Then
'
'                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text
'
'                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
'                Call Chama_Tela("Almoxarifado", objAlmoxarifado)
'
'            Else
'                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'            End If
'
'        Case 54566
'
'            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))
'
'            If vbMsg = vbYes Then
'
'                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)
'
'                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
'                Call Chama_Tela("Almoxarifado", objAlmoxarifado)
'
'            Else
'                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'            End If
'
'        Case 54567, 54568, 54569
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166445)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    'Verifica se Ccl foi preenchido
    If Len(Trim(Ccl.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", Ccl, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then Error 54570

        If lErro = 5703 Then Error 54571

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 54572

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err

    Select Case Err

        Case 54570
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, Ccl.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54572
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166446)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorUnitario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim dValorUnitario As Double
Dim dValorReal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim dSubTotal As Double
Dim lTamanho As Long

On Error GoTo Erro_Saida_Celula_ValorUnitario

    Set objGridInt.objControle = ValorUnitario

    'Se estiver preenchido
    If Len(Trim(ValorUnitario.ClipText)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorUnitario.Text)
        If lErro <> SUCESSO Then Error 54573

        dValorUnitario = CDbl(ValorUnitario.Text)

        'Coloca o valor Formatado na tela
        ValorUnitario.Text = Format(dValorUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 54574

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then Error 61975
    
    Saida_Celula_ValorUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorUnitario:

    Saida_Celula_ValorUnitario = Err

    Select Case Err

        Case 54573, 54574, 61975
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166447)

    End Select

    Exit Function

End Function

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_MascaraCcl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 54688

    Ccl.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_MascaraCcl:

    Inicializa_MascaraCcl = Err

    Select Case Err

        Case 54688

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166448)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    
'distribuicao
'    objGridInt.colColuna.Add ("Almoxarifado")
    
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Valor Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Valor Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    
'distribuicao
'    objGridInt.colCampo.Add (Almoxarifado.Name)

    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ValorUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)
        
    'Se é permitido que o valor unitário do ItemNF é diferente do valor unitário do ItemPC
    If gobjCOM.iNFDiferentePC = NFISCAL_NAO_ACEITA_DIFERENCA_PC Then
        ValorUnitario.Enabled = False
    Else
        ValorUnitario.Enabled = True
    End If
    
    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
'distribuicao
'    iGrid_Almoxarifado_Col = 5
    iGrid_Ccl_Col = 5
    iGrid_ValorUnitario_Col = 6
    iGrid_PercDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_ValorTotal_Col = 9
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_RECEB + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridItensPC(objGridInt As AdmGrid) As Long
'Inicializa o Grid
    
Dim bExibeTodos As Boolean
    
    If ComboPedidoCompras.Text = "TODOS" Then bExibeTodos = True
    
    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Ped Compra")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("A Receber")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Preço")
    If bExibeTodos Then
        objGridInt.colColuna.Add ("Moeda")
        objGridInt.colColuna.Add ("Taxa")
    End If
    objGridInt.colColuna.Add ("Unitário R$")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (CodigoPC.Name)
    objGridInt.colCampo.Add (ItemPC.Name)
    objGridInt.colCampo.Add (Prod.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (QuantAReceber.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    If bExibeTodos Then
        objGridInt.colCampo.Add (MoedaGrid.Name)
        objGridInt.colCampo.Add (TaxaGrid.Name)
    End If
    objGridInt.colCampo.Add (ValorRecebido.Name)

    'Colunas do Grid
    iGrid_PedCompra_Col = 1
    iGrid_Item_Col = 2
    iGrid_Prod_Col = 3
    iGrid_DescProduto_Col = 4
    iGrid_UM_Col = 5
    iGrid_AReceber_Col = 6
    iGrid_Recebido_Col = 7
    iGrid_PrecoUnitario_Col = 8
    
    If bExibeTodos Then
        
        iGrid_Moeda_Col = 9
        iGrid_Taxa_Col = 10
        iGrid_Recebido_RS_Col = 11
        
    Else
    
        iGrid_Recebido_RS_Col = 9
        
    End If

    'Grid do GridInterno
    objGridInt.objGrid = GridItensPC

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridItensPC.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItensPC = SUCESSO

    Exit Function

End Function

''''Private Function SubTotal_Calcula() As Long
'''''Soma a coluna de Valor Total e acumula em SubTotal
''''
''''Dim lErro As Long
''''Dim dSubTotal As Double
''''Dim iIndice As Integer
''''
''''On Error GoTo Erro_SubTotal_Calcula
''''
''''    For iIndice = 1 To objGrid.iLinhasExistentes
''''
''''        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col))) <> 0 Then
''''            dSubTotal = dSubTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col))
''''
''''        End If
''''
''''    Next
''''
''''    SubTotal.Caption = Format(CStr(dSubTotal), "Standard")
''''
''''    lErro = Total_Calcula
''''    If lErro <> SUCESSO Then Error 54575
''''
''''    SubTotal_Calcula = SUCESSO
''''
''''    Exit Function
''''
''''Erro_SubTotal_Calcula:
''''
''''    SubTotal_Calcula = Err
''''
''''    Select Case Err
''''
''''        Case 54575
''''
''''        Case Else
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166449)
''''
''''    End Select
''''
''''    Exit Function
''''
''''End Function

Private Sub Transportadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDespesas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)

    Call Valor_Saida(ValorFrete)

End Sub

Private Sub IPIValor1_Validate(Cancel As Boolean)

    Call Valor_Saida(IPIValor1)

End Sub

Private Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)

    Call Valor_Saida(ValorSeguro)

End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)

    Call Valor_Saida(ValorDespesas)

End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)

''''    Call Valor_Saida(ValorDesconto)

End Sub

Private Sub Valor_Saida(objControle As Object)

Dim lErro As Long

On Error GoTo Erro_Valor_Saida

    'Verifica se foi preenchido
    If Len(Trim(objControle.Text)) <> 0 Then

        'Criica se é Valor não negativo
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then Error 54576

        objControle.Text = Format(objControle.Text, "Fixed")

    End If

    lErro = Total_Calcula()
    If lErro <> SUCESSO Then Error 54577

    Exit Sub

Erro_Valor_Saida:

    Select Case Err

        Case 54576, 54577
            objControle.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166450)

    End Select

    Exit Sub

End Sub

Public Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 54633
        
        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado
        
        objProduto.sCodigo = sProdutoFormatado
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89658

        'Se não achou o Produto --> erro
        If lErro = 28030 Then gError 89659
        
        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_Descricao_Col)
        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then objItemNF.dQuantidade = CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) > 0 Then objItemNF.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        lTamanho = Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col)))
        
        If lTamanho > 0 Then objItemNF.dPercDesc = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col))

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))) > 0 Then objItemNF.dValorDesconto = CDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))

'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) > 0 Then
'
'            objItemNF.sAlmoxarifadoNomeRed = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
'            objAlmoxarifado.sNomeReduzido = objItemNF.sAlmoxarifadoNomeRed
'
'            'Busca o Código do Almoxarifado através do Nome Reduzido
'            lErro = CF("Almoxarifado_Le_NomeReduzido",objAlmoxarifado)
'            If lErro <> SUCESSO And lErro <> 25060 Then gError 54634
'
'            If lErro = 25060 Then gError 54636
'
'            objItemNF.iAlmoxarifado = objAlmoxarifado.iCodigo
'
'        End If
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Ccl_Col))) > 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", GridItens.TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 54635
            
            objItemNF.sCcl = sCclFormatada
            
        Else

            objItemNF.sCcl = ""

        End If
        
        'Adiciona na coleção de Ítens
        With objItemNF
            'distribuicao. retirado o codigo do almoxarifado
            objNFiscal.ColItensNF.Add 0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, 0, "", .sCcl, STATUS_LANCADO, 0, "", 0, 0, 0, objProduto.sSiglaUMEstoque, objProduto.iClasseUM, 0
        End With

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 54633, 54634, 54635, 89658

        Case 54636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 89659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166451)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long
'Verifica os dados para gravação de Recebimento de Material de Fornecedor

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objNFiscal As New ClassNFiscal
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objItemNF As ClassItemNF
Dim objItemNFItemPC As ClassItemNFItemPC
Dim colCodPedCompras As New Collection
Dim dQuantidade As Double
Dim iLinha As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim dTotal As Double
Dim vbMsg As VbMsgBoxResult
Dim dFator As Double
Dim dQuantidadeComMargem As Double 'Inserido por Wagner

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Chama Verifica_Preenchimento
    lErro = Verifica_Preenchimento()
    If lErro <> SUCESSO Then gError 54582

    'Critica os Valores
    lErro = Critica_Valores(objNFiscal)
    If lErro <> SUCESSO Then gError 61594
    
    lErro = Confere_Quantidade_PrecoTotal()
    If lErro <> SUCESSO Then gError 114501

    'Verifica se algum Item no Grid
    If objGrid.iLinhasExistentes = 0 Then gError 54578
    
    'Critica Total
    lErro = Total_Calcula()
    If lErro <> SUCESSO Then gError 61595
    
    'Valida os dados do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 65586

        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 61912

        objProduto.sCodigo = sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 61913
        If lErro = 28030 Then gError 61914

        dQuantidade = 0
        dQuantidadeComMargem = 0 'Inserido por Wagner
        
        For iLinha = 1 To objGridItensPC.iLinhasExistentes

            'Se a taxa nao estiver preenchida => Erro
            'If Len(Trim(GridItensPC.TextMatrix(iLinha, iGrid_Taxa_Col))) = 0 Then gError 114512
            
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridItensPC.TextMatrix(iLinha, iGrid_Prod_Col) Then
            
                'Converte a UM de GridItensPC para a UM do GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItensPC.TextMatrix(iLinha, iGrid_UM_Col), GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 89284
            
                Set objItemPCInfo = gcolItemPedCompraInfo.Item(iLinha)
            
                'Alterado por Wagner
                dQuantidade = dQuantidade + StrParaDbl(GridItensPC.TextMatrix(iLinha, iGrid_Recebido_Col)) * dFator '* (1 + objItemPCInfo.dPercentMaisReceb / 100)
                
                '#######################################################
                'Inserido por Wagner
                dQuantidadeComMargem = dQuantidadeComMargem + objItemPCInfo.dQuantPedida * dFator * (1 + objItemPCInfo.dPercentMaisReceb)
                dQuantidadeComMargem = dQuantidadeComMargem - (objItemPCInfo.dQuantPedida - CDbl(GridItensPC.TextMatrix(iLinha, iGrid_AReceber_Col)) * dFator)
                
                If dQuantidade - dQuantidadeComMargem > QTDE_ESTOQUE_DELTA Then gError 136840
                '#######################################################
                
            End If
            
        Next

        sCodProduto = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
        
        'Se for diferente, Erro
        If Abs(dQuantidade - StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > QTDE_ESTOQUE_DELTA Then gError 54585

'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 61590
        
'        'não pode ser um produto não comprável
'        If objProduto.iCompras = PRODUTO_NAO_COMPRAVEL Then gError 61591

        'Verifica se a Unidade de Medida foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 61593

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 61596

    Next
    
    dTotal = CDbl(IIf(Len(Trim(Total.Caption)) > 0, Total.Caption, 0))

    'Se o total for negativo --> Erro
    If dTotal < 0 Then gError 54579
    
    'Valida e recolhe os dados do grid de Recebimento
    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 54583

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 54580

    'distribuicao
    lErro = gobjDistribuicao.Move_GridDist_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 89655
    
    Call Atualiza_ItensPC
    
    'Transfere os dados de gColItemPedCompraInfo para os itens do Recebimento
    For Each objItemNF In objNFiscal.ColItensNF

        Set objItemNF.colItemNFItemPC = New Collection
        For Each objItemPCInfo In gcolItemPedCompraInfo

            If objItemNF.sProduto = objItemPCInfo.sProduto Then
                Set objItemNFItemPC = New ClassItemNFItemPC

                objItemNFItemPC.dQuantidade = objItemPCInfo.dQuantRecebida
                objItemNFItemPC.lItemPedCompra = objItemPCInfo.lNumIntDoc
                objItemNFItemPC.dTaxa = objItemPCInfo.dTaxa

                objItemNF.colItemNFItemPC.Add objItemNFItemPC
            End If

        Next
    Next

    'Guarda na coleção colCodPedCompras os Códigos dos pedidos selecionados em colCodPedCompras
    For iIndice = 0 To PedidosCompra.ListCount - 1
        If PedidosCompra.Selected(iIndice) = True Then
            colCodPedCompras.Add PedidosCompra.List(iIndice)
        End If
    Next

    lErro = CF("RecebMaterialFCom_Grava", objNFiscal, colCodPedCompras)
    If lErro <> SUCESSO Then gError 54581

    GL_objMDIForm.MousePointer = vbDefault
    
    If Len(Trim(NumRecebimento.Caption)) = 0 Then vbMsg = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_RECEBIMENTO_GRAVADO", objNFiscal.lNumRecebimento)
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 54578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSRECEB_NAO_INFORMADOS", gErr)

        Case 54579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 54580, 54581, 54582, 54583, 61594, 61595, 61912, 61913, 89284, 89655
        
        Case 54585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_DIFERENT_QUANTRECEBIDA", gErr, sCodProduto)
        
        Case 61590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
'        Case 61591
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
            
        Case 61593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 61596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 61914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 65586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 114501
        
        Case 114512
            Call Rotina_Erro(vbOKOnly, "ERRO_TAXA_GRID_IMCOMPLETA", gErr)
            
        '#########################
        'Inserido por Wagner
        Case 136840
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_MAIOR_SALDO_PC", gErr, sCodProduto)
        '#########################
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166452)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Function Verifica_Preenchimento() As Long
'Verifica se os principais campos da tela foram preenchidos

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer

On Error GoTo Erro_Verifica_Preenchimento

    'Verifica se o fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 54619

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Error 54620

    'Verifica se a DataEntrada foi preenchida
    If Len(Trim(DataEntrada.Text)) = 0 Then Error 54621

    'Verifica se foi selecionado algum Tipo de Nota fiscal
    If Not NFiscalForn.Value And Not NFiscalPropria.Value Then Error 54622

    'Verifica se a Série de NotaFiscal Propria está Cadastrada no BD
    If NFiscalForn.Value = True Then

        'Verifica se a Série foi preenchida
        If Len(Trim(Serie.Text)) = 0 Then Error 54623

        'Verifica se a Nota Fiscal foi preenchida
        If Len(Trim(NFiscal.Text)) = 0 Then Error 54624

        For iIndice = 0 To Serie.ListCount - 1

            If Serie.Text = Serie.List(iIndice) Then
                iAchou = 1
                Exit For
            End If
        Next

        If iAchou = 0 Then Error 54625

    End If

    'Verifica se o PesoBruto é maior que PesoLiq
    If Len(Trim(PesoLiquido.Text)) <> 0 And Len(Trim(PesoBruto.Text)) <> 0 Then

        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then Error 54626

    End If

    'Filial Compra não preenchida
    If Len(Trim(FilialCompra.Text)) = 0 Then Error 61592

    Verifica_Preenchimento = SUCESSO

    Exit Function

Erro_Verifica_Preenchimento:

    Verifica_Preenchimento = Err

    Select Case Err

        Case 54619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 54620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 54621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", Err)

        Case 54622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_INFORMADO", Err)

        Case 54623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case 54624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", Err)

        Case 54625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case 54626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESOBRUTO_MENOR_PESOLIQ", Err, CDbl(PesoBruto.Text), CDbl(PesoLiquido.Text))

        Case 61592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166453)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_RecebMaterialFCom()
'Limpa a tela de Recebimento de Material de Fornecedor

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    PedidosCompra.Clear

    'Limpa os Grids
    Call Grid_Limpa(objGrid)
    Call Grid_Limpa(objGridItensPC)
    
    Set gcolPedidoCompra = New Collection
    Set gcolItemPedCompraInfo = New Collection

    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao

    'Limpa o Label's
    SubTotal.Caption = ""
    Total.Caption = ""

    'Limpa e desseleciona a Combo Série
    Serie.Text = ""
    Serie.ListIndex = -1

    'Desseleciona as combos Transportadora e PlacaUF
    Transportadora.ListIndex = -1
    PlacaUF.Text = ""

    'Incluído por Luiz Nogueira em 21/08/03
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)

    'Limpa a Combo Filial
    Filial.Clear

    'Preenche a DataEntrada com a Data Atual
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    NFiscalPropria.Value = False
    NFiscalForn.Value = False
    Emitente.Value = True
    NumRecebimento.Caption = ""
  
    ComboPedidoCompras.Clear
    ComboPedidoCompras.AddItem "TODOS"
    
    Moeda.ListIndex = -1
    ComboPedidoCompras.ListIndex = 0
    
    Taxa.Text = ""
    Taxa.Enabled = False
    LabelTaxa.Enabled = False
    
    Set gcolItemPedCompraInfo = New Collection
    Set gColInfoBD = New Collection

    iAlterado = 0

End Sub

Private Sub VolumeEspecie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeEspecie_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie, "AVISO_CRIAR_VOLUMEESPECIE")
    If lErro <> SUCESSO Then gError 102440
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102440
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166454)

    End Select

End Sub

Private Sub VolumeMarca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeMarca_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca, "AVISO_CRIAR_VOLUMEMARCA")
    If lErro <> SUCESSO Then gError 102441
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102441
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166455)

    End Select

End Sub

Private Sub VolumeQuant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VolumeNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUMAnterior As String

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    'recolhe a UM anteriormente escolhida
    sUMAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text


    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 54627

    If Trim(sUMAnterior) <> Trim(UnidadeMed.Text) Then

        lErro = Atualiza_UM(GridItens.Row, sUMAnterior)
        If lErro <> SUCESSO Then gError 54628

'distribuicao
        lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 89666

    End If

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 54627, 54628, 89656, 89666
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166456)

    End Select

    Exit Function

End Function

Private Function Atualiza_UM(iLinha As Integer, sUM As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Atualiza_UM

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 54702

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridItens
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 54706
    If lErro = 28030 Then Error 54630

    'Converte a UM de GridItensPC para a UM do GridItens
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dFator)
    If lErro <> SUCESSO Then Error 54629

    'Atualiza o Grid de Recebimento
    If (StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) * dFator) <> 0 Then GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) * dFator)
    GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col)) / dFator, FORMATO_PRECO_UNITARIO_EXTERNO)

    Atualiza_UM = SUCESSO

    Exit Function

Erro_Atualiza_UM:

    Atualiza_UM = Err

    Select Case Err

        Case 54629, 54702, 54706

        Case 54630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166457)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal, Optional iGravacao = 1) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê Fornecedor no BD
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 54631

        'Se não achou o Fornecedor --> erro
        If lErro = 6681 Then Error 54632

        objNFiscal.lFornecedor = objFornecedor.lCodigo

    End If
    
    objNFiscal.lNumRecebimento = StrParaLong(NumRecebimento.Caption)
    objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)

    objNFiscal.dtDataEntrada = MaskedParaDate(DataEntrada)
    
'horaentrada
    If Len(Trim(HoraEntrada.ClipText)) > 0 Then
        objNFiscal.dtHoraEntrada = CDate(HoraEntrada.Text)
    Else
        objNFiscal.dtHoraEntrada = Time
    End If
    
    objNFiscal.dtDataEmissao = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.dtDataSaida = DATA_NULA

    If NFiscalPropria.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRFPCO
    ElseIf NFiscalForn.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRFFCO
    Else
        objNFiscal.iTipoNFiscal = 0
    End If

    objNFiscal.sSerie = Serie.Text
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iFilialPedido = Codigo_Extrai(FilialCompra.Text)
    
    objNFiscal.dValorProdutos = StrParaDbl(SubTotal.Caption)
    objNFiscal.dValorTotal = StrParaDbl(Total.Caption)
    objNFiscal.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objNFiscal.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objNFiscal.dValorFrete = StrParaDbl(ValorFrete.Text)
    objNFiscal.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)

    objNFiscal.lNumIntDoc = 0

    objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)

    'Armazena o responsável pelo frete
    If Emitente.Value Then
        objNFiscal.iFreteRespons = FRETE_EMITENTE
    Else
        objNFiscal.iFreteRespons = FRETE_DESTINATARIO
    End If

    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    objNFiscal.sVolumeNumero = VolumeNumero.Text

    objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)

    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sObservacao = Observacao.Text
    
    objNFiscal.dPesoLiq = StrParaDbl(PesoLiquido.Text)
    objNFiscal.dPesoBruto = StrParaDbl(PesoBruto.Text)

    objNFiscal.iFilialPedido = Codigo_Extrai(FilialCompra.Text)

    objNFiscal.iStatus = STATUS_LANCADO

    Set objNFiscal.objTributacaoNF = New ClassTributacaoDoc
    objNFiscal.objTributacaoNF.dIPIValor = StrParaDbl(IPIValor1.Text)
''''    Call Preenche_Tributacao(objNFiscal)

    objNFiscal.dtDataRegistro = gdtDataAtual
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 54631

        Case 54632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166458)

    End Select

    Exit Function

End Function

Function Preenche_GridItensPC() As Long
'Preenche GridItensPC

Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim iLinha As Integer
Dim iItem As Integer
Dim lErro As Long
Dim lPedCompra As Long
Dim sProdutoMascarado As String
Dim iIndice As Integer
Dim bTodos As Boolean

On Error GoTo Erro_Preenche_GridItensPC

    Call Grid_Limpa(objGridItensPC)
    
    Set objGridItensPC = New AdmGrid
    
    Call Inicializa_GridItensPC(objGridItensPC)
        
    If ComboPedidoCompras.List(ComboPedidoCompras.ListIndex) = "TODOS" Then bTodos = True
    
    'Para cada Item da coleção de ItensPC
    For Each objItemPedCompraInfo In gcolItemPedCompraInfo
        
        If Not bTodos Then
        
            If objItemPedCompraInfo.lPedCompra = ComboPedidoCompras.List(ComboPedidoCompras.ListIndex) Then
    
                'Incrementa o número de linhas
                iLinha = iLinha + 1
                iItem = iItem + 1
                
                'Se o Pedido de Compras mudou, Inicia novamente o contador de Itens
                If objItemPedCompraInfo.lPedCompra <> lPedCompra And iLinha > 1 Then iItem = 1
                
                GridItensPC.TextMatrix(iLinha, iGrid_PedCompra_Col) = CStr(objItemPedCompraInfo.lPedCompra)
                GridItensPC.TextMatrix(iLinha, iGrid_Item_Col) = iItem
                
                'Mascara o Produto
                lErro = Mascara_RetornaProdutoEnxuto(objItemPedCompraInfo.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 54675
                
                Prod.PromptInclude = False
                Prod.Text = sProdutoMascarado
                Prod.PromptInclude = True
        
                sProdutoMascarado = Prod.Text
                
                GridItensPC.TextMatrix(iLinha, iGrid_Prod_Col) = sProdutoMascarado
                GridItensPC.TextMatrix(iLinha, iGrid_DescProduto_Col) = objItemPedCompraInfo.sDescProduto
                GridItensPC.TextMatrix(iLinha, iGrid_UM_Col) = objItemPedCompraInfo.sUM
                GridItensPC.TextMatrix(iLinha, iGrid_AReceber_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantReceber)
                GridItensPC.TextMatrix(iLinha, iGrid_Recebido_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantRecebida)
                GridItensPC.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(objItemPedCompraInfo.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
                
                If gbLimpaTaxa Then Taxa.Text = ""
                
                If objItemPedCompraInfo.dTaxa = 0 And StrParaDbl(Taxa.Text) > 0 Then
                    objItemPedCompraInfo.dTaxa = StrParaDbl(Taxa.Text)
                    objItemPedCompraInfo.bTaxaPedido = True
                End If
                
                If ComboPedidoCompras.Text <> "TODOS" And objItemPedCompraInfo.dTaxa > 0 Then
                    Taxa.Text = Format(objItemPedCompraInfo.dTaxa, "STANDARD")
                End If
                
                If StrParaDbl(Taxa.Text) > 0 Then GridItensPC.TextMatrix(iLinha, iGrid_Recebido_RS_Col) = Format(objItemPedCompraInfo.dPrecoUnitario * objItemPedCompraInfo.dTaxa, FORMATO_PRECO_UNITARIO_EXTERNO)
                
                'Guarda o último Pedido de Compras
                lPedCompra = objItemPedCompraInfo.lPedCompra
                
                'Seleciona a moeda e a taxa (caso exista ...)
                For iIndice = 0 To Moeda.ListCount - 1
                    If Moeda.ItemData(iIndice) = objItemPedCompraInfo.iMoeda Then
                        Moeda.ListIndex = iIndice
                        Exit For
                    End If
                Next
                    
                'Se a taxa foi preenchida => Mostra na tela ...
                If objItemPedCompraInfo.dTaxa > 0 Then
                    
                    Taxa.Text = Format(objItemPedCompraInfo.dTaxa, "STANDARD")
                    
                    If objItemPedCompraInfo.bTaxaPedido Then
                        Taxa.Enabled = True
                        LabelTaxa.Enabled = True
                    Else
                        Taxa.Enabled = False
                        LabelTaxa.Enabled = False
                    End If
                
                'Apenas por Seguranca ...
                ElseIf objItemPedCompraInfo.dTaxa = 0 Then
                    
                    Taxa.Text = ""
                    Taxa.Enabled = True
                    LabelTaxa.Enabled = True
                    
                End If
                
                If objItemPedCompraInfo.iMoeda = MOEDA_REAL Then
                    Taxa.Text = ""
                    Taxa.Enabled = False
                    LabelTaxa.Enabled = False
                End If
                
            End If
                
        Else
            
            'Incrementa o número de linhas
            iLinha = iLinha + 1
            iItem = iItem + 1
            
            'Se o Pedido de Compras mudou, Inicia novamente o contador de Itens
            If objItemPedCompraInfo.lPedCompra <> lPedCompra And iLinha > 1 Then iItem = 1
            
            GridItensPC.TextMatrix(iLinha, iGrid_PedCompra_Col) = CStr(objItemPedCompraInfo.lPedCompra)
            GridItensPC.TextMatrix(iLinha, iGrid_Item_Col) = iItem
            
            'Mascara o Produto
            lErro = Mascara_RetornaProdutoEnxuto(objItemPedCompraInfo.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 54675
            
            Prod.PromptInclude = False
            Prod.Text = sProdutoMascarado
            Prod.PromptInclude = True
    
            sProdutoMascarado = Prod.Text
            
            GridItensPC.TextMatrix(iLinha, iGrid_Prod_Col) = sProdutoMascarado
            GridItensPC.TextMatrix(iLinha, iGrid_DescProduto_Col) = objItemPedCompraInfo.sDescProduto
            GridItensPC.TextMatrix(iLinha, iGrid_UM_Col) = objItemPedCompraInfo.sUM
            GridItensPC.TextMatrix(iLinha, iGrid_AReceber_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantReceber)
            GridItensPC.TextMatrix(iLinha, iGrid_Recebido_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantRecebida)
            GridItensPC.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(objItemPedCompraInfo.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
            
            For iIndice = 0 To Moeda.ListCount - 1
                If Moeda.ItemData(iIndice) = objItemPedCompraInfo.iMoeda Then
                    GridItensPC.TextMatrix(iLinha, iGrid_Moeda_Col) = Moeda.List(iIndice)
                    Exit For
                End If
            Next
            
            If objItemPedCompraInfo.dTaxa > 0 Then GridItensPC.TextMatrix(iLinha, iGrid_Taxa_Col) = Format(objItemPedCompraInfo.dTaxa, "STANDARD")
            
            Taxa.Text = ""
            Taxa.Enabled = False
            LabelTaxa.Enabled = False
                
            If objItemPedCompraInfo.dTaxa = 0 And StrParaDbl(Taxa.Text) > 0 Then
                objItemPedCompraInfo.dTaxa = StrParaDbl(Taxa.Text)
                objItemPedCompraInfo.bTaxaPedido = True
            End If
            
            If ComboPedidoCompras.Text <> "TODOS" Then
                Taxa.Text = Format(objItemPedCompraInfo.dTaxa, "STANDARD")
            End If
            
            If objItemPedCompraInfo.dTaxa > 0 Then GridItensPC.TextMatrix(iLinha, iGrid_Recebido_RS_Col) = Format(objItemPedCompraInfo.dPrecoUnitario * objItemPedCompraInfo.dTaxa, FORMATO_PRECO_UNITARIO_EXTERNO)
            
            'Guarda o último Pedido de Compras
            lPedCompra = objItemPedCompraInfo.lPedCompra
            
        End If
        
    Next
    
    'Alteração das linhas existentes
    objGridItensPC.iLinhasExistentes = gcolItemPedCompraInfo.Count
    
''''    If Not gbCarregandoTela Then
''''
''''        lErro = CalculaPrecoUnitario_GridItens
''''        If lErro <> SUCESSO Then gError 114511
''''
''''    End If
        
    Preenche_GridItensPC = SUCESSO

    Exit Function

Erro_Preenche_GridItensPC:

    Preenche_GridItensPC = gErr

    Select Case gErr

        Case 54675, 114511
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166459)

    End Select

    Exit Function

End Function

Function Traz_RecebMaterialFCom_Tela(objNFiscal As ClassNFiscal) As Long
'Preenche a tela com os dados passados como parâmetro em objNFiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objFornecedor As New ClassFornecedor
Dim objPedidoCompras As ClassPedidoCompras
Dim colPedCompras As New Collection
Dim bCancel As Boolean

On Error GoTo Erro_Traz_RecebMaterialFCom_Tela

    gbCarregandoTela = True
    
    NumRecebimento.Caption = objNFiscal.lNumRecebimento
    
    'Lê os ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 54643

    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89660

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Label's
    SubTotal.Caption = ""
    Total.Caption = ""

    'Seleciona NFiscalPropria
    NFiscalPropria.Value = True

    'Lê o NomeReduzido do Fornecedor no BD
    objFornecedor.lCodigo = objNFiscal.lFornecedor

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 54644

    'Se não achou o Fornecedor --> Erro
    If lErro = 12729 Then gError 54649

    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Dispara o Validate de Fornecedor
    Call Fornecedor_Validate(bCancel)

    Filial.Text = CStr(objNFiscal.iFilialForn)

    Call Filial_Validate(bCancel)

    Call DateParaMasked(DataEntrada, objNFiscal.dtDataEntrada)

'horaentrada
    HoraEntrada.PromptInclude = False
    'este teste está correto
    If objNFiscal.dtDataEntrada <> DATA_NULA Then HoraEntrada.Text = Format(objNFiscal.dtHoraEntrada, "hh:mm:ss")
    HoraEntrada.PromptInclude = True

    If objNFiscal.iTipoNFiscal = DOCINFO_NRFPCO Then
        NFiscalPropria.Value = True
    Else
        NFiscalForn.Value = True
    End If

    Serie.Text = objNFiscal.sSerie
    If objNFiscal.lNumNotaFiscal > 0 Then
        NFiscal.Text = CStr(objNFiscal.lNumNotaFiscal)
    End If
    
    NFiscal.Text = CStr(objNFiscal.lNumNotaFiscal)

    Call CF("Filial_Seleciona", FilialCompra, objNFiscal.iFilialPedido)

    If Len(Trim(FilialCompra.Text)) > 0 And Len(Trim(Filial.Text)) > 0 And Len(Trim(Fornecedor.Text)) > 0 Then
        Call Atualiza_ListaPedidos(objNFiscal)
    End If

    'Preenche o Tab Complemento
    'Seleciona a Transportadora através do Código no ItemData
    For iIndice = 0 To Transportadora.ListCount - 1
        If Transportadora.ItemData(iIndice) = objNFiscal.iCodTransportadora Then
            Transportadora.ListIndex = iIndice
            Exit For
        End If
    Next

    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    VolumeQuant.Text = CStr(objNFiscal.lVolumeQuant)
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        VolumeEspecie.Text = ""
    End If
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        VolumeMarca.Text = ""
    End If
    
    VolumeNumero.Text = objNFiscal.sVolumeNumero
    Mensagem.Text = objNFiscal.sMensagemNota
    PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.iFreteRespons = FRETE_EMITENTE Then
        Emitente.Value = True
    Else
        Destinatario.Value = True
    End If
    
    VolumeNumero.Text = objNFiscal.sVolumeNumero

    'Limpa o Grid de recebimento
    Call Grid_Limpa(objGrid)

    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 54645

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 89661

    'Lê apenas os Pedidos de Compras não baixados relacionados a nota fiscal
    lErro = CF("PedidoCompra_Le_Recebimento", objNFiscal, colPedCompras)
    If lErro <> SUCESSO And lErro <> 66136 Then gError 66400
    
    'Seleciona os Pedidos de Compras associados a Nota Fiscal
    For Each objPedidoCompras In colPedCompras
        For iIndice = 0 To PedidosCompra.ListCount - 1
            If objPedidoCompras.lCodigo = PedidosCompra.List(iIndice) Then
                PedidosCompra.Selected(iIndice) = True
            End If
        Next
    Next
    
    Set gcolItemPedCompraInfo = New Collection
    
    'Lê os Itens de Pedido de Compras associados a Nota Fiscal
    lErro = CF("ItensPedCompra_Le_NFiscalReceb", objNFiscal, gcolItemPedCompraInfo)
    If lErro <> SUCESSO Then gError 54692

    'Limpa o Grid de Pedido de Compras
    Call Grid_Limpa(objGridItensPC)

    'Coloca para mostrar todos os pedidos ...
    lErro = Preenche_GridItensPC()
    If lErro <> SUCESSO Then gError 108996
    
    SubTotal.Caption = CStr(objNFiscal.dValorProdutos)
    ValorDesconto.Text = CStr(objNFiscal.dValorDesconto)
    ValorSeguro.Text = CStr(objNFiscal.dValorSeguro)
    ValorFrete.Text = CStr(objNFiscal.dValorFrete)
    ValorDespesas.Text = CStr(objNFiscal.dValorOutrasDespesas)
    IPIValor1.Text = Format(objNFiscal.dValorTotal - objNFiscal.dValorProdutos + objNFiscal.dValorDesconto - objNFiscal.dValorSeguro - objNFiscal.dValorFrete - objNFiscal.dValorOutrasDespesas, "Standard")
        
    lErro = SubTotal_Calcula()
    If lErro <> SUCESSO Then gError 54647


    iAlterado = 0

    gbCarregandoTela = False
    
    Traz_RecebMaterialFCom_Tela = SUCESSO

    Exit Function

Erro_Traz_RecebMaterialFCom_Tela:

    Traz_RecebMaterialFCom_Tela = gErr

    Select Case gErr

        Case 54643, 54644, 54645, 54646, 54647, 54648, 54692, 54723, 89660, 89661

        Case 54649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166460)

    End Select

    Exit Function

End Function

Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os ítens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoMascarado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sCclMascarado As String
Dim iIndice2 As Integer
Dim dValorTotal As Double

On Error GoTo Erro_Preenche_GridItens

    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 54650
        
        Prod.PromptInclude = False
        Prod.Text = sProdutoMascarado
        Prod.PromptInclude = True

        sProdutoMascarado = Prod.Text
        
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        For iIndice2 = 0 To Produto.ListCount - 1
            If Produto.List(iIndice) = sProdutoMascarado Then
                Produto.ListIndex = iIndice
                Exit For
            End If
        Next
        
        If iIndice2 = Produto.ListCount Then
            Produto.AddItem sProdutoMascarado
            Produto.ListIndex = iIndice2
        End If
                
        sCclMascarado = ""
        
        'Formata o Ccl
        If Trim(objItemNF.sCcl) <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_RetornaCclEnxuta(objItemNF.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 54651
            
            'Preenche o campo Ccl com o Ccl encontrado
            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True
    
            'Joga o Ccl encontrado no Grid
            GridItens.TextMatrix(iIndice, iGrid_Ccl_Col) = Ccl.Text

        End If

        'Preenche o Grid
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        If objItemNF.dPercDesc <> 0 Then GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        If objItemNF.dValorDesconto <> 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")

    
        'Coloca o Valor Total na Coluna correspondente
        GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col) = Format(CStr(objItemNF.dValorTotal - objItemNF.dValorDesconto), "Standard")
        
        'Atualiza o número de linhas existentes
        objGrid.iLinhasExistentes = iIndice
    
    Next

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr
        
        Case 54650
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
        
        Case 54651
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objItemNF.sCcl)

        Case 54652
        
        Case 54653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166461)

    End Select

    Exit Function

End Function

Private Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 54654

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 54655

        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then Error 54656

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 54657

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case Err

        Case 54654, 54655

        Case 54656  'Não encontrou Transportadora no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 54657

            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166462)

    End Select

    Exit Sub

End Sub

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim lTamanho As Long
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then
        
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then Error 54658

        dPercentDesc = CDbl(PercentDesc.Text)

        lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
        If lTamanho > 0 Then dPercentDescAnterior = StrParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

        If dPercentDesc <> dPercentDescAnterior Then

            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then Error 54660

            PercentDesc.Text = Format(dPercentDesc, "Fixed")

        End If
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
     If lErro <> SUCESSO Then Error 54659

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then Error 54560

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = Err

    Select Case Err

        Case 54560, 54658, 54659
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166463)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dPrecoTotal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim iDescontoAlterado As Integer

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto

    iDescontoAlterado = False
    
    'Veifica o preenchimento de Desconto
    If Len(Trim(Desconto.ClipText)) > 0 Then

        'Critica o Desconto
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then Error 54662

        dDesconto = CDbl(Desconto.Text)
        
        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) <> dDesconto Then iDescontoAlterado = True

        If iDescontoAlterado = True Then

            dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
            dPrecoTotal = dQuantidade * dPrecoUnitario

            If dPrecoTotal > 0 Then
    
                If dDesconto >= dPrecoTotal Then Error 54661
            
                dPercentDesc = dDesconto / dPrecoTotal

                GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

            End If
        
        End If
    
    Else
    
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) <> 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = ""
        End If
            
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 54663

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then Error 61976

    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = Err

    Select Case Err

        Case 54661
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", Err, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54662, 54663, 61976
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166464)

    End Select

    Exit Function

End Function

Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = dValorUnitario * dQuantidade

    'Se o Percentual Desconto estiver preenchido
    If dPercentDesc > 0 Then

        'Testa se o desconto está preenchido
        If dDesconto = 0 Then
            dPercDesc2 = 0
        Else
            'Calcula o Percentual em cima dos valores passados
            dPercDesc2 = dDesconto / dValorTotal
            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
        End If
        'se os percentuais (passado e calculado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = dPercentDesc * dValorTotal

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Private Sub Fornecedor_GotFocus()

    sFornecedorAnterior = Trim(Fornecedor.Text)

End Sub

Private Sub PedidosCompra_ItemCheck(Item As Integer)
 
Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objItemPC As ClassItemPedCompra
Dim objProduto As New ClassProduto
Dim colReqCompras As New Collection
Dim iIndice As Integer
Dim iCont As Integer
Dim iProdutoEncontrado As Integer
Dim dFator As Double
Dim dQuantidade As Double
Dim dQuantAtual As Double
Dim sProdutoMascarado As String
Dim iItem As Integer
Dim lReqCompra As Long
Dim sProduto As String
Dim dDescontoAtual As Double
Dim dValorUnitario As Double
Dim dValorTotal As Double
Dim dQuantidadeAux As Double
Dim dPrecoAux As Double

On Error GoTo Erro_PedidosCompra_ItemCheck
    
    'Se não existem pedidos na list --> Sai
    If PedidosCompra.ListCount = 0 Then Exit Sub
    
    'Atualiza as quantidades das coleções globais
    Call Atualiza_ItensPC
    
    'Se o pedido clicado na Lista de Pedidos de Compra está marcado
    If PedidosCompra.Selected(Item) = True Then

        'Recolhe o pedido da Coleção global
        Set objPedidoCompra = gcolPedidoCompra.Item(Item + 1)

        'Adiciona na combo de pedido do frame 3
        ComboPedidoCompras.AddItem PedidosCompra.List(PedidosCompra.ListIndex)
        
        'Se não está trazendo uma Nota Fiscal para a tela
        If gbCarregandoTela = False Then
        
            'Lê os itens do Pedido de Compra passado por objPedidoCompra
            lErro = CF("ItensPC_Le_Codigo", objPedidoCompra)
            If lErro <> SUCESSO Then gError 65786
    
            'Para cada Item do Pedido de Compras
            For Each objItemPC In objPedidoCompra.colItens
    
                iItem = iItem + 1
    
                Set objItemPCInfo = New ClassItemPedCompraInfo
    
                sProdutoMascarado = String(STRING_PRODUTO, 0)
                
                lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 65787
                
                Prod.PromptInclude = False
                Prod.Text = sProdutoMascarado
                Prod.PromptInclude = True
        
                sProdutoMascarado = Prod.Text
                
                objProduto.sCodigo = objItemPC.sProduto

                'Lê o Produto passado por objProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 65788
                If lErro = 28030 Then gError 65789
                   
                iProdutoEncontrado = 0
    
                Set objItemPCInfo = New ClassItemPedCompraInfo
                
                
                If objItemPC.iMoeda = MOEDA_REAL Then objItemPC.dTaxa = 1
                
                'Preenche o objItemPCInfo a partir do objItemPC
                objItemPCInfo.dPercentMaisReceb = objItemPC.dPercentMaisReceb
                objItemPCInfo.dQuantReceber = objItemPC.dQuantidade - objItemPC.dQuantRecebida
                objItemPCInfo.dQuantRecebida = objItemPC.dQuantidade - objItemPC.dQuantRecebida
                objItemPCInfo.dQuantPedida = objItemPC.dQuantidade
                objItemPCInfo.iItem = iItem
                objItemPCInfo.iRecebForaFaixa = objItemPC.iRebebForaFaixa
                objItemPCInfo.lNumIntDoc = objItemPC.lNumIntDoc
                objItemPCInfo.lPedCompra = objPedidoCompra.lCodigo
                objItemPCInfo.sDescProduto = objItemPC.sDescProduto
                objItemPCInfo.sProduto = objItemPC.sProduto
                objItemPCInfo.sUM = objItemPC.sUM
                objItemPCInfo.dPrecoUnitario = objItemPC.dPrecoUnitario
                objItemPCInfo.dAliquotaIPI = objItemPC.dAliquotaIPI
                objItemPCInfo.dAliquotaICMS = objItemPC.dAliquotaICMS
                objItemPCInfo.iMoeda = objItemPC.iMoeda
                objItemPCInfo.dTaxa = objItemPC.dTaxa
                objItemPCInfo.dPrecoUnitario = objItemPC.dPrecoUnitario
                objItemPCInfo.dValorDesconto = objItemPC.dValorDesconto

                
                'Verifica se o produto está presente no grid de itens de recebimento
                For iIndice = 1 To objGridItens.iLinhasExistentes
    
                    'Se encontrou o produto no grid de Recebimento
                    If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado Then
    
                        'Atualiza a Quantidade do Grid de Recebimento
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                        If lErro <> SUCESSO Then gError 65790
    
                        dQuantidade = (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * dFator
                        objItemPCInfo.dQuantReceber = dQuantidade
                        objItemPCInfo.dQuantRecebida = dQuantidade
                        objItemPCInfo.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
                        
                        'Calcula a quantidade Atual levanod em conta o novo item inserido.
                        dQuantAtual = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) + dQuantidade
                                               
                        'Calcula por proporção um novo preço unitário
                        dPrecoAux = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) * StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
                        dPrecoAux = dPrecoAux + objItemPCInfo.dQuantReceber * objItemPC.dPrecoUnitario * objItemPC.dTaxa
                        dPrecoAux = dPrecoAux / dQuantAtual
                        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(dPrecoAux, IIf(gobjCOM.sFormatoPrecoUnitario <> "", gobjCOM.sFormatoPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO))
                                               
                        'Atualiza a quantidade do item de Recebimento do produto
                        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(dQuantAtual)
    
                        'calcula o novo percentual e valor de desconto
                        If objItemPC.dValorDesconto <> 0 Then
                        
                            dDescontoAtual = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) + (objItemPC.dValorDesconto * objItemPC.dTaxa)
                            
                            GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(dDescontoAtual, "Standard")
                            
                            GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dDescontoAtual / (dQuantAtual * dPrecoAux), "Percent")
                                
                        End If
    
                        lErro = Calcula_Valores(iIndice)
                        If lErro <> SUCESSO Then gError 89179
    
                        Exit For
    
                    End If
    
                Next
    
                'Se não encontrou o Produto
                If iIndice > objGridItens.iLinhasExistentes Then
                    
                    'Se o produto possui quantidade a receber
                    If objItemPCInfo.dQuantReceber > 0 Then
                        
                        'Preenche uma linha do GridItens
                        GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Produto_Col) = sProdutoMascarado
    
                        For iCont = 0 To Produto.ListCount - 1
                            If Produto.List(iCont) = sProdutoMascarado Then
                                Produto.ListIndex = iCont
                                Exit For
                            End If
                        Next
                    
                        
                    
                        lErro = PreencheLinha_GridItens(objItemPC)
                        If lErro <> SUCESSO Then gError 65791
                
                    'Se esse item não tem quantidade a Receber
                    Else
                        'O Próximo reinicia a procura
                        iIndice = 0
                
                    End If
                
                End If
        
                Call InsereOrdenadaPC(objItemPCInfo)
                
                'distribuicao - busca a distribuição dos produtos oriunda do pedido de compra
                lErro = CF("LocalizacaoItemPC_Le", objItemPC)
                If lErro <> SUCESSO And lErro <> 56361 Then gError 89614
                
                Set objItemPCInfo.colLocalizacao = objItemPC.colLocalizacao
                
            Next
        
            'Limpa o GridItensPC
            Call Grid_Limpa(objGridItensPC)
            
            'Preenche o GridItensPC
            For iIndice = 1 To gcolItemPedCompraInfo.Count
                lErro = PreencheLinha_ItensPC(gcolItemPedCompraInfo(iIndice))
                If lErro <> SUCESSO Then gError 65792
            Next
            
        End If
                                                             
        
    'Se o pedido foi desselecionado
    Else
                
        'Guarda o Pedido de compras que foi desselecionado
        Set objPedidoCompra = gcolPedidoCompra.Item(Item + 1)
        
        'Atualiza o Grid de NF e o Grid de Requisição de compras
        lErro = Atualiza_GridItens(objPedidoCompra, Item)
        If lErro <> SUCESSO Then gError 65801
        
        'Remove da combo
        For iIndice = 1 To ComboPedidoCompras.ListCount - 1
            If ComboPedidoCompras.List(iIndice) = objPedidoCompra.lCodigo Then
                ComboPedidoCompras.RemoveItem (iIndice)
                Exit For
            End If
        Next

        Set objPedidoCompra.colItens = New Collection

    End If

    Call Total_Calcula
       
    lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
    If lErro <> SUCESSO Then gError 89628
       
    If ComboPedidoCompras.ListIndex = 0 Then ComboPedidoCompras_Click
    
    ComboPedidoCompras.ListIndex = 0
       
    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_PedidosCompra_ItemCheck:

    Select Case gErr

        Case 65786, 65787, 65788, 65790, 65791, 65792, 65798, 65801, 66121, 66566, 89179, 89183, 89614, 89615, 89628

        Case 65789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 66000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_PRECOUNITARIO_DIFERENTE", gErr, objItemPC.sProduto)
            PedidosCompra.Selected(Item) = False
        
        Case 66001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_ALIQUOTAICMS_DIFERENTE", gErr, objItemPC.sProduto)
            PedidosCompra.Selected(Item) = False
            
        Case 66002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_ALIQUOTAIPI_DIFERENTE", gErr, objItemPC.sProduto)
            PedidosCompra.Selected(Item) = False
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166465)

    End Select

    Exit Sub

End Sub

Function Atualiza_Recebimento(objPedidoCompra As ClassPedidoCompras, iItem As Integer) As Long
'Remove do Grid de recebimento os produtos de Pedidos de compras desselecionados da lista

Dim objItemPC As ClassItemPedCompra
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim sProdutoMascarado As String
Dim dQuantAtual As Double
Dim lErro As Long
Dim iPossuiRelacionamento As Integer
Dim iIndice As Integer
Dim iCont As Integer

On Error GoTo Erro_Atualiza_Recebimento

    'Retira do Grid de Pedido de Compras os itens do pedido de compras desselecionado
    For iCont = objGridItensPC.iLinhasExistentes To 1 Step -1
        If PedidosCompra.List(iItem) = GridItensPC.TextMatrix(iCont, iGrid_PedCompra_Col) Then
            
            'Procura pelo mesmo Produto no GridItens
            For iIndice = 1 To objGrid.iLinhasExistentes
                If GridItensPC.TextMatrix(iCont, iGrid_Prod_Col) = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) Then
                    
                    'Atualiza a quantidade de GridItens retirando a quantidade do produtoPC removido
                    GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) - StrParaDbl(GridItensPC.TextMatrix(iCont, iGrid_Recebido_Col)))
                    Exit For
                End If
            Next
                        
            'Exclui a linha de GridItensPC
            lErro = Grid_Exclui_Linha(objGridItensPC, iCont)
            If lErro <> SUCESSO Then gError 54669
            
        End If
    Next
    
    'Retira da coleção os itens do Pedido de compras desselecionado
    For iCont = gcolItemPedCompraInfo.Count To 1 Step -1
        If gcolItemPedCompraInfo.Item(iCont).lPedCompra = CLng(PedidosCompra.List(iItem)) Then
            gcolItemPedCompraInfo.Remove (iCont)
        End If
    Next

    'Para cada Item do Pedido de Compras
    For Each objItemPC In objPedidoCompra.colItens
    
        iPossuiRelacionamento = 0
        
        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 54584
        
        Prod.PromptInclude = False
        Prod.Text = sProdutoMascarado
        Prod.PromptInclude = True

        sProdutoMascarado = Prod.Text

        'Para cada Item de Pedido de Compras que foi selecionado
        For Each objItemPCInfo In gcolItemPedCompraInfo
                
            'Se o Item de Pedido de Compras faz parte de outro pedido selecionado
            If objItemPCInfo.sProduto = objItemPC.sProduto Then
                                          
'                'Guarda a nova quantidade recebida desse produto
'                For iIndice = 1 To objGridItensPC.iLinhasExistentes
'                    If sProdutoMascarado = GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) Then
'                        dQuantAtual = dQuantAtual + StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
'                    End If
'                Next
'
'                'Atualiza o Grid de Itens
'                For iIndice = 1 To objGrid.iLinhasExistentes
'                    If sProdutoMascarado = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) Then
'                        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(dQuantAtual)
'                        Exit For
'                    End If
'                Next
            
                iPossuiRelacionamento = 1
                Exit For
                
            End If
        
        Next
    
        'Se o Item do Pedido que foi desselecionado não está presente em nenhuma linha de GridItensPC
        If iPossuiRelacionamento = 0 Then
            
            'Exclui as linhas onde esse Produto aparece
            For iIndice = objGrid.iLinhasExistentes To 1 Step -1
                If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado Then
                    lErro = Grid_Exclui_Linha(objGrid, iIndice)
                    If lErro <> SUCESSO Then Error 54606
                    
                    'Recalcula os valores
                    lErro = Calcula_Valores(iIndice)
                    If lErro <> SUCESSO Then gError 66633
                
                End If
            Next
        End If
        
    Next

    Atualiza_Recebimento = SUCESSO

    Exit Function

Erro_Atualiza_Recebimento:

    Atualiza_Recebimento = gErr
    
    Select Case gErr

        Case 54584, 54606, 54669, 66633

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166466)

    End Select

    Exit Function

End Function

Function Atualiza_ItensPC() As Long
'Atualiza as quantidades das coleções globais

Dim iIndice As Integer
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_Atualiza_ItensPC
    
    'Atualiza a quantidade recebida dos itens de pedido de Compras
    For Each objItemPCInfo In gcolItemPedCompraInfo
        
        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemPCInfo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 66992
        
        Prod.PromptInclude = False
        Prod.Text = sProdutoMascarado
        Prod.PromptInclude = True

        sProdutoMascarado = Prod.Text
        
        'Procura a quantidade Recebida do ItemPC
        For iIndice = 1 To objGrid.iLinhasExistentes
        
            'Se encontrou
            If sProdutoMascarado = GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) And CStr(objItemPCInfo.lPedCompra) = GridItensPC.TextMatrix(iIndice, iGrid_PedCompra_Col) Then
                objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                Exit For
            End If
        
        Next
        
    Next
    
    Atualiza_ItensPC = SUCESSO
    
    Exit Function
    
Erro_Atualiza_ItensPC:

    Atualiza_ItensPC = gErr
    
    Select Case gErr
            
        Case 66992
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166467)
    
    End Select
    
    Exit Function
    
End Function

Function PreencheLinha_GridItens(objItemPC As ClassItemPedCompra) As Long
'Preenche uma linha no Grid de Recebimento com os dados passados em objItemPC

Dim iAlmoxarifadoPadrao As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long
Dim dValorTotal As Double
Dim dTaxa As Double

On Error GoTo Erro_PreencheLinha_GridItens

    'Descrição do produto
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Descricao_Col) = objItemPC.sDescProduto

    'Unidade de Medida
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_UnidadeMed_Col) = objItemPC.sUM
   
    'Quantidade
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Quantidade_Col) = Formata_Estoque(objItemPC.dQuantidade - objItemPC.dQuantRecebida)

    'Preço unitário
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_ValorUnitario_Col) = Format(objItemPC.dPrecoUnitario * objItemPC.dTaxa, FORMATO_PRECO_UNITARIO_EXTERNO)
    
    'calcula o percentual e valor de desconto
    If objItemPC.dValorDesconto <> 0 Then

        dValorTotal = objItemPC.dPrecoUnitario * (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * objItemPC.dTaxa

        GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Desconto_Col) = Format(objItemPC.dValorDesconto * objItemPC.dTaxa, "Standard")

        If dValorTotal <> 0 Then
            
            GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_PercDesc_Col) = Format(objItemPC.dValorDesconto * objItemPC.dTaxa / dValorTotal, "Percent")
        
        End If

    End If
    objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    
    lErro = Calcula_Valores(objGridItens.iLinhasExistentes)
    If lErro <> SUCESSO Then gError 54700

    PreencheLinha_GridItens = SUCESSO

    Exit Function

Erro_PreencheLinha_GridItens:

    PreencheLinha_GridItens = gErr

    Select Case gErr

        Case 54698, 54699, 54700

        Case 61800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166468)

    End Select

    Exit Function

End Function

Function PreencheLinha_ItensPC(objItemPCInfo As ClassItemPedCompraInfo) As Long
'Preenche Linha Corrente do Grid de Pedido de Compras

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim objInfoPC As ClassItemPedCompraInfo

On Error GoTo Erro_PreencheLinha_ItensPC
        
    'Se o produto possui quantidade à Receber
    If objItemPCInfo.dQuantReceber > 0 Then
    
        'Pedido de Compra
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_PedCompra_Col) = CStr(objItemPCInfo.lPedCompra)
    
        'Item
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Item_Col) = CStr(objItemPCInfo.iItem)
    
        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemPCInfo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 54605
        
        Prod.PromptInclude = False
        Prod.Text = sProdutoMascarado
        Prod.PromptInclude = True

        sProdutoMascarado = Prod.Text
    
        'Produto
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Prod_Col) = sProdutoMascarado
    
        'Descrição do Produto
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_DescProduto_Col) = objItemPCInfo.sDescProduto
    
        'Unidade de Medida
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_UM_Col) = objItemPCInfo.sUM
    
        'Quantidade a receber
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_AReceber_Col) = Formata_Estoque(objItemPCInfo.dQuantReceber)
    
        'Quantidade recebida
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Recebido_Col) = Formata_Estoque(objItemPCInfo.dQuantRecebida)
        
        If ComboPedidoCompras.Text = "TODOS" Then
            For iIndice = 0 To Moeda.ListCount - 1
                If Moeda.ItemData(iIndice) = objItemPCInfo.iMoeda Then
                    GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Moeda_Col) = Moeda.List(iIndice)
                End If
            Next
            
            GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Taxa_Col) = Format(objItemPCInfo.dTaxa, "STANDARD")
            
        End If

    
        'ALTERAÇÃO DE LINHAS EXISTENTES
        objGridItensPC.iLinhasExistentes = objGridItensPC.iLinhasExistentes + 1
        
    End If
    
    PreencheLinha_ItensPC = SUCESSO

    Exit Function

Erro_PreencheLinha_ItensPC:

    PreencheLinha_ItensPC = gErr

    Select Case gErr

        Case 54605

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166469)

    End Select

    Exit Function

End Function

Private Sub Destinatario_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub


Private Sub VolumeQuant_GotFocus()
        
    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Recebimento de Material de Fornecedor - c/ Pedidos de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RecebMaterialFCom"

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
   
   ' Parent.UnloadDoFilho
   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Sub NFiscalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalLabel, Source, X, Y)
End Sub

Private Sub NFiscalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalLabel, Button, Shift, X, Y)
End Sub

Private Sub SerieLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieLabel, Source, X, Y)
End Sub

Private Sub SerieLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieLabel, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
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

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub SubTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SubTotal, Source, X, Y)
End Sub

Private Sub SubTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SubTotal, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Total_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Total, Source, X, Y)
End Sub

Private Sub Total_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Total, Button, Shift, X, Y)
End Sub

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property

'Tratamento dos Grids
Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer, lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iLinha As Integer
Dim sProduto As String

On Error GoTo Erro_GridItens_KeyDown

    iLinhasExistentesAnterior = objGrid.iLinhasExistentes
    iItemAtual = GridItens.Row
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)


    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then
        
        Call SubTotal_Calcula
        Call Total_Calcula

        'Remove do Grid de Itens Pedido de Compras os Itens que possuem o mesmo Produto
        For iIndice = objGridItensPC.iLinhasExistentes To 1 Step -1
            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto Then
                Call Grid_Exclui_Linha(objGridItensPC, iIndice)
            End If
        Next
                
        'Formata o Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 65999
        
        'Remove da coleção de Itens Pedidos de Compras
        For iIndice = gcolItemPedCompraInfo.Count To 1 Step -1
            
            If gcolItemPedCompraInfo.Item(iIndice).sProduto = sProdutoFormatado Then
                gcolItemPedCompraInfo.Remove (iIndice)
            End If
        
        Next
                
        'Verifica se há Pedidos de Compras selecionados que não possuem itens no GridPC
        For iIndice = 0 To PedidosCompra.ListCount - 1
            If PedidosCompra.Selected(iIndice) = True Then
                For iLinha = 1 To objGridItensPC.iLinhasExistentes
                    If GridItensPC.TextMatrix(iLinha, iGrid_PedCompra_Col) = PedidosCompra.List(iIndice) Then
                        Exit For
                    End If
                Next
                If iLinha > objGridItensPC.iLinhasExistentes Then
                    PedidosCompra.Selected(iIndice) = False
                End If
            End If
        Next
        
        'distribuicao
        lErro = gobjDistribuicao.Exclusao_Item_GridDist(iItemAtual)
        If lErro <> SUCESSO Then gError 89657
        
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 54683, 89657

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166470)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Sub GridDist_Click()
'distribuicao
    
    Call gobjDistribuicao.GridDist_Click

End Sub

Public Sub GridDist_EnterCell()
'distribuicao
    
    Call gobjDistribuicao.GridDist_EnterCell

End Sub

Public Sub GridDist_GotFocus()
'distribuicao
    
    Call gobjDistribuicao.GridDist_GotFocus

End Sub

Public Sub GridDist_KeyPress(KeyAscii As Integer)
'distribuicao
    
    Call gobjDistribuicao.GridDist_KeyPress(KeyAscii)

End Sub

Public Sub GridDist_LeaveCell()
'distribuicao
    
    Call gobjDistribuicao.GridDist_LeaveCell

End Sub

Public Sub GridDist_Validate(Cancel As Boolean)
'distribuicao
    
    Call gobjDistribuicao.GridDist_Validate(Cancel)
    
End Sub

Public Sub GridDist_RowColChange()
'distribuicao
    
    Call gobjDistribuicao.GridDist_RowColChange

End Sub

Public Sub GridDist_KeyDown(KeyCode As Integer, Shift As Integer)
'distribuicao
    
    Call gobjDistribuicao.GridDist_KeyDown(KeyCode, Shift)
    
End Sub

Public Sub GridDist_Scroll()
'distribuicao
    
    Call gobjDistribuicao.GridDist_Scroll


End Sub

Private Sub Produto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ValorUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ValorUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ValorUnitario
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PercentDesc_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub


Private Sub PercentDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridItensPC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensPC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensPC, iAlterado)
    End If

End Sub

Private Sub GridItensPC_EnterCell()

    Call Grid_Entrada_Celula(objGridItensPC, iAlterado)

End Sub

Private Sub GridItensPC_GotFocus()

    Call Grid_Recebe_Foco(objGridItensPC)

End Sub

Private Sub GridItensPC_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItensPC)

End Sub

Private Sub GridItensPC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensPC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensPC, iAlterado)
    End If

End Sub

Private Sub GridItensPC_LeaveCell()

    Call Saida_Celula(objGridItensPC)

End Sub

Private Sub GridItensPC_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensPC)

End Sub

Private Sub GridItensPC_RowColChange()

    Call Grid_RowColChange(objGridItensPC)

End Sub

Private Sub GridItensPC_Scroll()

    Call Grid_Scroll(objGridItensPC)

End Sub

Private Sub QuantRecebida_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantRecebida_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensPC)

End Sub

Private Sub QuantRecebida_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensPC)

End Sub

Private Sub QuantRecebida_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensPC.objControle = QuantRecebida
    lErro = Grid_Campo_Libera_Foco(objGridItensPC)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Emitente_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
'distribuicao
        ElseIf Me.ActiveControl Is gobjDistribuicao.AlmoxDist Then
            Call gobjDistribuicao.BotaoLocalizacaoDist_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        End If
    End If

End Sub


Private Sub NumRecebimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumRecebimento, Source, X, Y)
End Sub

Private Sub NumRecebimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumRecebimento, Button, Shift, X, Y)
End Sub

Private Sub LabelRecebimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRecebimento, Source, X, Y)
End Sub

Private Sub LabelRecebimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRecebimento, Button, Shift, X, Y)
End Sub

Private Sub LabelIPIValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelIPIValor, Source, X, Y)
End Sub

Private Sub LabelIPIValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelIPIValor, Button, Shift, X, Y)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166471)
    
    End Select

End Function

Private Sub Taxa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim dTaxa As Double

On Error GoTo Erro_Taxa_Validate

    'Verifica se foi preenchido
    If Len(Trim(Taxa.Text)) > 0 Then

        'Criica se é Valor não negativo
        lErro = Valor_NaoNegativo_Critica(Taxa.Text)
        If lErro <> SUCESSO Then gError 108993

        Taxa.Text = Format(Taxa.Text, "STANDARD")
        
        gbLimpaTaxa = False
        
        'Preenche o grid
        lErro = Preenche_GridItensPC
        If lErro <> SUCESSO Then gError 108994

    End If

    Exit Sub

Erro_Taxa_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 108993, 108994

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166472)

    End Select

End Sub

Private Function Confere_Quantidade_PrecoTotal() As Long
'Verifica se o preco total confere com o do Item NF

Dim iIndice As Integer
Dim iIndice1 As Integer
Dim lErro As Long
Dim dFator As Double
Dim sProduto As String
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim iPreenchido As Integer
Dim dValorItensPC As Double
Dim dQuantidadeItensPC As Double

On Error GoTo Erro_Confere_Quantidade_PrecoTotal
    
    'Descobre o Indice do produto alterado no grid de itens
    For iIndice1 = 1 To objGridItens.iLinhasExistentes
        
        sProduto = GridItens.TextMatrix(iIndice1, iGrid_Produto_Col)
        dValorItensPC = 0
        dQuantidadeItensPC = 0
        
        For iIndice = 1 To objGridItensPC.iLinhasExistentes
    
            If sProduto = GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) Then
            
                lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
                If lErro <> SUCESSO Then gError 114502
        
                objProduto.sCodigo = sProdutoFormatado
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 114503
        
                If lErro = 28030 Then gError 114504
        
                'Converte a UM de GridItensPC para a UM do GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItensPC.TextMatrix(iIndice, iGrid_UM_Col), GridItens.TextMatrix(iIndice1, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 114505
        
                dValorItensPC = dValorItensPC + StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_RS_Col)))
                dQuantidadeItensPC = dQuantidadeItensPC + StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) * dFator
            
            End If
    
        Next
        
        'Se o valor total for diferente do valor dos itens => Erro
        If Abs(StrParaDbl(GridItens.TextMatrix(iIndice1, iGrid_ValorTotal_Col)) - dValorItensPC) > DELTA_VALORMONETARIO Then gError 114506
        If Abs(StrParaDbl(GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col)) - dQuantidadeItensPC) > DELTA_VALORMONETARIO Then gError 114507
            
    Next
    
    Confere_Quantidade_PrecoTotal = SUCESSO

    Exit Function

Erro_Confere_Quantidade_PrecoTotal:

    Confere_Quantidade_PrecoTotal = gErr

    Select Case gErr

        Case 114502, 114503, 114505
        
        Case 114504
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 114506
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_DIFERENTE_VALORITENSPC", gErr, iIndice1)
            
        Case 114507
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_DIFERENTE_QUANTIDADEITENSPC", gErr, iIndice1)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166473)

    End Select

End Function

Private Function CalculaPrecoUnitario_GridItens() As Long
'Calcula o preco unitário baseado no grid de ItensPC

Dim lErro As Long
Dim dValorItensPC As Double
Dim dQuantidadeItensPC As Double
Dim dDescontoPC As Double
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim bAchou As Boolean
Dim objItemPedCompraInfo As ClassItemPedCompraInfo

On Error GoTo Erro_CalculaPrecoUnitario_GridItens

    'Descobre o Indice do produto alterado no grid de itens
    For iIndice1 = 1 To objGridItens.iLinhasExistentes
        
        sProduto = GridItens.TextMatrix(iIndice1, iGrid_Produto_Col)
        dValorItensPC = 0
        dQuantidadeItensPC = 0
        
        'Para cada Item da coleção de ItensPC
        For Each objItemPedCompraInfo In gcolItemPedCompraInfo
    
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
            If lErro <> SUCESSO Then gError 114507
        
            If sProdutoFormatado = objItemPedCompraInfo.sProduto Then
            
                bAchou = True
            
                objProduto.sCodigo = sProdutoFormatado
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 114508
        
                If lErro = 28030 Then gError 114509
        
                'Converte a UM de GridItensPC para a UM do GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPedCompraInfo.sUM, GridItens.TextMatrix(iIndice1, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 114510
        
                If objItemPedCompraInfo.dTaxa > 0 Then
                    dValorItensPC = dValorItensPC + (objItemPedCompraInfo.dQuantRecebida * objItemPedCompraInfo.dPrecoUnitario - objItemPedCompraInfo.dValorDesconto) * objItemPedCompraInfo.dTaxa
                Else
                    dValorItensPC = dValorItensPC + objItemPedCompraInfo.dQuantRecebida * objItemPedCompraInfo.dPrecoUnitario - objItemPedCompraInfo.dValorDesconto
                End If
                
                dQuantidadeItensPC = dQuantidadeItensPC + objItemPedCompraInfo.dQuantRecebida * dFator
            
            End If
    
        Next
        
        If bAchou Then
            GridItens.TextMatrix(iIndice1, iGrid_ValorTotal_Col) = Format(dValorItensPC - (dValorItensPC * PercentParaDbl(GridItens.TextMatrix(iIndice1, iGrid_PercDesc_Col))), "STANDARD")
            GridItens.TextMatrix(iIndice1, iGrid_ValorUnitario_Col) = Format(dValorItensPC / dQuantidadeItensPC, IIf(gobjCOM.sFormatoPrecoUnitario <> "", gobjCOM.sFormatoPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO))
            GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col) = Format(dQuantidadeItensPC, "STANDARD")
            
        End If
        
        Call SubTotal_Calcula
        
        bAchou = False
        
    Next

    CalculaPrecoUnitario_GridItens = SUCESSO
    
    Exit Function
    
Erro_CalculaPrecoUnitario_GridItens:

    CalculaPrecoUnitario_GridItens = gErr
    
    Select Case gErr
    
        Case 114507, 114508, 114510
        
        Case 114509
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166474)
    
    End Select

End Function



Function Atualiza_GridItens(objPedidoCompra As ClassPedidoCompras, iItem As Integer) As Long
'Remove do Grids os produtos de Pedidos de compras desselecionados da lista

Dim objItemPC As ClassItemPedCompra
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim iCont As Integer
Dim sProdutoMascarado As String
Dim dQuantAtual As Double
Dim lErro As Long
Dim lReqCompra As Long
Dim iPossuiRelacionamento As Integer
Dim iIndice As Integer
Dim iLinha As Integer
Dim colPedCompras As New Collection
Dim iIndice2 As Integer

On Error GoTo Erro_Atualiza_GridItens
            
    'Retira do Grid de Pedido de Compras os itens do pedido de compras desselecionado
    For iCont = objGridItensPC.iLinhasExistentes To 1 Step -1
        If PedidosCompra.List(iItem) = GridItensPC.TextMatrix(iCont, iGrid_PedCompra_Col) Then
            lErro = Grid_Exclui_Linha(objGridItensPC, iCont)
            If lErro <> SUCESSO Then gError 65800
        End If
    Next

    'Retira da coleção os itens do Pedido de compras desselecionado
    For iCont = gcolItemPedCompraInfo.Count To 1 Step -1
        If gcolItemPedCompraInfo.Item(iCont).lPedCompra = CLng(PedidosCompra.List(iItem)) Then
            gcolItemPedCompraInfo.Remove (iCont)
        End If
    Next
    
    'Para cada Item do Pedido de Compras
    For Each objItemPC In objPedidoCompra.colItens
    
        iPossuiRelacionamento = 0
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)
        lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 65802

        Prod.PromptInclude = False
        Prod.Text = sProdutoMascarado
        Prod.PromptInclude = True

        sProdutoMascarado = Prod.Text

        dQuantAtual = 0

        'Para cada Item de Pedido de Compras que foi selecionado
        For Each objItemPCInfo In gcolItemPedCompraInfo

            'Se o Item de Pedido de Compras faz parte de outro pedido selecionado
            If objItemPCInfo.sProduto = objItemPC.sProduto Then

                dQuantAtual = 0

                'Guarda a nova quantidade recebida desse produto
                For iIndice = 1 To objGridItensPC.iLinhasExistentes
                    If sProdutoMascarado = GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) Then
                        dQuantAtual = dQuantAtual + StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    End If
                Next

                'Atualiza o Grid de Itens
                For iIndice = 1 To objGridItens.iLinhasExistentes
                    If sProdutoMascarado = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) Then
                        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(dQuantAtual)

                        lErro = Calcula_Valores(iIndice)
                        If lErro <> SUCESSO Then gError 89181

                        Exit For

                    End If
                Next

                iPossuiRelacionamento = 1
                Exit For

            End If

        Next
    
        'Se o Item do Pedido que foi desselecionado não está presente em nenhuma outra linha de GridItensPC
        If iPossuiRelacionamento = 0 Then
            
            'Exclui as linhas onde esse Produto aparece
            For iIndice = objGridItens.iLinhasExistentes To 1 Step -1
                If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado Then
                    lErro = Grid_Exclui_Linha(objGridItens, iIndice)
                    If lErro <> SUCESSO Then gError 66078
                
                End If
            Next
        End If
        
    Next
                
    Atualiza_GridItens = SUCESSO

    Exit Function

Erro_Atualiza_GridItens:

    Atualiza_GridItens = gErr
    
    Select Case gErr

        Case 65800, 65802, 66078, 66079, 66115, 66129, 89181

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166475)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134059

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134059

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166476)

    End Select
    
    Exit Sub

End Sub

