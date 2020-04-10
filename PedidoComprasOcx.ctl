VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PedidoComprasOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   8175
      Index           =   2
      Left            =   150
      TabIndex        =   57
      Top             =   840
      Visible         =   0   'False
      Width           =   16665
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
         Left            =   135
         TabIndex        =   28
         Top             =   7770
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Itens"
         Height          =   6690
         Left            =   120
         TabIndex        =   129
         Top             =   15
         Width           =   16365
         Begin MSMask.MaskEdBox TempoTransito 
            Height          =   225
            Left            =   3675
            TabIndex        =   163
            Top             =   3795
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DeliveryDate 
            Height          =   225
            Left            =   2325
            TabIndex        =   162
            Top             =   3840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TotalMoedaReal 
            Height          =   228
            Left            =   6444
            TabIndex        =   139
            Top             =   1584
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
         Begin MSMask.MaskEdBox PrecoUnitarioMoedaReal 
            Height          =   228
            Left            =   5220
            TabIndex        =   138
            Top             =   1584
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin VB.TextBox DescCompleta 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1140
            MaxLength       =   50
            TabIndex        =   137
            Top             =   960
            Width           =   5460
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   7245
            TabIndex        =   65
            Top             =   720
            Width           =   1035
            _ExtentX        =   1826
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
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1875
            MaxLength       =   50
            TabIndex        =   60
            Top             =   285
            Width           =   4000
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3390
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   270
            Width           =   750
         End
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   5655
            TabIndex        =   63
            Top             =   330
            Width           =   1080
            _ExtentX        =   1905
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   4545
            TabIndex        =   62
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   600
            TabIndex        =   59
            Top             =   240
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   6855
            TabIndex        =   64
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin VB.ComboBox RecebForaFaixa 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoComprasOcx.ctx":0000
            Left            =   435
            List            =   "PedidoComprasOcx.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   2025
            Width           =   2235
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2790
            MaxLength       =   255
            TabIndex        =   67
            Top             =   2085
            Width           =   2445
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   135
            TabIndex        =   68
            Top             =   2475
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
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   225
            Left            =   1455
            TabIndex        =   69
            Top             =   2475
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaICMS 
            Height          =   225
            Left            =   4695
            TabIndex        =   72
            Top             =   2535
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox ValorIPIItem 
            Height          =   225
            Left            =   3615
            TabIndex        =   71
            Top             =   2430
            Width           =   960
            _ExtentX        =   1693
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
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   225
            Left            =   2625
            TabIndex        =   70
            Top             =   2550
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox PercentMaisReceb 
            Height          =   255
            Left            =   5985
            TabIndex        =   73
            Top             =   2430
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   450
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   6225
            Left            =   75
            TabIndex        =   58
            Top             =   285
            Width           =   16080
            _ExtentX        =   28363
            _ExtentY        =   10980
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
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
         Height          =   345
         Left            =   2040
         TabIndex        =   29
         Top             =   7770
         Width           =   1350
      End
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   840
         Index           =   1
         Left            =   120
         TabIndex        =   119
         Top             =   6795
         Width           =   16365
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   2580
            TabIndex        =   23
            Top             =   450
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   4965
            TabIndex        =   24
            Top             =   450
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox OutrasDespesas 
            Height          =   285
            Left            =   7350
            TabIndex        =   25
            Top             =   450
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   9735
            TabIndex        =   26
            Top             =   450
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox ValorIPI 
            Height          =   285
            Left            =   12120
            TabIndex        =   27
            Top             =   450
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   503
            _Version        =   393216
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
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   14505
            TabIndex        =   128
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   14505
            TabIndex        =   127
            Top             =   450
            Width           =   1740
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
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
            Height          =   165
            Index           =   1
            Left            =   12120
            TabIndex        =   126
            Top             =   270
            Width           =   1740
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
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
            Left            =   4965
            TabIndex        =   125
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   9735
            TabIndex        =   124
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
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
            Left            =   2580
            TabIndex        =   123
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Left            =   7350
            TabIndex        =   122
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   195
            TabIndex        =   121
            Top             =   450
            Width           =   1740
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   195
            TabIndex        =   120
            Top             =   240
            Width           =   1740
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8250
      Index           =   5
      Left            =   135
      TabIndex        =   85
      Top             =   810
      Visible         =   0   'False
      Width           =   16665
      Begin VB.CommandButton BotaoLiberaBloqueio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   45
         Picture         =   "PedidoComprasOcx.ctx":003F
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   7530
         Width           =   1905
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   7320
         Left            =   75
         TabIndex        =   117
         Top             =   105
         Width           =   16425
         Begin VB.ComboBox TipoBloqueio 
            Height          =   315
            ItemData        =   "PedidoComprasOcx.ctx":2639
            Left            =   1170
            List            =   "PedidoComprasOcx.ctx":263B
            TabIndex        =   87
            Top             =   3570
            Width           =   3000
         End
         Begin MSMask.MaskEdBox CodUsuario 
            Height          =   270
            Left            =   4245
            TabIndex        =   89
            Top             =   4320
            Width           =   2500
            _ExtentX        =   4419
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ResponsavelBL 
            Height          =   270
            Left            =   6345
            TabIndex        =   90
            Top             =   4365
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataBloqueio 
            Height          =   270
            Left            =   2535
            TabIndex        =   88
            Top             =   465
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ResponsavelLib 
            Height          =   270
            Left            =   8880
            TabIndex        =   92
            Top             =   4920
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLiberacao 
            Height          =   270
            Left            =   6240
            TabIndex        =   91
            Top             =   465
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridBloqueios 
            Height          =   2775
            Left            =   105
            TabIndex        =   86
            Top             =   390
            Width           =   16125
            _ExtentX        =   28443
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8190
      Index           =   4
      Left            =   135
      TabIndex        =   77
      Top             =   840
      Visible         =   0   'False
      Width           =   16620
      Begin VB.CommandButton BotaoCcl 
         Caption         =   "Centros de Custo/Lucro"
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
         Left            =   2175
         TabIndex        =   36
         Top             =   7710
         Width           =   2535
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
         Height          =   345
         Left            =   105
         TabIndex        =   35
         Top             =   7710
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         Caption         =   "Distribuição dos Produtos"
         Height          =   7350
         Left            =   135
         TabIndex        =   118
         Top             =   195
         Width           =   16440
         Begin VB.TextBox UnidMed 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   4860
            TabIndex        =   130
            Top             =   510
            Width           =   1365
         End
         Begin VB.TextBox DescProd 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   80
            Top             =   855
            Width           =   4000
         End
         Begin VB.ComboBox Prod 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   285
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   330
            Width           =   1400
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   225
            Left            =   7350
            TabIndex        =   84
            Top             =   390
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quant 
            Height          =   225
            Left            =   6075
            TabIndex        =   83
            Top             =   210
            Width           =   1335
            _ExtentX        =   2355
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
            Left            =   3015
            TabIndex        =   81
            Top             =   1200
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   4125
            TabIndex        =   82
            Top             =   210
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDistribuicao 
            Height          =   6690
            Left            =   255
            TabIndex        =   78
            Top             =   480
            Width           =   16050
            _ExtentX        =   28310
            _ExtentY        =   11800
            _Version        =   393216
            Rows            =   10
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoContasContabeis 
         Caption         =   "Plano de Contas "
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
         Left            =   4965
         TabIndex        =   37
         Top             =   7710
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8190
      Index           =   3
      Left            =   135
      TabIndex        =   74
      Top             =   840
      Visible         =   0   'False
      Width           =   16665
      Begin VB.Frame Frame6 
         Caption         =   "Frete"
         Height          =   870
         Left            =   105
         TabIndex        =   113
         Top             =   3150
         Width           =   8505
         Begin VB.ComboBox Transportadora 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5445
            TabIndex        =   34
            Top             =   315
            Width           =   2190
         End
         Begin VB.ComboBox TipoFrete 
            Height          =   315
            ItemData        =   "PedidoComprasOcx.ctx":263D
            Left            =   1245
            List            =   "PedidoComprasOcx.ctx":2647
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   330
            Width           =   2550
         End
         Begin VB.Label TransportadoraLabel 
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
            Height          =   210
            Left            =   4050
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   115
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Frete:"
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
            Left            =   270
            TabIndex        =   114
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Local de Entrega"
         Height          =   2865
         Left            =   120
         TabIndex        =   94
         Top             =   225
         Width           =   8505
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   1035
            Index           =   0
            Left            =   4680
            TabIndex        =   95
            Top             =   240
            Width           =   3645
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               ItemData        =   "PedidoComprasOcx.ctx":2655
               Left            =   960
               List            =   "PedidoComprasOcx.ctx":2657
               TabIndex        =   32
               Text            =   "FilialEmpresa"
               Top             =   360
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
               Index           =   1
               Left            =   360
               TabIndex        =   96
               Top             =   420
               Width           =   465
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Height          =   795
            Index           =   1
            Left            =   4800
            TabIndex        =   98
            Top             =   240
            Visible         =   0   'False
            Width           =   3645
            Begin VB.ComboBox FilialFornec 
               Height          =   315
               Left            =   1215
               TabIndex        =   76
               Top             =   480
               Width           =   2160
            End
            Begin MSMask.MaskEdBox Fornec 
               Height          =   300
               Left            =   1215
               TabIndex        =   75
               Top             =   120
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label FornLabel 
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
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   100
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label Label40 
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
               TabIndex        =   99
               Top             =   540
               Width           =   465
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Tipo"
            Height          =   555
            Left            =   270
            TabIndex        =   97
            Top             =   450
            Width           =   3900
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
               Left            =   2130
               TabIndex        =   31
               Top             =   240
               Width           =   1335
            End
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
               Left            =   330
               TabIndex        =   30
               Top             =   225
               Value           =   -1  'True
               Width           =   1515
            End
         End
         Begin VB.Label Pais 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   112
            Top             =   2355
            Width           =   1995
         End
         Begin VB.Label Estado 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   111
            Top             =   2370
            Width           =   495
         End
         Begin VB.Label CEP 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6675
            TabIndex        =   110
            Top             =   1920
            Width           =   930
         End
         Begin VB.Label Cidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   109
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Bairro 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   108
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Endereco 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   107
            Top             =   1500
            Width           =   6345
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "País:"
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
            Left            =   3465
            TabIndex        =   106
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Left            =   6150
            TabIndex        =   105
            Top             =   1995
            Width           =   465
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            TabIndex        =   104
            Top             =   1995
            Width           =   585
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   510
            TabIndex        =   103
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   3285
            TabIndex        =   102
            Top             =   1995
            Width           =   675
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            TabIndex        =   101
            Top             =   1515
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8160
      Index           =   1
      Left            =   165
      TabIndex        =   43
      Top             =   855
      Width           =   16605
      Begin VB.Frame Frame14 
         Caption         =   "Outros"
         Height          =   2505
         Left            =   165
         TabIndex        =   150
         Top             =   5040
         Width           =   11055
         Begin VB.TextBox ObservacaoPC 
            Height          =   1140
            Left            =   1440
            MaxLength       =   255
            TabIndex        =   13
            Top             =   1245
            Width           =   7695
         End
         Begin VB.ComboBox ObsEmbalagem 
            Height          =   315
            Left            =   5910
            TabIndex        =   11
            Top             =   180
            Width           =   3255
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   315
            Left            =   1425
            TabIndex        =   12
            Top             =   765
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoPV 
            Height          =   300
            Left            =   1425
            TabIndex        =   10
            Top             =   255
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label ObsLabel 
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   161
            Top             =   1275
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
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
            Left            =   645
            TabIndex        =   160
            Top             =   825
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
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
            Left            =   4815
            TabIndex        =   159
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Codigo PV:"
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
            Left            =   420
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   158
            Top             =   285
            Width           =   960
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Preços"
         Height          =   1455
         Left            =   165
         TabIndex        =   149
         Top             =   3450
         Width           =   11070
         Begin VB.ComboBox CondPagto 
            Height          =   315
            Left            =   1395
            TabIndex        =   6
            Top             =   315
            Width           =   3210
         End
         Begin VB.ComboBox Moeda 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   900
            Width           =   1935
         End
         Begin VB.CommandButton BotaoTrazCotacao 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7470
            Style           =   1  'Graphical
            TabIndex        =   153
            ToolTipText     =   "Numeração Automática"
            Top             =   840
            Width           =   345
         End
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   5940
            TabIndex        =   7
            Top             =   330
            Width           =   3090
         End
         Begin MSMask.MaskEdBox Taxa 
            Height          =   315
            Left            =   5925
            TabIndex        =   9
            Top             =   855
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            Format          =   "###,##0.00##"
            PromptChar      =   " "
         End
         Begin VB.Label CondPagtoLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   255
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   157
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label Label5 
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
            Height          =   195
            Left            =   5355
            TabIndex        =   156
            Top             =   930
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Moeda:"
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
            Left            =   675
            TabIndex        =   155
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tabela Preço:"
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
            Index           =   62
            Left            =   4710
            TabIndex        =   154
            Top             =   390
            Width           =   1215
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Dados do Fornecedor"
         Height          =   750
         Left            =   165
         TabIndex        =   148
         Top             =   1215
         Width           =   11100
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5940
            TabIndex        =   2
            Top             =   270
            Width           =   3165
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1410
            TabIndex        =   1
            Top             =   285
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   152
            Top             =   330
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   151
            Top             =   330
            Width           =   1035
         End
      End
      Begin VB.CommandButton BotaoPedidosAvulsos 
         BackColor       =   &H80000007&
         Caption         =   "Pedidos Avulsos"
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
         Left            =   90
         TabIndex        =   14
         Top             =   7740
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificação"
         Height          =   945
         Left            =   165
         TabIndex        =   44
         Top             =   150
         Width           =   11085
         Begin VB.Frame Frame11 
            Caption         =   "Projeto - Invisível"
            Height          =   450
            Left            =   8175
            TabIndex        =   142
            Top             =   2310
            Visible         =   0   'False
            Width           =   5265
            Begin VB.CommandButton BotaoProjetos 
               Caption         =   "..."
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
               Left            =   2715
               TabIndex        =   144
               Top             =   120
               Width           =   495
            End
            Begin VB.ComboBox Etapa 
               Height          =   315
               Left            =   4635
               Style           =   2  'Dropdown List
               TabIndex        =   143
               Top             =   135
               Width           =   2550
            End
            Begin MSMask.MaskEdBox Projeto 
               Height          =   300
               Left            =   840
               TabIndex        =   145
               Top             =   150
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Etapa:"
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
               Index           =   53
               Left            =   4020
               TabIndex        =   147
               Top             =   195
               Width           =   570
            End
            Begin VB.Label LabelProjeto 
               AutoSize        =   -1  'True
               Caption         =   "Projeto:"
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
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   146
               Top             =   330
               Width           =   675
            End
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2340
            Picture         =   "PedidoComprasOcx.ctx":2659
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Numeração Automática"
            Top             =   375
            Width           =   300
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1440
            TabIndex        =   0
            Top             =   360
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Embalagem 
            Height          =   315
            Left            =   -10000
            TabIndex        =   41
            Top             =   2190
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label LabelEmbalagem 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
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
            Left            =   -10000
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   136
            Top             =   2250
            Width           =   1035
         End
         Begin VB.Label Comprador 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5940
            TabIndex        =   47
            Top             =   375
            Width           =   3120
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
            Left            =   4920
            TabIndex        =   46
            Top             =   435
            Width           =   975
         End
         Begin VB.Label CodigoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nº Pedido:"
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
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   45
            Top             =   420
            Width           =   930
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Datas"
         Height          =   1275
         Left            =   165
         TabIndex        =   48
         Top             =   2100
         Width           =   11085
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   2490
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1410
            TabIndex        =   3
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
         Begin MSComCtl2.UpDown UpDownDataEnvio 
            Height          =   300
            Left            =   7050
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   765
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvio 
            Height          =   300
            Left            =   5955
            TabIndex        =   5
            Top             =   765
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataRefFluxo 
            Height          =   300
            Left            =   10680
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataRefFluxo 
            Height          =   300
            Left            =   9585
            TabIndex        =   4
            ToolTipText     =   "Data de Referência usada em Fluxo de Caixa"
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fluxo:"
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
            Left            =   8955
            TabIndex        =   141
            ToolTipText     =   "Data de Referência usada em Fluxo de Caixa"
            Top             =   330
            Width           =   525
         End
         Begin VB.Label DataAlteracao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1410
            TabIndex        =   54
            Top             =   765
            Width           =   1095
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Alterado:"
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
            TabIndex        =   53
            Top             =   810
            Width           =   780
         End
         Begin VB.Label DataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5940
            TabIndex        =   52
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Emitido:"
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
            Left            =   5145
            TabIndex        =   51
            Top             =   375
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Enviado:"
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
            Left            =   5070
            TabIndex        =   55
            Top             =   825
            Width           =   765
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
            Left            =   840
            TabIndex        =   49
            Top             =   375
            Width           =   480
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8160
      Index           =   6
      Left            =   135
      TabIndex        =   131
      Top             =   870
      Visible         =   0   'False
      Width           =   16665
      Begin VB.CommandButton BotaoIncluirNota 
         Caption         =   "Incluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   900
         TabIndex        =   40
         Top             =   1860
         Width           =   1650
      End
      Begin VB.TextBox Nota 
         Height          =   1380
         Left            =   930
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   375
         Width           =   14910
      End
      Begin VB.Frame Frame10 
         Caption         =   "Notas"
         Height          =   5325
         Left            =   105
         TabIndex        =   132
         Top             =   2775
         Width           =   16185
         Begin VB.TextBox NotaPC 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1260
            MaxLength       =   150
            TabIndex        =   135
            Top             =   570
            Width           =   14220
         End
         Begin MSFlexGridLib.MSFlexGrid GridNotas 
            Height          =   1455
            Left            =   210
            TabIndex        =   133
            Top             =   255
            Width           =   15645
            _ExtentX        =   27596
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Label LabelNota 
         Caption         =   "Nota :"
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   134
         Top             =   390
         Width           =   675
      End
   End
   Begin VB.CommandButton BotaoInfoAdic 
      Caption         =   "Informações Adicionais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   11355
      TabIndex        =   16
      Top             =   45
      Width           =   2280
   End
   Begin VB.CheckBox ImprimePedido 
      Caption         =   "Imprimir o pedido ao gravar"
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
      Left            =   8175
      TabIndex        =   15
      Top             =   225
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   13665
      ScaleHeight     =   495
      ScaleWidth      =   3075
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   45
      Width           =   3135
      Begin VB.CommandButton BotaoEmail 
         Height          =   360
         Left            =   90
         Picture         =   "PedidoComprasOcx.ctx":2743
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Enviar email"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1575
         Picture         =   "PedidoComprasOcx.ctx":30E5
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   570
         Picture         =   "PedidoComprasOcx.ctx":326F
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2580
         Picture         =   "PedidoComprasOcx.ctx":3371
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   2085
         Picture         =   "PedidoComprasOcx.ctx":34EF
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   1080
         Picture         =   "PedidoComprasOcx.ctx":3A21
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8775
      Left            =   90
      TabIndex        =   42
      Top             =   345
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   15478
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Entrega"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bloqueios"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notas"
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
Attribute VB_Name = "PedidoComprasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTPedidoCompras
Attribute objCT.VB_VarHelpID = -1


Private Sub CodigoPV_Validate(Cancel As Boolean)
    Call objCT.CodigoPV_Validate(Cancel)
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTPedidoCompras
    Set objCT.objUserControl = Me
End Sub

Private Sub Almoxarifado_Change()
     Call objCT.Almoxarifado_Change
End Sub

Private Sub Almoxarifado_GotFocus()
     Call objCT.Almoxarifado_GotFocus
End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)
     Call objCT.Almoxarifado_KeyPress(KeyAscii)
End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)
     Call objCT.Almoxarifado_Validate(Cancel)
End Sub

Private Sub BotaoAlmoxarifados_Click()
     Call objCT.BotaoAlmoxarifados_Click
End Sub

Private Sub BotaoCcl_Click()
     Call objCT.BotaoCcl_Click
End Sub

Private Sub BotaoContasContabeis_Click()
     Call objCT.BotaoContasContabeis_Click
End Sub

Private Sub BotaoEmail_Click()
     Call objCT.BotaoEmail_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoImprimir_Click()
     Call objCT.BotaoImprimir_Click
End Sub

Private Sub BotaoIncluirNota_Click()
     Call objCT.BotaoIncluirNota_Click
End Sub

Private Sub BotaoLiberaBloqueio_Click()
     Call objCT.BotaoLiberaBloqueio_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoPedidosAvulsos_Click()
     Call objCT.BotaoPedidosAvulsos_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoTrazCotacao_Click()
     Call objCT.BotaoTrazCotacao_Click
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub CodigoLabel_Click()
     Call objCT.CodigoLabel_Click
End Sub

Private Sub CondPagto_Validate(Cancel As Boolean)
     Call objCT.CondPagto_Validate(Cancel)
End Sub

Private Sub CondPagtoLabel_Click()
     Call objCT.CondPagtoLabel_Click
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub Data_Validate(Cancel As Boolean)
     Call objCT.Data_Validate(Cancel)
End Sub

Private Sub DataEnvio_GotFocus()
     Call objCT.DataEnvio_GotFocus
End Sub

Private Sub DataEnvio_Validate(Cancel As Boolean)
     Call objCT.DataEnvio_Validate(Cancel)
End Sub

Private Sub DescProduto_Change()
     Call objCT.DescProduto_Change
End Sub

Private Sub DescProduto_GotFocus()
     Call objCT.DescProduto_GotFocus
End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)
     Call objCT.DescProduto_KeyPress(KeyAscii)
End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)
     Call objCT.DescProduto_Validate(Cancel)
End Sub

Private Sub Embalagem_Change()
     Call objCT.Embalagem_Change
End Sub

Private Sub Embalagem_Validate(Cancel As Boolean)
     Call objCT.Embalagem_Validate(Cancel)
End Sub

Private Sub FilialFornec_Click()
     Call objCT.FilialFornec_Click
End Sub

Private Sub GridDistribuicao_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDistribuicao_KeyDown(KeyCode, Shift)
End Sub

Private Sub LabelEmbalagem_Click()
     Call objCT.LabelEmbalagem_Click
End Sub

Private Sub LabelNota_Click()
     Call objCT.LabelNota_Click
End Sub

Private Sub Moeda_Change()
     Call objCT.Moeda_Change
End Sub

Private Sub Moeda_Click()
     Call objCT.Moeda_Click
End Sub

Private Sub Nota_Change()
     Call objCT.Nota_Change
End Sub

Private Sub Taxa_Change()
     Call objCT.Taxa_Change
End Sub

Private Sub Taxa_Validate(Cancel As Boolean)
     Call objCT.Taxa_Validate(Cancel)
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub UnidadeMed_Change()
     Call objCT.UnidadeMed_Change
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub UnidadeMed_GotFocus()
     Call objCT.UnidadeMed_GotFocus
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMed_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)
     Call objCT.UnidadeMed_Validate(Cancel)
End Sub

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub Quantidade_GotFocus()
     Call objCT.Quantidade_GotFocus
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
     Call objCT.Quantidade_KeyPress(KeyAscii)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub PrecoUnitario_Change()
     Call objCT.PrecoUnitario_Change
End Sub

Private Sub PrecoUnitario_GotFocus()
     Call objCT.PrecoUnitario_GotFocus
End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoUnitario_KeyPress(KeyAscii)
End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)
     Call objCT.PrecoUnitario_Validate(Cancel)
End Sub

Private Sub PercentDesc_Change()
     Call objCT.PercentDesc_Change
End Sub

Private Sub PercentDesc_GotFocus()
     Call objCT.PercentDesc_GotFocus
End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)
     Call objCT.PercentDesc_KeyPress(KeyAscii)
End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)
     Call objCT.PercentDesc_Validate(Cancel)
End Sub

Private Sub Desconto_Change()
     Call objCT.Desconto_Change
End Sub

Private Sub Desconto_GotFocus()
     Call objCT.Desconto_GotFocus
End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto_KeyPress(KeyAscii)
End Sub

Private Sub Desconto_Validate(Cancel As Boolean)
     Call objCT.Desconto_Validate(Cancel)
End Sub

Private Sub DataLimite_Change()
     Call objCT.DataLimite_Change
End Sub

Private Sub DataLimite_GotFocus()
     Call objCT.DataLimite_GotFocus
End Sub

Private Sub DataLimite_KeyPress(KeyAscii As Integer)
     Call objCT.DataLimite_KeyPress(KeyAscii)
End Sub

Private Sub DataLimite_Validate(Cancel As Boolean)
     Call objCT.DataLimite_Validate(Cancel)
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub FilialEmpresa_Click()
     Call objCT.FilialEmpresa_Click
End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)
     Call objCT.FilialEmpresa_Validate(Cancel)
End Sub

Private Sub FilialFornec_Validate(Cancel As Boolean)
     Call objCT.FilialFornec_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Fornec_Change()
     Call objCT.Fornec_Change
End Sub

Private Sub Fornec_Validate(Cancel As Boolean)
     Call objCT.Fornec_Validate(Cancel)
End Sub

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub FornecedorLabel_Click()
     Call objCT.FornecedorLabel_Click
End Sub

Private Sub FornLabel_Click()
     Call objCT.FornLabel_Click
End Sub

Private Sub GridBloqueios_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridBloqueios_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItens_KeyDown(KeyCode, Shift)
End Sub

Private Sub ObsLabel_Click()
     Call objCT.ObsLabel_Click
End Sub

Private Sub OutrasDespesas_Validate(Cancel As Boolean)
     Call objCT.OutrasDespesas_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub TipoDestino_Click(Index As Integer)
     Call objCT.TipoDestino_Click(Index)
End Sub

Private Sub TipoFrete_Click()
     Call objCT.TipoFrete_Click
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub UpDownData_DownClick()
     Call objCT.UpDownData_DownClick
End Sub

Private Sub UpDownData_UpClick()
     Call objCT.UpDownData_UpClick
End Sub

Private Sub UpDownDataEnvio_DownClick()
     Call objCT.UpDownDataEnvio_DownClick
End Sub

Private Sub UpDownDataEnvio_UpClick()
     Call objCT.UpDownDataEnvio_UpClick
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)
     Call objCT.ValorFrete_Validate(Cancel)
End Sub

Private Sub ValorIPI_Validate(Cancel As Boolean)
     Call objCT.ValorIPI_Validate(Cancel)
End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)
     Call objCT.ValorSeguro_Validate(Cancel)
End Sub

Public Function Trata_Parametros(Optional objPedidoCompra As ClassPedidoCompras) As Long
     Trata_Parametros = objCT.Trata_Parametros(objPedidoCompra)
End Function

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_GotFocus()
     Call objCT.Produto_GotFocus
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
     Call objCT.Produto_KeyPress(KeyAscii)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub GridItens_Click()
     Call objCT.GridItens_Click
End Sub

Private Sub GridItens_EnterCell()
     Call objCT.GridItens_EnterCell
End Sub

Private Sub GridItens_GotFocus()
     Call objCT.GridItens_GotFocus
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)
     Call objCT.GridItens_KeyPress(KeyAscii)
End Sub

Private Sub GridItens_LeaveCell()
     Call objCT.GridItens_LeaveCell
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
     Call objCT.GridItens_Validate(Cancel)
End Sub

Private Sub GridItens_RowColChange()
     Call objCT.GridItens_RowColChange
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub AliquotaIPI_Change()
     Call objCT.AliquotaIPI_Change
End Sub

Private Sub AliquotaIPI_GotFocus()
     Call objCT.AliquotaIPI_GotFocus
End Sub

Private Sub AliquotaIPI_KeyPress(KeyAscii As Integer)
     Call objCT.AliquotaIPI_KeyPress(KeyAscii)
End Sub

Private Sub AliquotaIPI_Validate(Cancel As Boolean)
     Call objCT.AliquotaIPI_Validate(Cancel)
End Sub

Private Sub GridBloqueios_Click()
     Call objCT.GridBloqueios_Click
End Sub

Private Sub GridBloqueios_EnterCell()
     Call objCT.GridBloqueios_EnterCell
End Sub

Private Sub GridBloqueios_GotFocus()
     Call objCT.GridBloqueios_GotFocus
End Sub

Private Sub GridBloqueios_KeyPress(KeyAscii As Integer)
     Call objCT.GridBloqueios_KeyPress(KeyAscii)
End Sub

Private Sub GridBloqueios_LeaveCell()
     Call objCT.GridBloqueios_LeaveCell
End Sub

Private Sub GridBloqueios_Validate(Cancel As Boolean)
     Call objCT.GridBloqueios_Validate(Cancel)
End Sub

Private Sub GridBloqueios_RowColChange()
     Call objCT.GridBloqueios_RowColChange
End Sub

Private Sub GridBloqueios_Scroll()
     Call objCT.GridBloqueios_Scroll
End Sub

Private Sub GridDistribuicao_Click()
     Call objCT.GridDistribuicao_Click
End Sub

Private Sub GridDistribuicao_EnterCell()
     Call objCT.GridDistribuicao_EnterCell
End Sub

Private Sub GridDistribuicao_GotFocus()
     Call objCT.GridDistribuicao_GotFocus
End Sub

Private Sub GridDistribuicao_KeyPress(KeyAscii As Integer)
     Call objCT.GridDistribuicao_KeyPress(KeyAscii)
End Sub

Private Sub GridDistribuicao_LeaveCell()
     Call objCT.GridDistribuicao_LeaveCell
End Sub

Private Sub GridDistribuicao_Validate(Cancel As Boolean)
     Call objCT.GridDistribuicao_Validate(Cancel)
End Sub

Private Sub GridDistribuicao_RowColChange()
     Call objCT.GridDistribuicao_RowColChange
End Sub

Private Sub GridDistribuicao_Scroll()
     Call objCT.GridDistribuicao_Scroll
End Sub

Private Sub ValorIPIItem_Change()
     Call objCT.ValorIPIItem_Change
End Sub

Private Sub ValorIPIItem_GotFocus()
     Call objCT.ValorIPIItem_GotFocus
End Sub

Private Sub ValorIPIItem_KeyPress(KeyAscii As Integer)
     Call objCT.ValorIPIItem_KeyPress(KeyAscii)
End Sub

Private Sub ValorIPIItem_Validate(Cancel As Boolean)
     Call objCT.ValorIPIItem_Validate(Cancel)
End Sub

Private Sub AliquotaICMS_Change()
     Call objCT.AliquotaICMS_Change
End Sub

Private Sub AliquotaICMS_GotFocus()
     Call objCT.AliquotaICMS_GotFocus
End Sub

Private Sub AliquotaICMS_KeyPress(KeyAscii As Integer)
     Call objCT.AliquotaICMS_KeyPress(KeyAscii)
End Sub

Private Sub AliquotaICMS_Validate(Cancel As Boolean)
     Call objCT.AliquotaICMS_Validate(Cancel)
End Sub

Private Sub PercentMaisReceb_Change()
     Call objCT.PercentMaisReceb_Change
End Sub

Private Sub PercentMaisReceb_GotFocus()
     Call objCT.PercentMaisReceb_GotFocus
End Sub

Private Sub PercentMaisReceb_KeyPress(KeyAscii As Integer)
     Call objCT.PercentMaisReceb_KeyPress(KeyAscii)
End Sub

Private Sub PercentMaisReceb_Validate(Cancel As Boolean)
     Call objCT.PercentMaisReceb_Validate(Cancel)
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub Observacao_GotFocus()
     Call objCT.Observacao_GotFocus
End Sub

Private Sub Observacao_KeyPress(KeyAscii As Integer)
     Call objCT.Observacao_KeyPress(KeyAscii)
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)
     Call objCT.Observacao_Validate(Cancel)
End Sub

Private Sub RecebForaFaixa_Change()
     Call objCT.RecebForaFaixa_Change
End Sub

Private Sub RecebForaFaixa_Click()
     Call objCT.RecebForaFaixa_Click
End Sub

Private Sub RecebForaFaixa_GotFocus()
     Call objCT.RecebForaFaixa_GotFocus
End Sub

Private Sub RecebForaFaixa_KeyPress(KeyAscii As Integer)
     Call objCT.RecebForaFaixa_KeyPress(KeyAscii)
End Sub

Private Sub RecebForaFaixa_Validate(Cancel As Boolean)
     Call objCT.RecebForaFaixa_Validate(Cancel)
End Sub

Private Sub Prod_Change()
     Call objCT.Prod_Change
End Sub

Private Sub Prod_Click()
     Call objCT.Prod_Click
End Sub

Private Sub Prod_GotFocus()
     Call objCT.Prod_GotFocus
End Sub

Private Sub Prod_KeyPress(KeyAscii As Integer)
     Call objCT.Prod_KeyPress(KeyAscii)
End Sub

Private Sub Prod_Validate(Cancel As Boolean)
     Call objCT.Prod_Validate(Cancel)
End Sub

Private Sub DescProd_Change()
     Call objCT.DescProd_Change
End Sub

Private Sub DescProd_Click()
     Call objCT.DescProd_Click
End Sub

Private Sub DescProd_GotFocus()
     Call objCT.DescProd_GotFocus
End Sub

Private Sub DescProd_KeyPress(KeyAscii As Integer)
     Call objCT.DescProd_KeyPress(KeyAscii)
End Sub

Private Sub DescProd_Validate(Cancel As Boolean)
     Call objCT.DescProd_Validate(Cancel)
End Sub

Private Sub CentroCusto_Change()
     Call objCT.CentroCusto_Change
End Sub

Private Sub CentroCusto_Click()
     Call objCT.CentroCusto_Click
End Sub

Private Sub CentroCusto_GotFocus()
     Call objCT.CentroCusto_GotFocus
End Sub

Private Sub CentroCusto_KeyPress(KeyAscii As Integer)
     Call objCT.CentroCusto_KeyPress(KeyAscii)
End Sub

Private Sub CentroCusto_Validate(Cancel As Boolean)
     Call objCT.CentroCusto_Validate(Cancel)
End Sub

Private Sub UnidMed_Change()
     Call objCT.UnidMed_Change
End Sub

Private Sub UnidMed_Click()
     Call objCT.UnidMed_Click
End Sub

Private Sub UnidMed_GotFocus()
     Call objCT.UnidMed_GotFocus
End Sub

Private Sub UnidMed_KeyPress(KeyAscii As Integer)
     Call objCT.UnidMed_KeyPress(KeyAscii)
End Sub

Private Sub UnidMed_Validate(Cancel As Boolean)
     Call objCT.UnidMed_Validate(Cancel)
End Sub

Private Sub Quant_Change()
     Call objCT.Quant_Change
End Sub

Private Sub Quant_GotFocus()
     Call objCT.Quant_GotFocus
End Sub

Private Sub Quant_KeyPress(KeyAscii As Integer)
     Call objCT.Quant_KeyPress(KeyAscii)
End Sub

Private Sub Quant_Validate(Cancel As Boolean)
     Call objCT.Quant_Validate(Cancel)
End Sub

Private Sub ContaContabil_Change()
     Call objCT.ContaContabil_Change
End Sub

Private Sub ContaContabil_Click()
     Call objCT.ContaContabil_Click
End Sub

Private Sub ContaContabil_GotFocus()
     Call objCT.ContaContabil_GotFocus
End Sub

Private Sub ContaContabil_KeyPress(KeyAscii As Integer)
     Call objCT.ContaContabil_KeyPress(KeyAscii)
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)
     Call objCT.ContaContabil_Validate(Cancel)
End Sub

Private Sub TipoBloqueio_Change()
     Call objCT.TipoBloqueio_Change
End Sub

Private Sub TipoBloqueio_Click()
     Call objCT.TipoBloqueio_Click
End Sub

Private Sub TipoBloqueio_GotFocus()
     Call objCT.TipoBloqueio_GotFocus
End Sub

Private Sub TipoBloqueio_KeyPress(KeyAscii As Integer)
     Call objCT.TipoBloqueio_KeyPress(KeyAscii)
End Sub

Private Sub TipoBloqueio_Validate(Cancel As Boolean)
     Call objCT.TipoBloqueio_Validate(Cancel)
End Sub

Private Sub ResponsavelBL_Change()
     Call objCT.ResponsavelBL_Change
End Sub

Private Sub ResponsavelBL_GotFocus()
     Call objCT.ResponsavelBL_GotFocus
End Sub

Private Sub ResponsavelBL_KeyPress(KeyAscii As Integer)
     Call objCT.ResponsavelBL_KeyPress(KeyAscii)
End Sub

Private Sub ResponsavelBL_Validate(Cancel As Boolean)
     Call objCT.ResponsavelBL_Validate(Cancel)
End Sub

Private Sub Label6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label6(Index), Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6(Index), Button, Shift, X, Y)
End Sub
Private Sub ObsLabel_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(ObsLabel, Source, X, Y)
End Sub
Private Sub ObsLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ObsLabel, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
End Sub
Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub
Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub
Private Sub DataAlteracao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataAlteracao, Source, X, Y)
End Sub
Private Sub DataAlteracao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataAlteracao, Button, Shift, X, Y)
End Sub
Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub
Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub
Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub
Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub
Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub
Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub
Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub
Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub
Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub
Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub
Private Sub Comprador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Comprador, Source, X, Y)
End Sub
Private Sub Comprador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Comprador, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub
Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub
Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub
Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
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
Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub
Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub
Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub
Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub
Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub
Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub
Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub
Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub
Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub
Private Sub Label41_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label41, Source, X, Y)
End Sub
Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label41, Button, Shift, X, Y)
End Sub
Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub
Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub
Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub
Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub
Private Sub FornLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornLabel, Source, X, Y)
End Sub
Private Sub FornLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornLabel, Button, Shift, X, Y)
End Sub
Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub
Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub
Private Sub Pais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Pais, Source, X, Y)
End Sub
Private Sub Pais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Pais, Button, Shift, X, Y)
End Sub
Private Sub Estado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Estado, Source, X, Y)
End Sub
Private Sub Estado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Estado, Button, Shift, X, Y)
End Sub
Private Sub CEP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CEP, Source, X, Y)
End Sub
Private Sub CEP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CEP, Button, Shift, X, Y)
End Sub
Private Sub Cidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cidade, Source, X, Y)
End Sub
Private Sub Cidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cidade, Button, Shift, X, Y)
End Sub
Private Sub Bairro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Bairro, Source, X, Y)
End Sub
Private Sub Bairro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Bairro, Button, Shift, X, Y)
End Sub
Private Sub Endereco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Endereco, Source, X, Y)
End Sub
Private Sub Endereco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Endereco, Button, Shift, X, Y)
End Sub
Private Sub Label63_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label63, Source, X, Y)
End Sub
Private Sub Label63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label63, Button, Shift, X, Y)
End Sub
Private Sub Label65_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label65, Source, X, Y)
End Sub
Private Sub Label65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label65, Button, Shift, X, Y)
End Sub
Private Sub Label70_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label70, Source, X, Y)
End Sub
Private Sub Label70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label70, Button, Shift, X, Y)
End Sub
Private Sub Label71_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label71, Source, X, Y)
End Sub
Private Sub Label71_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label71, Button, Shift, X, Y)
End Sub
Private Sub Label72_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label72, Source, X, Y)
End Sub
Private Sub Label72_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label72, Button, Shift, X, Y)
End Sub
Private Sub Label73_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label73, Source, X, Y)
End Sub
Private Sub Label73_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label73, Button, Shift, X, Y)
End Sub
Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub
Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub
Private Sub GridNotas_Click()
     Call objCT.GridNotas_Click
End Sub

Private Sub GridNotas_EnterCell()
     Call objCT.GridNotas_EnterCell
End Sub

Private Sub GridNotas_GotFocus()
     Call objCT.GridNotas_GotFocus
End Sub

Private Sub GridNotas_KeyPress(KeyAscii As Integer)
     Call objCT.GridNotas_KeyPress(KeyAscii)
End Sub

Private Sub GridNotas_LeaveCell()
     Call objCT.GridNotas_LeaveCell
End Sub

Private Sub GridNotas_Validate(Cancel As Boolean)
     Call objCT.GridNotas_Validate(Cancel)
End Sub

Private Sub GridNotas_RowColChange()
     Call objCT.GridNotas_RowColChange
End Sub

Private Sub GridNotas_Scroll()
     Call objCT.GridNotas_Scroll
End Sub

Private Sub GridNotas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridNotas_KeyDown(KeyCode, Shift)
End Sub

Private Sub NotaPC_Change()
     Call objCT.NotaPC_Change
End Sub

Private Sub NotaPC_GotFocus()
     Call objCT.NotaPC_GotFocus
End Sub

Private Sub NotaPC_KeyPress(KeyAscii As Integer)
     Call objCT.NotaPC_KeyPress(KeyAscii)
End Sub

Private Sub NotaPC_Validate(Cancel As Boolean)
     Call objCT.NotaPC_Validate(Cancel)
End Sub

Private Sub Fornecedor_Preenche()
     Call objCT.Fornecedor_Preenche
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
             Set objCT.objUserControl = Nothing
             Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub DataRefFluxo_GotFocus()
     Call objCT.DataRefFluxo_GotFocus
End Sub

Private Sub DataRefFluxo_Validate(Cancel As Boolean)
     Call objCT.DataRefFluxo_Validate(Cancel)
End Sub

Private Sub BotaoInfoAdic_Click()
     Call objCT.BotaoInfoAdic_Click
End Sub

Private Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub

Private Sub LabelProjeto_Click()
    Call objCT.LabelProjeto_Click
End Sub

Private Sub Projeto_Change()
     Call objCT.Projeto_Change
End Sub

Private Sub Projeto_GotFocus()
     Call objCT.Projeto_GotFocus
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub

Sub Etapa_Change()
     Call objCT.Projeto_Change
End Sub

Sub Etapa_Click()
     Call objCT.Projeto_Change
End Sub

Sub Etapa_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub BotaoEntrega_Click()
    Call objCT.BotaoEntrega_Click
End Sub

Private Sub DeliveryDate_Change()
     Call objCT.DeliveryDate_Change
End Sub

Private Sub DeliveryDate_GotFocus()
     Call objCT.DeliveryDate_GotFocus
End Sub

Private Sub DeliveryDate_KeyPress(KeyAscii As Integer)
     Call objCT.DeliveryDate_KeyPress(KeyAscii)
End Sub

Private Sub DeliveryDate_Validate(Cancel As Boolean)
     Call objCT.DeliveryDate_Validate(Cancel)
End Sub

Private Sub TempoTransito_Change()
     Call objCT.TempoTransito_Change
End Sub

Private Sub TempoTransito_GotFocus()
     Call objCT.TempoTransito_GotFocus
End Sub

Private Sub TempoTransito_KeyPress(KeyAscii As Integer)
     Call objCT.TempoTransito_KeyPress(KeyAscii)
End Sub

Private Sub TempoTransito_Validate(Cancel As Boolean)
     Call objCT.TempoTransito_Validate(Cancel)
End Sub
