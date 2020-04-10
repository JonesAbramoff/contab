VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrcamentoVendaOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9200
   ScaleMode       =   0  'User
   ScaleWidth      =   17000
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   8280
      Index           =   6
      Left            =   180
      TabIndex        =   155
      Top             =   720
      Visible         =   0   'False
      Width           =   16695
      Begin VB.CommandButton BotaoAtualizarFP 
         Caption         =   "Atualizar"
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
         Left            =   15180
         TabIndex        =   90
         Top             =   7830
         Width           =   1200
      End
      Begin VB.CommandButton BotaoCotacoesRecebidas 
         Caption         =   "Cotações Atualizadas"
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
         Left            =   2355
         TabIndex        =   89
         Top             =   7860
         Width           =   2145
      End
      Begin VB.Frame Frame7 
         Caption         =   "Formação de preço do item selecionado acima"
         Height          =   3945
         Left            =   75
         TabIndex        =   165
         Top             =   3750
         Width           =   16335
         Begin MSMask.MaskEdBox FPQtde 
            Height          =   225
            Left            =   1500
            TabIndex        =   171
            Top             =   1155
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FPPrecoUnit 
            Height          =   225
            Left            =   3135
            TabIndex        =   170
            Top             =   1710
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FPPercentMargem 
            Height          =   225
            Left            =   2025
            TabIndex        =   169
            Top             =   1695
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSMask.MaskEdBox FPCustoUnit 
            Height          =   225
            Left            =   525
            TabIndex        =   168
            Top             =   1695
            Width           =   1440
            _ExtentX        =   2540
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
         Begin VB.ComboBox FPUnidMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":0000
            Left            =   645
            List            =   "OrcamentoVendaOcx.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   172
            Top             =   1110
            Width           =   975
         End
         Begin MSMask.MaskEdBox FPProduto 
            Height          =   225
            Left            =   630
            TabIndex        =   174
            Top             =   585
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox FPDescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1815
            MaxLength       =   250
            TabIndex        =   173
            Top             =   600
            Width           =   4000
         End
         Begin VB.ComboBox FPSituacao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":0004
            Left            =   5880
            List            =   "OrcamentoVendaOcx.ctx":0014
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   1650
            Width           =   1605
         End
         Begin MSMask.MaskEdBox FPPrecoTotal 
            Height          =   225
            Left            =   4590
            TabIndex        =   167
            Top             =   1710
            Width           =   1440
            _ExtentX        =   2540
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
         Begin MSFlexGridLib.MSFlexGrid GridFormacaoPreco 
            Height          =   3450
            Left            =   105
            TabIndex        =   175
            Top             =   210
            Width           =   15960
            _ExtentX        =   28152
            _ExtentY        =   6085
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Preços Calculados"
         Height          =   3450
         Index           =   0
         Left            =   60
         TabIndex        =   156
         Top             =   195
         Width           =   16335
         Begin VB.OptionButton PCSelecionado 
            Caption         =   "Option1"
            Height          =   225
            Left            =   105
            TabIndex        =   177
            Top             =   255
            Width           =   915
         End
         Begin MSMask.MaskEdBox PCPrecoUnitCalc 
            Height          =   225
            Left            =   5550
            TabIndex        =   176
            Top             =   390
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.ComboBox PCSituacao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":0035
            Left            =   6345
            List            =   "OrcamentoVendaOcx.ctx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   1125
            Width           =   1575
         End
         Begin VB.TextBox PCDescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2700
            MaxLength       =   50
            TabIndex        =   162
            Top             =   1440
            Width           =   4000
         End
         Begin MSMask.MaskEdBox PCProduto 
            Height          =   225
            Left            =   750
            TabIndex        =   163
            Top             =   1035
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PCPrecoTotal 
            Height          =   225
            Left            =   6360
            TabIndex        =   161
            Top             =   810
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PCPrecoUnit 
            Height          =   225
            Left            =   4860
            TabIndex        =   160
            Top             =   780
            Width           =   1590
            _ExtentX        =   2805
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
         Begin MSMask.MaskEdBox PCQtde 
            Height          =   225
            Left            =   3135
            TabIndex        =   159
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
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
         Begin VB.ComboBox PCUnidMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":0066
            Left            =   2625
            List            =   "OrcamentoVendaOcx.ctx":0068
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   360
            Width           =   885
         End
         Begin MSFlexGridLib.MSFlexGrid GridPrecosCalculados 
            Height          =   2955
            Left            =   165
            TabIndex        =   157
            Top             =   195
            Width           =   15960
            _ExtentX        =   28152
            _ExtentY        =   5212
            _Version        =   393216
         End
      End
      Begin VB.CommandButton BotaoCotacoesPendentes 
         Caption         =   "Cotações a Atualizar"
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
         Left            =   120
         TabIndex        =   88
         Top             =   7860
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   8310
      Index           =   5
      Left            =   -30
      TabIndex        =   147
      Top             =   690
      Visible         =   0   'False
      Width           =   16890
      Begin TelasFAT.TabTributacaoFat TabTrib 
         Height          =   4590
         Left            =   225
         TabIndex        =   180
         Top             =   45
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   8096
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   8340
      Index           =   4
      Left            =   90
      TabIndex        =   124
      Top             =   690
      Visible         =   0   'False
      Width           =   16800
      Begin VB.CheckBox CobrancaAutomatica 
         Caption         =   "Calcula cobrança automaticamente"
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
         Left            =   300
         TabIndex        =   85
         Top             =   150
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   7860
         Left            =   45
         TabIndex        =   132
         Top             =   435
         Width           =   16530
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3045
            Picture         =   "OrcamentoVendaOcx.ctx":006A
            Style           =   1  'Graphical
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   480
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3045
            Picture         =   "OrcamentoVendaOcx.ctx":00C4
            Style           =   1  'Graphical
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   330
            Width           =   240
         End
         Begin VB.ComboBox TipoDesconto1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   112
            Top             =   1215
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            TabIndex        =   115
            Top             =   1530
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3105
            TabIndex        =   119
            Top             =   1845
            Width           =   1965
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7470
            TabIndex        =   114
            Top             =   1260
            Width           =   900
            _ExtentX        =   1588
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
         Begin MSMask.MaskEdBox Desconto3Valor 
            Height          =   225
            Left            =   6105
            TabIndex        =   121
            Top             =   1905
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   120
            Top             =   1905
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
         Begin MSMask.MaskEdBox Desconto2Valor 
            Height          =   225
            Left            =   6135
            TabIndex        =   117
            Top             =   1590
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   116
            Top             =   1590
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
         Begin MSMask.MaskEdBox Desconto1Valor 
            Height          =   225
            Left            =   6120
            TabIndex        =   125
            Top             =   1260
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   113
            Top             =   1260
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
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   570
            TabIndex        =   110
            Top             =   1230
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
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   240
            Left            =   1695
            TabIndex        =   111
            Top             =   1245
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Percentual 
            Height          =   225
            Left            =   7500
            TabIndex        =   118
            Top             =   1605
            Width           =   900
            _ExtentX        =   1588
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
         Begin MSMask.MaskEdBox Desconto3Percentual 
            Height          =   225
            Left            =   7455
            TabIndex        =   122
            Top             =   1935
            Width           =   900
            _ExtentX        =   1588
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
         Begin MSMask.MaskEdBox DataReferencia 
            Height          =   300
            Left            =   1935
            TabIndex        =   86
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   6855
            Left            =   135
            TabIndex        =   126
            Top             =   855
            Width           =   16200
            _ExtentX        =   28575
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox ValorDescontoTit 
            Height          =   300
            Left            =   7605
            TabIndex        =   87
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total a Receber:"
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
            Left            =   9465
            TabIndex        =   197
            Top             =   375
            Width           =   1455
         End
         Begin VB.Label ValorTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   11010
            TabIndex        =   196
            Top             =   345
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Original:"
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
            Index           =   6
            Left            =   3840
            TabIndex        =   195
            Top             =   375
            Width           =   1215
         End
         Begin VB.Label ValorOriginalTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5145
            TabIndex        =   194
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
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
            Index           =   8
            Left            =   6675
            TabIndex        =   193
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Data de Referência:"
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
            Index           =   9
            Left            =   105
            TabIndex        =   146
            Top             =   375
            Width           =   1740
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Base ICMS Subst"
      Height          =   8325
      Index           =   3
      Left            =   90
      TabIndex        =   213
      Top             =   750
      Visible         =   0   'False
      Width           =   16770
      Begin VB.Frame Frame8 
         Caption         =   "Volumes"
         Height          =   825
         Left            =   180
         TabIndex        =   227
         Top             =   2745
         Width           =   15210
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   11925
            MaxLength       =   20
            TabIndex        =   77
            Top             =   315
            Width           =   1440
         End
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   7725
            TabIndex        =   76
            Top             =   315
            Width           =   2730
         End
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   4095
            TabIndex        =   75
            Top             =   315
            Width           =   2730
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1860
            TabIndex        =   74
            Top             =   315
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   52
            Left            =   7125
            TabIndex        =   231
            Top             =   375
            Width           =   600
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   51
            Left            =   3270
            TabIndex        =   230
            Top             =   375
            Width           =   750
         End
         Begin VB.Label Label1 
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
            Index           =   50
            Left            =   630
            TabIndex        =   229
            Top             =   375
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   32
            Left            =   11085
            TabIndex        =   228
            Top             =   330
            Width           =   705
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dados de Entrega"
         Height          =   2505
         Left            =   180
         TabIndex        =   222
         Top             =   75
         Width           =   15210
         Begin VB.Frame Frame11 
            Caption         =   "Transportadora"
            Height          =   2070
            Left            =   7575
            TabIndex        =   262
            Top             =   195
            Width           =   6525
            Begin VB.Frame Frame1 
               Caption         =   "Redespacho"
               Height          =   1065
               Index           =   13
               Left            =   420
               TabIndex        =   263
               Top             =   765
               Width           =   5730
               Begin VB.ComboBox TranspRedespacho 
                  Height          =   315
                  Left            =   1515
                  TabIndex        =   72
                  Top             =   285
                  Width           =   3825
               End
               Begin VB.CheckBox RedespachoCli 
                  Caption         =   "por conta do cliente"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   225
                  TabIndex        =   73
                  Top             =   705
                  Width           =   2100
               End
               Begin VB.Label TranspRedLabel 
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
                  Left            =   90
                  MousePointer    =   14  'Arrow and Question
                  TabIndex        =   264
                  Top             =   330
                  Width           =   1365
               End
            End
            Begin VB.ComboBox Transportadora 
               Height          =   315
               Left            =   1935
               TabIndex        =   71
               Top             =   375
               Width           =   3825
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
               Left            =   510
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   265
               Top             =   420
               Width           =   1365
            End
         End
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   4710
            TabIndex        =   70
            Top             =   1725
            Width           =   735
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   1845
            MaxLength       =   10
            TabIndex        =   69
            Top             =   1785
            Width           =   1290
         End
         Begin VB.ComboBox FilialEntrega 
            Height          =   315
            Left            =   1830
            TabIndex        =   67
            Top             =   315
            Width           =   3630
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frete por conta do"
            Height          =   870
            Index           =   1
            Left            =   255
            TabIndex        =   223
            Top             =   720
            Width           =   5355
            Begin VB.ComboBox TipoFrete 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   390
               Width           =   5070
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "U.F. da Placa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   53
            Left            =   3405
            TabIndex        =   226
            Top             =   1785
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Placa do Veículo:"
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
            Index           =   54
            Left            =   150
            TabIndex        =   225
            Top             =   1830
            Width           =   1530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial para Entrega:"
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
            Height          =   390
            Index           =   100
            Left            =   75
            TabIndex        =   224
            Top             =   375
            Width           =   1620
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Complemento"
         Height          =   4245
         Index           =   7
         Left            =   195
         TabIndex        =   214
         Top             =   3885
         Width           =   15180
         Begin VB.TextBox Mensagem 
            Height          =   2550
            Left            =   1830
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   78
            Top             =   270
            Width           =   11505
         End
         Begin VB.ComboBox CanalVenda 
            Height          =   315
            Left            =   1860
            TabIndex        =   82
            Top             =   3480
            Width           =   1440
         End
         Begin MSMask.MaskEdBox PedidoCliente 
            Height          =   300
            Left            =   6705
            TabIndex        =   83
            Top             =   3525
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   300
            Left            =   6705
            TabIndex        =   80
            Top             =   3060
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cubagem 
            Height          =   300
            Left            =   11880
            TabIndex        =   81
            Top             =   3090
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PedidoRepr 
            Height          =   300
            Left            =   11880
            TabIndex        =   84
            Top             =   3525
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   300
            Left            =   1860
            TabIndex        =   79
            Top             =   3060
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cubagem:"
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
            Index           =   57
            Left            =   10890
            TabIndex        =   221
            Top             =   3135
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Cliente:"
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
            Index           =   22
            Left            =   5265
            TabIndex        =   220
            Top             =   3570
            Width           =   1305
         End
         Begin VB.Label MensagemLabel 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem:"
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
            TabIndex        =   219
            Top             =   255
            Width           =   975
         End
         Begin VB.Label CanalVendaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Canal de Venda:"
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
            TabIndex        =   218
            Top             =   3555
            Width           =   1425
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   55
            Left            =   5370
            TabIndex        =   217
            Top             =   3105
            Width           =   1215
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   56
            Left            =   690
            TabIndex        =   216
            Top             =   3105
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ped. Representante:"
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
            Index           =   35
            Left            =   9960
            TabIndex        =   215
            Top             =   3570
            Width           =   1770
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   8340
      Index           =   2
      Left            =   60
      TabIndex        =   97
      Top             =   750
      Visible         =   0   'False
      Width           =   16815
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   6345
         Index           =   3
         Left            =   60
         TabIndex        =   130
         Top             =   15
         Width           =   16695
         Begin VB.Frame FrameTS 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   255
            Left            =   13635
            TabIndex        =   242
            Top             =   5970
            Visible         =   0   'False
            Width           =   2940
            Begin VB.Label TS 
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   1725
               TabIndex        =   244
               Top             =   0
               Width           =   1200
            End
            Begin VB.Label LabelTS 
               Caption         =   "Total Selecionado:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   60
               TabIndex        =   243
               Top             =   0
               Width           =   1800
            End
         End
         Begin VB.CheckBox SomaItem 
            Height          =   195
            Left            =   -1000
            TabIndex        =   241
            Top             =   2025
            Width           =   400
         End
         Begin VB.CommandButton BotaoDesce 
            Height          =   195
            Left            =   285
            Picture         =   "OrcamentoVendaOcx.ctx":011E
            Style           =   1  'Graphical
            TabIndex        =   240
            Top             =   285
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.CommandButton BotaoSobe 
            Height          =   195
            Left            =   90
            Picture         =   "OrcamentoVendaOcx.ctx":02E0
            Style           =   1  'Graphical
            TabIndex        =   239
            Top             =   285
            Visible         =   0   'False
            Width           =   180
         End
         Begin MSMask.MaskEdBox PrazoEntregaItem 
            Height          =   225
            Left            =   6885
            TabIndex        =   232
            Top             =   915
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotalB 
            Height          =   225
            Left            =   3705
            TabIndex        =   212
            Top             =   1215
            Width           =   1185
            _ExtentX        =   2090
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
         Begin VB.ComboBox StatusItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   1080
            Width           =   1920
         End
         Begin VB.ComboBox MotivoPerdaItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   2100
            Width           =   1905
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5310
            MaxLength       =   50
            TabIndex        =   152
            Top             =   5355
            Width           =   6000
         End
         Begin MSMask.MaskEdBox VersaoKit 
            Height          =   225
            Left            =   5220
            TabIndex        =   150
            Top             =   1200
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox VersaoKitBase 
            Height          =   225
            Left            =   5970
            TabIndex        =   151
            Top             =   1350
            Width           =   1485
            _ExtentX        =   2619
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   5790
            TabIndex        =   98
            Top             =   2670
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   5190
            MaxLength       =   250
            TabIndex        =   105
            Top             =   3690
            Width           =   4000
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":04A2
            Left            =   1575
            List            =   "OrcamentoVendaOcx.ctx":04A4
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   240
            Width           =   720
         End
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   2640
            TabIndex        =   103
            Top             =   660
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
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   1440
            TabIndex        =   101
            Top             =   585
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   270
            TabIndex        =   99
            Top             =   2565
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   4185
            TabIndex        =   104
            Top             =   360
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2580
            TabIndex        =   102
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
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
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   5670
            TabIndex        =   106
            Top             =   360
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1455
            Left            =   30
            TabIndex        =   123
            Top             =   210
            Width           =   16575
            _ExtentX        =   29236
            _ExtentY        =   2566
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
      Begin VB.CommandButton BotaoEstoqueProd 
         Caption         =   "Estoque - Produto"
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
         Left            =   10542
         TabIndex        =   65
         Top             =   7860
         Width           =   1785
      End
      Begin VB.CommandButton BotaoInfoAdicItem 
         Caption         =   "Informações Adicionais do Item"
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
         Left            =   13560
         TabIndex        =   66
         Top             =   7860
         Width           =   3195
      End
      Begin VB.CommandButton BotaoKitVenda 
         Caption         =   "Kits de Venda"
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
         Left            =   2493
         TabIndex        =   62
         Top             =   7860
         Width           =   1500
      End
      Begin VB.CommandButton BotaoVersaoKitBase 
         Caption         =   "Versão Kit Base"
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
         Left            =   5226
         TabIndex        =   63
         Top             =   7860
         Width           =   1725
      End
      Begin VB.CommandButton BotaoGrade 
         Caption         =   "Grade ..."
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
         Left            =   45
         TabIndex        =   61
         Top             =   7860
         Width           =   1215
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
         Left            =   8184
         TabIndex        =   64
         Top             =   7860
         Width           =   1125
      End
      Begin VB.Frame Frame2 
         Caption         =   "Totais"
         Height          =   1290
         Index           =   4
         Left            =   60
         TabIndex        =   131
         Top             =   6390
         Width           =   16695
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   75
            TabIndex        =   56
            Top             =   915
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   75
            TabIndex        =   107
            Top             =   405
            Visible         =   0   'False
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   4935
            TabIndex        =   58
            Top             =   915
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   2520
            TabIndex        =   57
            Top             =   915
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercDescontoItens 
            Height          =   285
            Left            =   7365
            TabIndex        =   59
            ToolTipText     =   "Percentual de desconto dos itens"
            Top             =   915
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDescontoItens 
            Height          =   285
            Left            =   9795
            TabIndex        =   60
            ToolTipText     =   "Soma dos descontos dos itens"
            Top             =   915
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ISSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9790
            TabIndex        =   261
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Base ISS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   15
            Left            =   9790
            TabIndex        =   260
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label ValorProdutos2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   14640
            TabIndex        =   211
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   21
            Left            =   14670
            TabIndex        =   210
            Top             =   705
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   20
            Left            =   12225
            TabIndex        =   209
            Top             =   705
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   19
            Left            =   4980
            TabIndex        =   208
            Top             =   705
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   17
            Left            =   2580
            TabIndex        =   207
            Top             =   705
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   16
            Left            =   105
            TabIndex        =   206
            Top             =   705
            Width           =   1650
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   12
            Left            =   9825
            TabIndex        =   205
            Top             =   705
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "% Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   11
            Left            =   7425
            TabIndex        =   204
            Top             =   705
            Width           =   1830
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ISS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   10
            Left            =   12215
            TabIndex        =   203
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   7
            Left            =   14685
            TabIndex        =   202
            Top             =   210
            Width           =   1830
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ICMS ST"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   5
            Left            =   7365
            TabIndex        =   201
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "BC ICMS ST"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   4940
            TabIndex        =   200
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ICMS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   2515
            TabIndex        =   199
            Top             =   195
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Base ICMS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   198
            Top             =   195
            Width           =   1905
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   12215
            TabIndex        =   181
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7365
            TabIndex        =   133
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4940
            TabIndex        =   134
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2515
            TabIndex        =   135
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   90
            TabIndex        =   136
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   14640
            TabIndex        =   137
            Top             =   405
            Width           =   1860
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   12210
            TabIndex        =   138
            Top             =   915
            Width           =   1905
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   14640
            TabIndex        =   139
            Top             =   915
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8175
      Index           =   1
      Left            =   120
      TabIndex        =   95
      Top             =   870
      Width           =   16770
      Begin VB.Frame Frame10 
         Caption         =   "Configurações"
         Height          =   690
         Left            =   105
         TabIndex        =   258
         Top             =   6765
         Width           =   16500
         Begin VB.CheckBox ImprimirOVComCodProd 
            Caption         =   "Exibir os código do produtos ao imprimir"
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
            Left            =   11205
            TabIndex        =   44
            Top             =   225
            Visible         =   0   'False
            Width           =   3930
         End
         Begin VB.CheckBox ImprimirOVComPreco 
            Caption         =   "Exibir os valores do orçamento ao imprimir"
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
            Left            =   6375
            TabIndex        =   43
            Top             =   225
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   4500
         End
         Begin VB.CheckBox ImprimeOrcamentoGravacao 
            Caption         =   "Imprimir ao gravar"
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
            Left            =   3645
            TabIndex        =   42
            Top             =   255
            Width           =   2385
         End
         Begin VB.CheckBox EmailOrcamentoGravacao 
            Caption         =   "Enviar email ao gravar"
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
            Left            =   255
            TabIndex        =   41
            Top             =   255
            Width           =   2235
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Outros"
         Height          =   2145
         Left            =   105
         TabIndex        =   245
         Top             =   3675
         Width           =   16500
         Begin VB.ComboBox MotivoPerda 
            Height          =   315
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1185
            Width           =   11265
         End
         Begin VB.TextBox Contato 
            Height          =   330
            Left            =   1425
            TabIndex        =   29
            Top             =   1650
            Width           =   2610
         End
         Begin VB.TextBox Email 
            Height          =   330
            Left            =   5250
            TabIndex        =   30
            Top             =   1635
            Width           =   7425
         End
         Begin VB.ComboBox Status 
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":04A6
            Left            =   1425
            List            =   "OrcamentoVendaOcx.ctx":04A8
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   705
            Width           =   3585
         End
         Begin MSMask.MaskEdBox PrazoValidade 
            Height          =   300
            Left            =   8580
            TabIndex        =   23
            Top             =   720
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   315
            Left            =   1425
            TabIndex        =   18
            Top             =   240
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Vendedor2 
            Height          =   315
            Left            =   8580
            TabIndex        =   19
            Top             =   240
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSComCtl2.UpDown UpDownEnvio 
            Height          =   300
            Left            =   15960
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvio 
            Height          =   300
            Left            =   14910
            TabIndex        =   20
            Top             =   255
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownPerda 
            Height          =   300
            Left            =   15975
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataPerda 
            Height          =   300
            Left            =   14925
            TabIndex        =   27
            Top             =   1200
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownFechamento 
            Height          =   300
            Left            =   15975
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   1665
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFechamento 
            Height          =   300
            Left            =   14925
            TabIndex        =   31
            Top             =   1665
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownProxContato 
            Height          =   300
            Left            =   15975
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   735
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataProxContato 
            Height          =   300
            Left            =   14925
            TabIndex        =   24
            Top             =   735
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label11 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9345
            TabIndex        =   259
            Top             =   750
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Perda:"
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
            Index           =   23
            Left            =   14280
            TabIndex        =   256
            Top             =   1245
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fechamento:"
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
            Index           =   24
            Left            =   13785
            TabIndex        =   255
            Top             =   1710
            Width           =   1110
         End
         Begin VB.Label LabelContato 
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
            Left            =   525
            TabIndex        =   254
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label LabelEmail 
            AutoSize        =   -1  'True
            Caption         =   "E-mail:"
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
            Left            =   4605
            TabIndex        =   253
            Top             =   1695
            Width           =   585
         End
         Begin VB.Label Vendedor2Label 
            AutoSize        =   -1  'True
            Caption         =   "2o. Vendedor:"
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
            Left            =   7200
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   252
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            TabIndex        =   251
            Top             =   765
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Motivo Perda:"
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
            TabIndex        =   250
            Top             =   1230
            Width           =   1200
         End
         Begin VB.Label VendedorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
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
            TabIndex        =   249
            Top             =   285
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Prazo de Validade:"
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
            Left            =   6795
            TabIndex        =   248
            Top             =   765
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Envio:"
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
            Index           =   27
            Left            =   14295
            TabIndex        =   247
            Top             =   285
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Próximo Contato:"
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
            Index           =   28
            Left            =   13410
            TabIndex        =   246
            Top             =   780
            Width           =   2010
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Entrega"
         Height          =   705
         Left            =   105
         TabIndex        =   235
         Top             =   5955
         Width           =   16500
         Begin VB.Frame FrameDataPrazoEnt 
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   345
            Index           =   0
            Left            =   3810
            TabIndex        =   92
            Top             =   240
            Width           =   2025
            Begin MSMask.MaskEdBox DataEntregaPV 
               Height          =   300
               Left            =   570
               TabIndex        =   37
               Top             =   15
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEntregaPV 
               Height          =   300
               Left            =   1695
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   25
               Left            =   60
               TabIndex        =   236
               Top             =   30
               Width           =   480
            End
         End
         Begin VB.Frame FrameDataPrazoEnt 
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   315
            Index           =   2
            Left            =   6585
            TabIndex        =   238
            Top             =   225
            Width           =   9690
            Begin VB.ComboBox PrazoTexto 
               Height          =   315
               ItemData        =   "OrcamentoVendaOcx.ctx":04AA
               Left            =   645
               List            =   "OrcamentoVendaOcx.ctx":04AC
               TabIndex        =   40
               Top             =   15
               Width           =   9015
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Texto:"
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
               Index           =   26
               Left            =   0
               TabIndex        =   93
               Top             =   75
               Width           =   555
            End
         End
         Begin VB.OptionButton OptDataPrazoEnt 
            Height          =   270
            Index           =   2
            Left            =   6345
            TabIndex        =   39
            Top             =   285
            Width           =   240
         End
         Begin VB.Frame FrameDataPrazoEnt 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   330
            Index           =   1
            Left            =   510
            TabIndex        =   237
            Top             =   255
            Width           =   3015
            Begin VB.ComboBox PrazoEntVar 
               Height          =   315
               ItemData        =   "OrcamentoVendaOcx.ctx":04AE
               Left            =   1245
               List            =   "OrcamentoVendaOcx.ctx":04BE
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   0
               Width           =   1440
            End
            Begin MSMask.MaskEdBox PrazoEntrega 
               Height          =   300
               Left            =   540
               TabIndex        =   34
               Top             =   0
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   5
               Mask            =   "#####"
               PromptChar      =   " "
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Prazo:"
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
               Left            =   0
               TabIndex        =   91
               Top             =   45
               Width           =   555
            End
         End
         Begin VB.OptionButton OptDataPrazoEnt 
            Height          =   285
            Index           =   1
            Left            =   255
            TabIndex        =   33
            Top             =   270
            Width           =   240
         End
         Begin VB.OptionButton OptDataPrazoEnt 
            Height          =   270
            Index           =   0
            Left            =   3615
            TabIndex        =   36
            Top             =   270
            Value           =   -1  'True
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Preços"
         Height          =   1110
         Index           =   2
         Left            =   105
         TabIndex        =   129
         Top             =   2400
         Width           =   16500
         Begin VB.ComboBox Idioma 
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":04ED
            Left            =   8610
            List            =   "OrcamentoVendaOcx.ctx":04EF
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   2775
         End
         Begin VB.ComboBox StatusComercial 
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":04F1
            Left            =   13770
            List            =   "OrcamentoVendaOcx.ctx":0504
            TabIndex        =   15
            Top             =   210
            Width           =   2535
         End
         Begin VB.ComboBox Moeda 
            Height          =   315
            ItemData        =   "OrcamentoVendaOcx.ctx":053C
            Left            =   5250
            List            =   "OrcamentoVendaOcx.ctx":053E
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   180
            Width           =   2160
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   1410
            TabIndex        =   12
            Top             =   225
            Width           =   2985
         End
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   1395
            TabIndex        =   16
            Top             =   660
            Width           =   3000
         End
         Begin MSMask.MaskEdBox PercAcrescFin 
            Height          =   315
            Left            =   8610
            TabIndex        =   17
            Top             =   630
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Idioma:"
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
            Left            =   7815
            TabIndex        =   257
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label9 
            Caption         =   "Análise de Preços:"
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
            Left            =   12090
            TabIndex        =   234
            Top             =   255
            Width           =   1620
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
            Height          =   210
            Left            =   4515
            TabIndex        =   233
            Top             =   240
            Width           =   615
         End
         Begin VB.Label CondPagtoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cond.Pagto:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   142
            Top             =   285
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Acréscimo Financeiro:"
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
            Index           =   18
            Left            =   6375
            TabIndex        =   143
            Top             =   690
            Width           =   2070
         End
         Begin VB.Label Label8 
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
            Index           =   2
            Left            =   90
            TabIndex        =   144
            Top             =   720
            Width           =   1200
         End
      End
      Begin VB.CommandButton BotaoOVVersoes 
         Caption         =   "Versões anteriores desse Orçamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   60
         TabIndex        =   45
         Top             =   7635
         Width           =   2235
      End
      Begin VB.CommandButton BotaoVersoesOVs 
         Caption         =   "Versões anteriores de todos Orçamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2490
         TabIndex        =   46
         Top             =   7635
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame FrameVersao 
         Caption         =   "Dados da Versão"
         Height          =   675
         Left            =   105
         TabIndex        =   185
         Top             =   765
         Width           =   16500
         Begin VB.CheckBox TrocarVersao 
            Caption         =   "Trocar versão ao gravar"
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
            Left            =   2490
            TabIndex        =   5
            Top             =   225
            Width           =   3315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Versão:"
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
            TabIndex        =   192
            Top             =   270
            Width           =   645
         End
         Begin VB.Label OVVersao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1425
            TabIndex        =   191
            Top             =   240
            Width           =   870
         End
         Begin VB.Label DataUltAlt 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   13755
            TabIndex        =   190
            Top             =   210
            Width           =   1440
         End
         Begin VB.Label HoraUltAlt 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   15195
            TabIndex        =   189
            Top             =   210
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data\Hora Alteração:"
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
            Index           =   14
            Left            =   11850
            TabIndex        =   188
            Top             =   255
            Width           =   1830
         End
         Begin VB.Label UsuarioLabel 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
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
            Left            =   7725
            TabIndex        =   187
            Top             =   255
            Width           =   705
         End
         Begin VB.Label Usuario 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8610
            TabIndex        =   186
            Top             =   210
            Width           =   2250
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Projeto"
         Height          =   645
         Left            =   9435
         TabIndex        =   182
         Top             =   1635
         Width           =   7170
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   4425
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   225
            Width           =   2550
         End
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
            Left            =   3105
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   990
            TabIndex        =   9
            Top             =   240
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Index           =   41
            Left            =   3795
            TabIndex        =   184
            Top             =   285
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
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   183
            Top             =   285
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   645
         Index           =   6
         Left            =   105
         TabIndex        =   128
         Top             =   1635
         Width           =   8985
         Begin VB.CheckBox CalcularST 
            Caption         =   "Calcular ST"
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
            Left            =   7500
            TabIndex        =   8
            Top             =   240
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5250
            TabIndex        =   7
            Top             =   255
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1410
            TabIndex        =   6
            Top             =   255
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   13
            Left            =   4695
            TabIndex        =   140
            Top             =   300
            Width           =   480
         End
         Begin VB.Label LabelCliente 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   141
            Top             =   300
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   645
         Index           =   0
         Left            =   105
         TabIndex        =   127
         Top             =   -15
         Width           =   16500
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2400
            Picture         =   "OrcamentoVendaOcx.ctx":0540
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   240
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   6285
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   5235
            TabIndex        =   2
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1425
            TabIndex        =   0
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoBase 
            Height          =   300
            Left            =   9795
            TabIndex        =   4
            Top             =   225
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label NatOpInternaEspelho 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   15330
            TabIndex        =   179
            Top             =   195
            Width           =   915
         End
         Begin VB.Label LblNatOpInternaEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Operação (CFOP):"
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
            Left            =   12645
            TabIndex        =   178
            Top             =   255
            Width           =   2640
         End
         Begin VB.Label NumeroBaseLabel 
            AutoSize        =   -1  'True
            Caption         =   "Número do Orçamento  Base:"
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
            Left            =   7155
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   149
            Top             =   270
            Width           =   2490
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
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
            Left            =   4440
            TabIndex        =   145
            Top             =   285
            Width           =   765
         End
         Begin VB.Label NumeroLabel 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   615
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   96
            Top             =   285
            Width           =   720
         End
      End
   End
   Begin VB.CommandButton BotaoAnaliseVenda 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11520
      TabIndex        =   47
      Top             =   90
      Width           =   330
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
      Height          =   450
      Left            =   11880
      TabIndex        =   48
      Top             =   75
      Width           =   1470
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   13395
      ScaleHeight     =   450
      ScaleWidth      =   3480
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   45
      Width           =   3540
      Begin VB.CommandButton BtnExportaDoc 
         Height          =   345
         Left            =   60
         Picture         =   "OrcamentoVendaOcx.ctx":062A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Exporta o Orçamento para um documento do Word"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoEmail 
         Height          =   345
         Left            =   570
         Picture         =   "OrcamentoVendaOcx.ctx":0BFC
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Enviar email"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   3000
         Picture         =   "OrcamentoVendaOcx.ctx":159E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   2520
         Picture         =   "OrcamentoVendaOcx.ctx":171C
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   2040
         Picture         =   "OrcamentoVendaOcx.ctx":1C4E
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   1560
         Picture         =   "OrcamentoVendaOcx.ctx":1DD8
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   1065
         Picture         =   "OrcamentoVendaOcx.ctx":1F32
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   8895
      Left            =   30
      TabIndex        =   94
      Top             =   285
      Width           =   16905
      _ExtentX        =   29819
      _ExtentY        =   15690
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formação de Preços"
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
Attribute VB_Name = "OrcamentoVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTOrcamentoVenda
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoAnaliseVenda_Click()
    Call objCT.BotaoAnaliseVenda_Click
End Sub

Private Sub BotaoCotacoesPendentes_Click()
    Call objCT.BotaoCotacoesPendentes_Click
End Sub

Private Sub BotaoCotacoesRecebidas_Click()
    Call objCT.BotaoCotacoesRecebidas_Click
End Sub

Private Sub CalcularST_Click()
    Call objCT.CalcularST_Click
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
    Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub Status_Change()
    Call objCT.Status_Change
End Sub

Private Sub Status_Click()
    Call objCT.Status_Click
End Sub

Private Sub StatusComercial_Change()
    Call objCT.StatusComercial_Change
End Sub

Private Sub StatusComercial_Click()
    Call objCT.StatusComercial_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTOrcamentoVenda
    Set objCT.objUserControl = Me
End Sub

Function Trata_Parametros(Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long
     Trata_Parametros = objCT.Trata_Parametros(objOrcamentoVenda)
End Function

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub BotaoDataReferenciaDown_Click()
     Call objCT.BotaoDataReferenciaDown_Click
End Sub

Private Sub BotaoDataReferenciaUp_Click()
     Call objCT.BotaoDataReferenciaUp_Click
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

Private Sub BotaoVersaoKitBase_Click()
     Call objCT.BotaoVersaoKitBase_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub CobrancaAutomatica_Click()
     Call objCT.CobrancaAutomatica_Click
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub CodigoBase_Validate(Cancel As Boolean)
     Call objCT.CodigoBase_Validate(Cancel)
End Sub

Private Sub CondicaoPagamento_Change()
     Call objCT.CondicaoPagamento_Change
End Sub

Private Sub CondicaoPagamento_Click()
     Call objCT.CondicaoPagamento_Click
End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)
     Call objCT.CondicaoPagamento_Validate(Cancel)
End Sub

Private Sub CondPagtoLabel_Click()
     Call objCT.CondPagtoLabel_Click
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub DataEntrega_Change()
     Call objCT.DataEntrega_Change
End Sub

Private Sub DataEntrega_GotFocus()
     Call objCT.DataEntrega_GotFocus
End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)
     Call objCT.DataEntrega_KeyPress(KeyAscii)
End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)
     Call objCT.DataEntrega_Validate(Cancel)
End Sub

Private Sub DataReferencia_Change()
     Call objCT.DataReferencia_Change
End Sub

Private Sub DataReferencia_GotFocus()
     Call objCT.DataReferencia_GotFocus
End Sub

Private Sub DataReferencia_Validate(Cancel As Boolean)
     Call objCT.DataReferencia_Validate(Cancel)
End Sub

Private Sub DataVencimento_Change()
     Call objCT.DataVencimento_Change
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
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

Private Sub Desconto1Ate_Change()
     Call objCT.Desconto1Ate_Change
End Sub

Private Sub Desconto1Ate_GotFocus()
     Call objCT.Desconto1Ate_GotFocus
End Sub

Private Sub Desconto1Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto1Ate_Validate(Cancel)
End Sub

Private Sub Desconto1Percentual_Change()
     Call objCT.Desconto1Percentual_Change
End Sub

Private Sub Desconto1Percentual_GotFocus()
     Call objCT.Desconto1Percentual_GotFocus
End Sub

Private Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto1Percentual_Validate(Cancel)
End Sub

Private Sub Desconto1Valor_Change()
     Call objCT.Desconto1Valor_Change
End Sub

Private Sub Desconto1Valor_GotFocus()
     Call objCT.Desconto1Valor_GotFocus
End Sub

Private Sub Desconto1Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto1Valor_Validate(Cancel)
End Sub

Private Sub Desconto2Ate_Change()
     Call objCT.Desconto2Ate_Change
End Sub

Private Sub Desconto2Ate_GotFocus()
     Call objCT.Desconto2Ate_GotFocus
End Sub

Private Sub Desconto2Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto2Ate_Validate(Cancel)
End Sub

Private Sub Desconto2Percentual_Change()
     Call objCT.Desconto2Percentual_Change
End Sub

Private Sub Desconto2Percentual_GotFocus()
     Call objCT.Desconto2Percentual_GotFocus
End Sub

Private Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto2Percentual_Validate(Cancel)
End Sub

Private Sub Desconto2Valor_Change()
     Call objCT.Desconto2Valor_Change
End Sub

Private Sub Desconto2Valor_GotFocus()
     Call objCT.Desconto2Valor_GotFocus
End Sub

Private Sub Desconto2Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto2Valor_Validate(Cancel)
End Sub

Private Sub Desconto3Ate_Change()
     Call objCT.Desconto3Ate_Change
End Sub

Private Sub Desconto3Ate_GotFocus()
     Call objCT.Desconto3Ate_GotFocus
End Sub

Private Sub Desconto3Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto3Ate_Validate(Cancel)
End Sub

Private Sub Desconto3Percentual_Change()
     Call objCT.Desconto3Percentual_Change
End Sub

Private Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto3Percentual_Validate(Cancel)
End Sub

Private Sub Desconto3Valor_Change()
     Call objCT.Desconto3Valor_Change
End Sub

Private Sub Desconto3Valor_GotFocus()
     Call objCT.Desconto3Valor_GotFocus
End Sub

Private Sub Desconto3Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto3Valor_Validate(Cancel)
End Sub

Private Sub DescricaoProduto_Change()
     Call objCT.DescricaoProduto_Change
End Sub

Private Sub DescricaoProduto_GotFocus()
     Call objCT.DescricaoProduto_GotFocus
End Sub

Private Sub DescricaoProduto_KeyPress(KeyAscii As Integer)
     Call objCT.DescricaoProduto_KeyPress(KeyAscii)
End Sub

Private Sub DescricaoProduto_Validate(Cancel As Boolean)
     Call objCT.DescricaoProduto_Validate(Cancel)
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
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

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItens_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)
     Call objCT.GridItens_KeyPress(KeyAscii)
End Sub

Private Sub GridItens_LeaveCell()
     Call objCT.GridItens_LeaveCell
End Sub

Private Sub GridItens_RowColChange()
     Call objCT.GridItens_RowColChange
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
     Call objCT.GridItens_Validate(Cancel)
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub LabelCliente_Click()
     Call objCT.LabelCliente_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub PercAcrescFin_Change()
     Call objCT.PercAcrescFin_Change
End Sub

Private Sub PercAcrescFin_Validate(Cancel As Boolean)
     Call objCT.PercAcrescFin_Validate(Cancel)
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

Private Sub PrazoValidade_Change()
     Call objCT.PrazoValidade_Change
End Sub

Private Sub PrecoTotal_Change()
     Call objCT.PrecoTotal_Change
End Sub

Private Sub PrecoTotal_GotFocus()
     Call objCT.PrecoTotal_GotFocus
End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoTotal_KeyPress(KeyAscii)
End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)
     Call objCT.PrecoTotal_Validate(Cancel)
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

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub TipoDesconto1_Change()
     Call objCT.TipoDesconto1_Change
End Sub

Private Sub TipoDesconto1_GotFocus()
     Call objCT.TipoDesconto1_GotFocus
End Sub

Private Sub TipoDesconto1_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto1_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto1_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto1_Validate(Cancel)
End Sub

Private Sub TipoDesconto2_Change()
     Call objCT.TipoDesconto2_Change
End Sub

Private Sub TipoDesconto2_GotFocus()
     Call objCT.TipoDesconto2_GotFocus
End Sub

Private Sub TipoDesconto2_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto2_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto2_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto2_Validate(Cancel)
End Sub

Private Sub TipoDesconto3_Change()
     Call objCT.TipoDesconto3_Change
End Sub

Private Sub TipoDesconto3_GotFocus()
     Call objCT.TipoDesconto3_GotFocus
End Sub

Private Sub TipoDesconto3_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto3_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto3_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto3_Validate(Cancel)
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

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorDespesas_Change()
     Call objCT.ValorDespesas_Change
End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)
     Call objCT.ValorDespesas_Validate(Cancel)
End Sub

Private Sub ValorFrete_Change()
     Call objCT.ValorFrete_Change
End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)
     Call objCT.ValorFrete_Validate(Cancel)
End Sub

Private Sub ValorParcela_Change()
     Call objCT.ValorParcela_Change
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub ValorSeguro_Change()
     Call objCT.ValorSeguro_Change
End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)
     Call objCT.ValorSeguro_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Private Sub VendedorLabel_Click()
     Call objCT.VendedorLabel_Click
End Sub

Private Sub Vendedor2_Change()
     Call objCT.Vendedor2_Change
End Sub

Private Sub Vendedor2_Validate(Cancel As Boolean)
     Call objCT.Vendedor2_Validate(Cancel)
End Sub

Private Sub Vendedor2Label_Click()
     Call objCT.Vendedor2Label_Click
End Sub

Private Sub BotaoGrade_Click()
     Call objCT.BotaoGrade_Click
End Sub

Public Sub BotaoImprimir_Click()
     Call objCT.BotaoImprimir_Click
End Sub

Public Sub BotaoEmail_Click()
     Call objCT.BotaoEmail_Click
End Sub

Private Sub VersaoKit_Change()
     Call objCT.VersaoKit_Change
End Sub

Private Sub VersaoKit_GotFocus()
     Call objCT.VersaoKit_GotFocus
End Sub

Private Sub VersaoKit_KeyPress(KeyAscii As Integer)
     Call objCT.VersaoKit_KeyPress(KeyAscii)
End Sub

Private Sub VersaoKit_Validate(Cancel As Boolean)
     Call objCT.VersaoKit_Validate(Cancel)
End Sub

Private Sub VersaoKitBase_Change()
     Call objCT.VersaoKitBase_Change
End Sub

Private Sub VersaoKitBase_GotFocus()
     Call objCT.VersaoKitBase_GotFocus
End Sub

Private Sub VersaoKitBase_KeyPress(KeyAscii As Integer)
     Call objCT.VersaoKitBase_KeyPress(KeyAscii)
End Sub

Private Sub VersaoKitBase_Validate(Cancel As Boolean)
     Call objCT.VersaoKitBase_Validate(Cancel)
End Sub

Private Sub NumeroBaseLabel_Click()
     Call objCT.NumeroBaseLabel_Click
End Sub

Private Sub CodigoBase_Change()
     Call objCT.CodigoBase_Change
End Sub

Private Sub CodigoBase_GotFocus()
     Call objCT.CodigoBase_GotFocus
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

'########################################
'Inserido por Wagner 09/01/2006
Private Sub MotivoPerdaItem_Change()
     Call objCT.MotivoPerdaItem_Change
End Sub

Private Sub MotivoPerdaItem_GotFocus()
     Call objCT.MotivoPerdaItem_GotFocus
End Sub

Private Sub MotivoPerdaItem_KeyPress(KeyAscii As Integer)
     Call objCT.MotivoPerdaItem_KeyPress(KeyAscii)
End Sub

Private Sub MotivoPerdaItem_Validate(Cancel As Boolean)
     Call objCT.MotivoPerdaItem_Validate(Cancel)
End Sub

Private Sub StatusItem_Change()
     Call objCT.StatusItem_Change
End Sub

Private Sub StatusItem_GotFocus()
     Call objCT.StatusItem_GotFocus
End Sub

Private Sub StatusItem_KeyPress(KeyAscii As Integer)
     Call objCT.StatusItem_KeyPress(KeyAscii)
End Sub

Private Sub StatusItem_Validate(Cancel As Boolean)
     Call objCT.StatusItem_Validate(Cancel)
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
'########################################

'#####################################
'Inserido por Wagner 18/05/2006
Private Sub BotaoKitVenda_Click()
    Call objCT.BotaoKitVenda_Click
End Sub
'#####################################

'#####################################
'Inserido por Wagner 03/08/2006
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
'#####################################

Private Sub GridPrecosCalculados_Click()
     Call objCT.GridPrecosCalculados_Click
End Sub

Private Sub GridPrecosCalculados_EnterCell()
     Call objCT.GridPrecosCalculados_EnterCell
End Sub

Private Sub GridPrecosCalculados_GotFocus()
     Call objCT.GridPrecosCalculados_GotFocus
End Sub

Private Sub GridPrecosCalculados_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridPrecosCalculados_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridPrecosCalculados_KeyPress(KeyAscii As Integer)
     Call objCT.GridPrecosCalculados_KeyPress(KeyAscii)
End Sub

Private Sub GridPrecosCalculados_LeaveCell()
     Call objCT.GridPrecosCalculados_LeaveCell
End Sub

Private Sub GridPrecosCalculados_RowColChange()
     Call objCT.GridPrecosCalculados_RowColChange
End Sub

Private Sub GridPrecosCalculados_Scroll()
     Call objCT.GridPrecosCalculados_Scroll
End Sub

Private Sub GridPrecosCalculados_Validate(Cancel As Boolean)
     Call objCT.GridPrecosCalculados_Validate(Cancel)
End Sub

Private Sub GridFormacaoPreco_Click()
     Call objCT.GridFormacaoPreco_Click
End Sub

Private Sub GridFormacaoPreco_EnterCell()
     Call objCT.GridFormacaoPreco_EnterCell
End Sub

Private Sub GridFormacaoPreco_GotFocus()
     Call objCT.GridFormacaoPreco_GotFocus
End Sub

Private Sub GridFormacaoPreco_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridFormacaoPreco_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridFormacaoPreco_KeyPress(KeyAscii As Integer)
     Call objCT.GridFormacaoPreco_KeyPress(KeyAscii)
End Sub

Private Sub GridFormacaoPreco_LeaveCell()
     Call objCT.GridFormacaoPreco_LeaveCell
End Sub

Private Sub GridFormacaoPreco_RowColChange()
     Call objCT.GridFormacaoPreco_RowColChange
End Sub

Private Sub GridFormacaoPreco_Scroll()
     Call objCT.GridFormacaoPreco_Scroll
End Sub

Private Sub GridFormacaoPreco_Validate(Cancel As Boolean)
     Call objCT.GridFormacaoPreco_Validate(Cancel)
End Sub

Private Sub PCPrecoUnit_Change()
     Call objCT.PCPrecoUnit_Change
End Sub

Private Sub PCPrecoUnit_GotFocus()
     Call objCT.PCPrecoUnit_GotFocus
End Sub

Private Sub PCPrecoUnit_KeyPress(KeyAscii As Integer)
     Call objCT.PCPrecoUnit_KeyPress(KeyAscii)
End Sub

Private Sub PCPrecoUnit_Validate(Cancel As Boolean)
     Call objCT.PCPrecoUnit_Validate(Cancel)
End Sub

Private Sub PCSituacao_Change()
     Call objCT.PCSituacao_Change
End Sub

Private Sub PCSituacao_GotFocus()
     Call objCT.PCSituacao_GotFocus
End Sub

Private Sub PCSituacao_KeyPress(KeyAscii As Integer)
     Call objCT.PCSituacao_KeyPress(KeyAscii)
End Sub

Private Sub PCSituacao_Validate(Cancel As Boolean)
     Call objCT.PCSituacao_Validate(Cancel)
End Sub

Private Sub FPPercentMargem_Change()
     Call objCT.FPPercentMargem_Change
End Sub

Private Sub FPPercentMargem_GotFocus()
     Call objCT.FPPercentMargem_GotFocus
End Sub

Private Sub FPPercentMargem_KeyPress(KeyAscii As Integer)
     Call objCT.FPPercentMargem_KeyPress(KeyAscii)
End Sub

Private Sub FPPercentMargem_Validate(Cancel As Boolean)
     Call objCT.FPPercentMargem_Validate(Cancel)
End Sub

Private Sub FPSituacao_Change()
     Call objCT.FPSituacao_Change
End Sub

Private Sub FPSituacao_GotFocus()
     Call objCT.FPSituacao_GotFocus
End Sub

Private Sub FPSituacao_KeyPress(KeyAscii As Integer)
     Call objCT.FPSituacao_KeyPress(KeyAscii)
End Sub

Private Sub FPSituacao_Validate(Cancel As Boolean)
     Call objCT.FPSituacao_Validate(Cancel)
End Sub

Private Sub PCSelecionado_GotFocus()
     Call objCT.PCSelecionado_GotFocus
End Sub

Private Sub PCSelecionado_KeyPress(KeyAscii As Integer)
     Call objCT.PCSelecionado_KeyPress(KeyAscii)
End Sub

Private Sub PCSelecionado_Validate(Cancel As Boolean)
     Call objCT.PCSelecionado_Validate(Cancel)
End Sub

Private Sub PCSelecionado_Click()
     Call objCT.PCSelecionado_Click
End Sub

Private Sub BotaoVersoesOVs_Click()
    Call objCT.BotaoVersoesOVs_Click
End Sub

Private Sub BotaoOVVersoes_Click()
    Call objCT.BotaoOVVersoes_Click
End Sub

Private Sub ValorDescontoItens_Change()
     Call objCT.ValorDescontoItens_Change
End Sub

Private Sub ValorDescontoItens_Validate(Cancel As Boolean)
     Call objCT.ValorDescontoItens_Validate(Cancel)
End Sub

Private Sub PercDescontoItens_Change()
     Call objCT.PercDescontoItens_Change
End Sub

Private Sub PercDescontoItens_Validate(Cancel As Boolean)
     Call objCT.PercDescontoItens_Validate(Cancel)
End Sub

Private Sub ValorDescontoTit_Change()
     Call objCT.ValorDescontoTit_Change
End Sub

Private Sub ValorDescontoTit_Validate(Cancel As Boolean)
     Call objCT.ValorDescontoTit_Validate(Cancel)
End Sub

Private Sub BotaoInfoAdic_Click()
     Call objCT.BotaoInfoAdic_Click
End Sub

Private Sub Contato_Change()
     Call objCT.Contato_Change
End Sub

Private Sub Email_Change()
     Call objCT.Email_Change
End Sub

Private Sub BotaoInfoAdicItem_Click()
    Call objCT.BotaoInfoAdicItem_Click
End Sub

Private Sub PrazoEntregaItem_GotFocus()
     Call objCT.PrazoEntregaItem_GotFocus
End Sub

Private Sub PrazoEntregaItem_KeyPress(KeyAscii As Integer)
     Call objCT.PrazoEntregaItem_KeyPress(KeyAscii)
End Sub

Private Sub PrazoEntregaItem_Validate(Cancel As Boolean)
     Call objCT.PrazoEntregaItem_Validate(Cancel)
End Sub

Private Sub FilialEntrega_Change()
     Call objCT.FilialEntrega_Change
End Sub

Private Sub FilialEntrega_Click()
     Call objCT.FilialEntrega_Click
End Sub

Private Sub FilialEntrega_Validate(Cancel As Boolean)
     Call objCT.FilialEntrega_Validate(Cancel)
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub TipoFrete_Click()
    Call objCT.TipoFrete_Click
End Sub

Private Sub Placa_Change()
     Call objCT.Placa_Change
End Sub

Private Sub PlacaUF_Change()
     Call objCT.PlacaUF_Change
End Sub

Private Sub PlacaUF_Click()
     Call objCT.PlacaUF_Click
End Sub

Private Sub PlacaUF_Validate(Cancel As Boolean)
     Call objCT.PlacaUF_Validate(Cancel)
End Sub

Private Sub DataEntregaPV_Change()
    Call objCT.DataEntregaPV_Change
End Sub

Private Sub DataEntregaPV_GotFocus()
    Call objCT.DataEntregaPV_GotFocus
End Sub

Private Sub DataEntregaPV_Validate(Cancel As Boolean)
    Call objCT.DataEntregaPV_Validate(Cancel)
End Sub

Private Sub TranspRedespacho_Change()
     Call objCT.TranspRedespacho_Change
End Sub

Private Sub TranspRedespacho_Click()
     Call objCT.TranspRedespacho_Click
End Sub

Private Sub TranspRedespacho_Validate(Cancel As Boolean)
     Call objCT.TranspRedespacho_Validate(Cancel)
End Sub

Private Sub Cubagem_Change()
     Call objCT.Cubagem_Change
End Sub

Private Sub Cubagem_Validate(Cancel As Boolean)
    Call objCT.Cubagem_Validate(Cancel)
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub TranspRedLabel_Click()
     Call objCT.TranspRedLabel_Click
End Sub

Private Sub RedespachoCli_Click()
    Call objCT.RedespachoCli_Click
End Sub

Private Sub UpDownEntregaPV_DownClick()
    Call objCT.UpDownEntregaPV_DownClick
End Sub

Private Sub UpDownEntregaPV_UpClick()
    Call objCT.UpDownEntregaPV_UpClick
End Sub

Private Sub VolumeEspecie_Change()
     Call objCT.VolumeEspecie_Change
End Sub

Private Sub VolumeMarca_Change()
     Call objCT.VolumeMarca_Change
End Sub

Private Sub VolumeEspecie_Validate(Cancel As Boolean)
     Call objCT.VolumeEspecie_Validate(Cancel)
End Sub

Private Sub VolumeMarca_Validate(Cancel As Boolean)
     Call objCT.VolumeMarca_Validate(Cancel)
End Sub

Private Sub VolumeNumero_Change()
     Call objCT.VolumeNumero_Change
End Sub

Private Sub VolumeQuant_Change()
     Call objCT.VolumeQuant_Change
End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)
     Call objCT.PesoLiquido_Validate(Cancel)
End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)
     Call objCT.PesoBruto_Validate(Cancel)
End Sub

Private Sub VolumeQuant_GotFocus()
     Call objCT.VolumeQuant_GotFocus
End Sub

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub PesoBruto_Change()
     Call objCT.PesoBruto_Change
End Sub

Private Sub PesoLiquido_Change()
     Call objCT.PesoLiquido_Change
End Sub

Private Sub CanalVenda_Change()
     Call objCT.CanalVenda_Change
End Sub

Private Sub CanalVenda_Click()
     Call objCT.CanalVenda_Click
End Sub

Private Sub CanalVenda_Validate(Cancel As Boolean)
     Call objCT.CanalVenda_Validate(Cancel)
End Sub

Private Sub PedidoRepr_Change()
     Call objCT.PedidoRepr_Change
End Sub

Private Sub PedidoRepr_Validate(Cancel As Boolean)
     Call objCT.PedidoRepr_Validate(Cancel)
End Sub

Private Sub PedidoCliente_Change()
     Call objCT.PedidoCliente_Change
End Sub

Private Sub DataEnvio_Change()
    Call objCT.DataEnvio_Change
End Sub

Private Sub DataEnvio_GotFocus()
    Call objCT.DataEnvio_GotFocus
End Sub

Private Sub DataEnvio_Validate(Cancel As Boolean)
    Call objCT.DataEnvio_Validate(Cancel)
End Sub

Private Sub UpDownEnvio_DownClick()
    Call objCT.UpDownEnvio_DownClick
End Sub

Private Sub UpDownEnvio_UpClick()
    Call objCT.UpDownEnvio_UpClick
End Sub

Private Sub PrazoEntrega_Change()
    Call objCT.PrazoEntrega_Change
End Sub

Private Sub PrazoEntrega_GotFocus()
    Call objCT.PrazoEntrega_GotFocus
End Sub

Private Sub PrazoEntrega_Validate(Cancel As Boolean)
    Call objCT.PrazoEntrega_Validate(Cancel)
End Sub

Private Sub MensagemLabel_Click()
    Call objCT.MensagemLabel_Click
End Sub

Private Sub OptDataPrazoEnt_Click(Index As Integer)
    Call objCT.OptDataPrazoEnt_Click(Index)
End Sub

Private Sub BotaoEstoqueProd_Click()
     Call objCT.BotaoEstoqueProd_Click
End Sub

Private Sub DataPerda_Change()
    Call objCT.DataPerda_Change
End Sub

Private Sub DataPerda_GotFocus()
    Call objCT.DataPerda_GotFocus
End Sub

Private Sub DataPerda_Validate(Cancel As Boolean)
    Call objCT.DataPerda_Validate(Cancel)
End Sub

Private Sub UpDownPerda_DownClick()
    Call objCT.UpDownPerda_DownClick
End Sub

Private Sub UpDownPerda_UpClick()
    Call objCT.UpDownPerda_UpClick
End Sub

Private Sub Moeda_Change()
    Call objCT.Moeda_Change
End Sub

Private Sub Moeda_Click()
    Call objCT.Moeda_Click
End Sub

Private Sub PrazoTexto_Change()
    Call objCT.PrazoTexto_Change
End Sub

Private Sub PrazoTexto_Click()
    Call objCT.PrazoTexto_Click
End Sub

Private Sub PrazoEntVar_Click()
    Call objCT.PrazoEntVar_Click
End Sub

Private Sub MotivoPerda_Change()
    Call objCT.MotivoPerda_Change
End Sub

Private Sub MotivoPerda_Click()
    Call objCT.MotivoPerda_Click
End Sub

Private Sub BotaoDesce_Click()
    Call objCT.BotaoDesce_Click
End Sub

Private Sub BotaoSobe_Click()
    Call objCT.BotaoSobe_Click
End Sub

Private Sub SomaItem_Click()
    Call objCT.SomaItem_Click
End Sub

Private Sub DataFechamento_Change()
    Call objCT.DataFechamento_Change
End Sub

Private Sub DataFechamento_GotFocus()
    Call objCT.DataFechamento_GotFocus
End Sub

Private Sub DataFechamento_Validate(Cancel As Boolean)
    Call objCT.DataFechamento_Validate(Cancel)
End Sub

Private Sub UpDownFechamento_DownClick()
    Call objCT.UpDownFechamento_DownClick
End Sub

Private Sub UpDownFechamento_UpClick()
    Call objCT.UpDownFechamento_UpClick
End Sub

Private Sub DataProxContato_Change()
    Call objCT.DataProxContato_Change
End Sub

Private Sub DataProxContato_GotFocus()
    Call objCT.DataProxContato_GotFocus
End Sub

Private Sub DataProxContato_Validate(Cancel As Boolean)
    Call objCT.DataProxContato_Validate(Cancel)
End Sub

Private Sub UpDownProxContato_DownClick()
    Call objCT.UpDownProxContato_DownClick
End Sub

Private Sub UpDownProxContato_UpClick()
    Call objCT.UpDownProxContato_UpClick
End Sub

Private Sub Idioma_Change()
    Call objCT.Idioma_Change
End Sub

Private Sub Idioma_Click()
    Call objCT.Idioma_Click
End Sub

Private Sub BtnExportaDoc_Click()
    Call objCT.BtnExportaDoc_Click
End Sub
