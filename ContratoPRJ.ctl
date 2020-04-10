VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ContratoPRJ 
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4995
      Index           =   3
      Left            =   45
      TabIndex        =   68
      Top             =   1005
      Visible         =   0   'False
      Width           =   9345
      Begin VB.Frame Frame2 
         Caption         =   "Totais"
         Height          =   1290
         Index           =   1
         Left            =   60
         TabIndex        =   147
         Top             =   3270
         Width           =   9285
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   90
            TabIndex        =   148
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
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
            TabIndex        =   149
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
            Left            =   2745
            TabIndex        =   150
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   1410
            TabIndex        =   151
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercDescontoItens 
            Height          =   285
            Left            =   4065
            TabIndex        =   152
            ToolTipText     =   "Percentual de desconto dos itens"
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDescontoItens 
            Height          =   285
            Left            =   5400
            TabIndex        =   153
            ToolTipText     =   "Soma dos descontos dos itens"
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   177
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   176
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   90
            TabIndex        =   175
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1410
            TabIndex        =   174
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2745
            TabIndex        =   173
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4065
            TabIndex        =   172
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   171
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label Label1 
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
            Index           =   28
            Left            =   165
            TabIndex        =   170
            Top             =   195
            Width           =   1020
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
            Index           =   27
            Left            =   1470
            TabIndex        =   169
            Top             =   195
            Width           =   1065
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
            Index           =   26
            Left            =   2745
            TabIndex        =   168
            Top             =   210
            Width           =   1170
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
            Index           =   25
            Left            =   4080
            TabIndex        =   167
            Top             =   210
            Width           =   1065
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
            Index           =   11
            Left            =   8100
            TabIndex        =   166
            Top             =   210
            Width           =   1065
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
            Index           =   22
            Left            =   6735
            TabIndex        =   165
            Top             =   210
            Width           =   1065
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
            Index           =   10
            Left            =   4125
            TabIndex        =   164
            Top             =   705
            Width           =   1065
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
            Index           =   7
            Left            =   5430
            TabIndex        =   163
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label ISSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   162
            Top             =   405
            Width           =   1140
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
            Left            =   5430
            TabIndex        =   161
            Top             =   210
            Width           =   1065
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
            Index           =   6
            Left            =   105
            TabIndex        =   160
            Top             =   705
            Width           =   1125
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
            Index           =   4
            Left            =   1470
            TabIndex        =   159
            Top             =   705
            Width           =   1125
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
            Index           =   3
            Left            =   2790
            TabIndex        =   158
            Top             =   705
            Width           =   1125
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
            Index           =   8
            Left            =   6735
            TabIndex        =   157
            Top             =   705
            Width           =   1125
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
            Left            =   8085
            TabIndex        =   156
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label ValorProdutos2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   155
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   154
            Top             =   405
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin VB.CommandButton BotaoEtapas 
         Caption         =   "Etapas"
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
         Left            =   105
         TabIndex        =   30
         Top             =   4590
         Width           =   1575
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
         Height          =   360
         Left            =   2070
         TabIndex        =   31
         Top             =   4590
         Width           =   1575
      End
      Begin VB.CommandButton BotaoRefazer 
         Caption         =   "Refazer Proposta"
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
         Left            =   6795
         TabIndex        =   33
         Top             =   4590
         Width           =   2385
      End
      Begin VB.CommandButton BotaoCondPag 
         Caption         =   "Condições de Pagamento"
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
         Left            =   3990
         TabIndex        =   32
         Top             =   4590
         Width           =   2385
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   3270
         Index           =   3
         Left            =   60
         TabIndex        =   69
         Top             =   -45
         Width           =   9285
         Begin MSMask.MaskEdBox PrecoTotalB 
            Height          =   225
            Left            =   7260
            TabIndex        =   178
            Top             =   750
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
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3060
            MaxLength       =   250
            TabIndex        =   93
            Top             =   1260
            Width           =   2490
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   1755
            TabIndex        =   92
            Top             =   1545
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "ContratoPRJ.ctx":0000
            Left            =   2100
            List            =   "ContratoPRJ.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   465
            Width           =   720
         End
         Begin VB.TextBox DescEtapa 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   5160
            MaxLength       =   250
            TabIndex        =   72
            Top             =   1230
            Width           =   2490
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3855
            MaxLength       =   50
            TabIndex        =   70
            Top             =   795
            Width           =   2250
         End
         Begin MSMask.MaskEdBox Etapa 
            Height          =   225
            Left            =   825
            TabIndex        =   71
            Top             =   555
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   3165
            TabIndex        =   74
            Top             =   885
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
            Left            =   1950
            TabIndex        =   75
            Top             =   795
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
            Left            =   795
            TabIndex        =   76
            Top             =   900
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
            Left            =   4680
            TabIndex        =   77
            Top             =   570
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   3105
            TabIndex        =   78
            Top             =   540
            Width           =   1500
            _ExtentX        =   2646
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
            Left            =   6195
            TabIndex        =   79
            Top             =   570
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
            TabIndex        =   29
            Top             =   210
            Width           =   9225
            _ExtentX        =   16272
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
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5055
      Index           =   2
      Left            =   105
      TabIndex        =   81
      Top             =   975
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame FrameE 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4425
         Index           =   1
         Left            =   75
         TabIndex        =   114
         Top             =   435
         Width           =   9045
         Begin VB.CheckBox ExibirProdutos 
            Caption         =   "Exibir Produtos"
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
            TabIndex        =   22
            Top             =   0
            Width           =   1695
         End
         Begin VB.CheckBox ExibirCustoCalc 
            Caption         =   "Exibir Custo Caculado"
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
            Left            =   1995
            TabIndex        =   23
            Top             =   15
            Width           =   2355
         End
         Begin VB.CheckBox ExibirCustoInf 
            Caption         =   "Exibir Custo Informado"
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
            TabIndex        =   24
            Top             =   30
            Width           =   2355
         End
         Begin VB.CheckBox ExibirPreco 
            Caption         =   "Exibir Preço"
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
            Left            =   7080
            TabIndex        =   25
            Top             =   45
            Width           =   1695
         End
         Begin VB.Frame Frame1 
            Caption         =   "Totais"
            Height          =   570
            Index           =   12
            Left            =   90
            TabIndex        =   115
            Top             =   2535
            Width           =   8880
            Begin VB.Label TotalCustoInf 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1695
               TabIndex        =   121
               Top             =   195
               Width           =   1380
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Custo Informado:"
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
               Index           =   19
               Left            =   180
               TabIndex        =   120
               Top             =   255
               Width           =   1455
            End
            Begin VB.Label TotalCustoCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4950
               TabIndex        =   119
               Top             =   180
               Width           =   1380
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Custo Calculado:"
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
               Index           =   20
               Left            =   3375
               TabIndex        =   118
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label TotalPreco 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   7260
               TabIndex        =   117
               Top             =   180
               Width           =   1410
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Preço:"
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
               Index           =   23
               Left            =   6570
               TabIndex        =   116
               Top             =   240
               Width           =   570
            End
         End
         Begin MSComctlLib.TreeView TvwEtapas 
            Height          =   2250
            Left            =   90
            TabIndex        =   26
            Top             =   240
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   3969
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   6
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
         Begin VB.Frame Frame1 
            Caption         =   "Etapa/Produto"
            Height          =   1305
            Index           =   11
            Left            =   90
            TabIndex        =   122
            Top             =   3105
            Width           =   8880
            Begin MSMask.MaskEdBox CustoInformado 
               Height          =   315
               Left            =   1695
               TabIndex        =   27
               Top             =   540
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox Preco 
               Height          =   315
               Left            =   1695
               TabIndex        =   28
               Top             =   900
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   556
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
            Begin VB.Label Label1 
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
               Index           =   5
               Left            =   705
               TabIndex        =   128
               Top             =   240
               Width           =   930
            End
            Begin VB.Label CustoCalculado 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   7245
               TabIndex        =   127
               Top             =   555
               Width           =   1425
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Custo Calculado:"
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
               Left            =   5700
               TabIndex        =   126
               Top             =   615
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Custo Informado:"
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
               Left            =   180
               TabIndex        =   125
               Top             =   585
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Preço:"
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
               Left            =   990
               TabIndex        =   124
               Top             =   945
               Width           =   570
            End
            Begin VB.Label DescricaoEtapa 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1695
               TabIndex        =   123
               Top             =   180
               Width           =   6990
            End
         End
      End
      Begin VB.Frame FrameE 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4545
         Index           =   2
         Left            =   60
         TabIndex        =   130
         Top             =   405
         Visible         =   0   'False
         Width           =   9045
         Begin MSMask.MaskEdBox ObservacaoGrid 
            Height          =   270
            Left            =   4920
            TabIndex        =   135
            Top             =   735
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DescricaoGrid 
            Height          =   270
            Left            =   2235
            TabIndex        =   134
            Top             =   750
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.CheckBox Imprimir 
            DragMode        =   1  'Automatic
            Height          =   270
            Left            =   345
            TabIndex        =   133
            Top             =   720
            Width           =   675
         End
         Begin MSMask.MaskEdBox EtapaGrid 
            Height          =   270
            Left            =   1155
            TabIndex        =   132
            Top             =   720
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox ObsEtapa 
            Height          =   1050
            Left            =   1260
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   3345
            Width           =   7680
         End
         Begin MSFlexGridLib.MSFlexGrid GridEtapa 
            Height          =   3075
            Left            =   105
            TabIndex        =   20
            Top             =   0
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   5424
            _Version        =   393216
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   24
            Left            =   150
            TabIndex        =   131
            Top             =   3330
            Width           =   1080
         End
      End
      Begin MSComctlLib.TabStrip OpcaoEtapa 
         Height          =   4980
         Left            =   30
         TabIndex        =   129
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   8784
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhamento de cobrança"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Impressão"
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
      Caption         =   "Frame1"
      Height          =   5070
      Index           =   1
      Left            =   90
      TabIndex        =   64
      Top             =   990
      Width           =   9225
      Begin VB.Frame Frame1 
         Caption         =   "Identificação"
         Height          =   1995
         Index           =   10
         Left            =   120
         TabIndex        =   84
         Top             =   15
         Width           =   9075
         Begin VB.TextBox Obs 
            Height          =   810
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1065
            Width           =   7545
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1335
            TabIndex        =   2
            Top             =   660
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   315
            Left            =   1335
            TabIndex        =   0
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzidoPRJ 
            Height          =   315
            Left            =   4815
            TabIndex        =   1
            Top             =   240
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataCriacao 
            Height          =   300
            Left            =   5970
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   660
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCriacao 
            Height          =   315
            Left            =   4815
            TabIndex        =   3
            Top             =   660
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
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
            Left            =   8370
            TabIndex        =   146
            Top             =   660
            Width           =   525
         End
         Begin VB.Label LblNatOpInternaEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Oper.:"
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
            Left            =   6645
            TabIndex        =   145
            Top             =   720
            Width           =   1650
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   1035
            Width           =   1095
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   17
            Left            =   4260
            TabIndex        =   90
            Top             =   735
            Width           =   480
         End
         Begin VB.Label LabelNomeRedPRJ 
            Caption         =   "Nome do Projeto:"
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
            Height          =   315
            Left            =   3270
            TabIndex        =   87
            Top             =   285
            Width           =   1560
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   600
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   86
            Top             =   285
            Width           =   675
         End
         Begin VB.Label LabelCodigo 
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
            Left            =   615
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   85
            Top             =   705
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Geração"
         Height          =   2355
         Index           =   6
         Left            =   135
         TabIndex        =   83
         Top             =   2685
         Width           =   9075
         Begin VB.TextBox NomeDiretorio 
            Height          =   315
            Left            =   5490
            TabIndex        =   12
            ToolTipText     =   "Diretório onde a proposta será gerada"
            Top             =   615
            Width           =   3375
         End
         Begin VB.DirListBox Dir1 
            Height          =   1215
            Left            =   1335
            TabIndex        =   11
            Top             =   1035
            Width           =   3180
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   1335
            TabIndex        =   10
            Top             =   645
            Width           =   2175
         End
         Begin VB.CommandButton BotaoMnemonicos 
            Caption         =   "Mnemônicos Válidos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   5535
            TabIndex        =   14
            ToolTipText     =   "Mnemônicos válidos para utilização em modelo do word"
            Top             =   1545
            Width           =   1515
         End
         Begin VB.TextBox Modelo 
            Height          =   315
            Left            =   1335
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   8
            ToolTipText     =   "Modelo base para geração da proposta (.doc)"
            Top             =   195
            Width           =   6990
         End
         Begin VB.CommandButton BotaoProcurar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8370
            TabIndex        =   9
            Top             =   195
            Width           =   495
         End
         Begin VB.CommandButton BotaoGerarArq 
            Caption         =   "Gerar Arquivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   7350
            TabIndex        =   15
            ToolTipText     =   "Gera um arquivo de proposta com base no modelo escolhido"
            Top             =   1530
            Width           =   1515
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   405
            Top             =   1755
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox NomeArquivo 
            Height          =   315
            Left            =   5505
            MaxLength       =   80
            TabIndex        =   13
            ToolTipText     =   "Nome do arquivo de proposta a ser gerado"
            Top             =   1035
            Width           =   3360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Diretório:"
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
            Left            =   4620
            TabIndex        =   136
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Arquivo:"
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
            Index           =   16
            Left            =   4755
            TabIndex        =   89
            Top             =   1095
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Modelo:"
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
            Index           =   9
            Left            =   570
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   88
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados do Cliente"
         Height          =   630
         Index           =   91
         Left            =   120
         TabIndex        =   65
         Top             =   2025
         Width           =   9075
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4845
            TabIndex        =   7
            Top             =   165
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1335
            TabIndex        =   6
            Top             =   210
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   600
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   67
            Top             =   240
            Width           =   660
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
            Left            =   4290
            TabIndex        =   66
            Top             =   255
            Width           =   465
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   5055
      Index           =   4
      Left            =   120
      TabIndex        =   80
      Top             =   960
      Visible         =   0   'False
      Width           =   9240
      Begin TelasPRJ.TabTributacaoFat TabTrib 
         Height          =   4845
         Left            =   150
         TabIndex        =   144
         Top             =   180
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   8546
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Proposta"
      Height          =   645
      Index           =   0
      Left            =   75
      TabIndex        =   137
      Top             =   -30
      Width           =   7260
      Begin VB.CommandButton BotaoProposta 
         Caption         =   "Trazer Dados"
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
         Left            =   4785
         TabIndex        =   139
         Top             =   240
         Width           =   1350
      End
      Begin VB.CommandButton BotaoVerProposta 
         Caption         =   "Consultar"
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
         Left            =   6210
         TabIndex        =   138
         Top             =   240
         Width           =   1005
      End
      Begin MSMask.MaskEdBox PRJ 
         Height          =   300
         Left            =   780
         TabIndex        =   140
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Proposta 
         Height          =   300
         Left            =   2985
         TabIndex        =   142
         Top             =   240
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProposta 
         AutoSize        =   -1  'True
         Caption         =   "Proposta:"
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
         Left            =   2160
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   143
         Top             =   285
         Width           =   825
      End
      Begin VB.Label PRJLabel 
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
         Height          =   255
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   141
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5010
      Index           =   5
      Left            =   135
      TabIndex        =   91
      Top             =   975
      Visible         =   0   'False
      Width           =   9240
      Begin VB.ComboBox Controles 
         Height          =   315
         ItemData        =   "ContratoPRJ.ctx":0004
         Left            =   7395
         List            =   "ContratoPRJ.ctx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   4515
         Width           =   1770
      End
      Begin VB.CommandButton BotaoDadosCustNovo 
         Height          =   405
         Left            =   6435
         Picture         =   "ContratoPRJ.ctx":0034
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   4455
         Width           =   435
      End
      Begin VB.CommandButton BotaoDadosCustDel 
         Height          =   405
         Left            =   6930
         Picture         =   "ContratoPRJ.ctx":0546
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4455
         Width           =   435
      End
      Begin VB.TextBox Texto 
         Height          =   315
         Index           =   1
         Left            =   3960
         MaxLength       =   255
         TabIndex        =   44
         Top             =   255
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Height          =   315
         Index           =   3
         Left            =   3960
         MaxLength       =   255
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Height          =   315
         Index           =   2
         Left            =   3960
         MaxLength       =   255
         TabIndex        =   45
         Top             =   615
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Height          =   315
         Index           =   4
         Left            =   3960
         MaxLength       =   255
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Height          =   315
         Index           =   5
         Left            =   3960
         MaxLength       =   255
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   4755
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   1
         Left            =   2445
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   1
         Left            =   1275
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   2
         Left            =   2445
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   615
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   2
         Left            =   1275
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   3
         Left            =   2445
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   3
         Left            =   1275
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   4
         Left            =   2445
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   4
         Left            =   1275
         TabIndex        =   40
         Top             =   1335
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   5
         Left            =   2445
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1710
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   5
         Left            =   1275
         TabIndex        =   42
         Top             =   1695
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   1
         Left            =   1275
         TabIndex        =   49
         Top             =   2160
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   2
         Left            =   1275
         TabIndex        =   50
         Top             =   2520
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   3
         Left            =   1275
         TabIndex        =   51
         Top             =   2880
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   4
         Left            =   1275
         TabIndex        =   52
         Top             =   3240
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   5
         Left            =   1275
         TabIndex        =   53
         Top             =   3615
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   54
         Top             =   2160
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   2
         Left            =   3960
         TabIndex        =   55
         Top             =   2520
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   3
         Left            =   3960
         TabIndex        =   56
         Top             =   2880
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   4
         Left            =   3960
         TabIndex        =   57
         Top             =   3240
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   5
         Left            =   3960
         TabIndex        =   58
         Top             =   3615
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data1:"
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
         Index           =   1001
         Left            =   120
         TabIndex        =   113
         Top             =   315
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data2:"
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
         Index           =   1002
         Left            =   135
         TabIndex        =   112
         Top             =   660
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data3:"
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
         Index           =   1003
         Left            =   135
         TabIndex        =   111
         Top             =   990
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data4:"
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
         Index           =   1004
         Left            =   135
         TabIndex        =   110
         Top             =   1380
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data5:"
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
         Index           =   1005
         Left            =   135
         TabIndex        =   109
         Top             =   1740
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto1:"
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
         Index           =   4001
         Left            =   2910
         TabIndex        =   108
         Top             =   300
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto2:"
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
         Index           =   4002
         Left            =   3225
         TabIndex        =   107
         Top             =   660
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto3:"
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
         Index           =   4003
         Left            =   3225
         TabIndex        =   106
         Top             =   990
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto4:"
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
         Index           =   4004
         Left            =   3225
         TabIndex        =   105
         Top             =   1380
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto5:"
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
         Index           =   4005
         Left            =   3225
         TabIndex        =   104
         Top             =   1740
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número1:"
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
         Height          =   285
         Index           =   3001
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   103
         Top             =   2205
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número2:"
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
         Height          =   285
         Index           =   3002
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   102
         Top             =   2565
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número3:"
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
         Height          =   285
         Index           =   3003
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   101
         Top             =   2925
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número4:"
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
         Height          =   285
         Index           =   3004
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   100
         Top             =   3285
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número5:"
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
         Height          =   285
         Index           =   3005
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   99
         Top             =   3660
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor1:"
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
         Index           =   2001
         Left            =   2610
         TabIndex        =   98
         Top             =   2220
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor2:"
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
         Index           =   2002
         Left            =   2610
         TabIndex        =   97
         Top             =   2595
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor3:"
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
         Index           =   2003
         Left            =   2610
         TabIndex        =   96
         Top             =   2985
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor4:"
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
         Index           =   2004
         Left            =   2610
         TabIndex        =   95
         Top             =   3345
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor5:"
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
         Index           =   2005
         Left            =   2610
         TabIndex        =   94
         Top             =   3705
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   570
      Left            =   7425
      ScaleHeight     =   510
      ScaleWidth      =   1890
      TabIndex        =   62
      Top             =   45
      Width           =   1950
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   30
         Picture         =   "ContratoPRJ.ctx":09FC
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   495
         Picture         =   "ContratoPRJ.ctx":0B56
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   945
         Picture         =   "ContratoPRJ.ctx":0CE0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1410
         Picture         =   "ContratoPRJ.ctx":1212
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5445
      Left            =   15
      TabIndex        =   63
      Top             =   645
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   9604
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Iniciais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Etapas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Customizados"
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
Attribute VB_Name = "ContratoPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

''constantes do word
'Private Const wdCharacter = 1
'Private Const wdGoToField = 7
'Private Const wdWord9TableBehavior = 1
'Private Const wdAutoFitContent = 1
'Private Const wdGoToLine = 3
'Private Const wdCell = 12

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim dValorDescontoItensAnt As Double
Dim dPercDescontoItensAnt As Double
Dim iDescontoAlterado As Integer

Dim iListIndexDefault As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProjeto As AdmEvento
Attribute objEventoProjeto.VB_VarHelpID = -1
Private WithEvents objEventoPRJ As AdmEvento
Attribute objEventoPRJ.VB_VarHelpID = -1
Private WithEvents objEventoProposta As AdmEvento
Attribute objEventoProposta.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoEtapa As AdmEvento
Attribute objEventoEtapa.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
Private WithEvents objEventoTiposDeTributacao As AdmEvento
Attribute objEventoTiposDeTributacao.VB_VarHelpID = -1

Dim gobjProjeto As ClassProjetos
Public gobjProposta As ClassPRJPropostas
Dim gobjEtapa As ClassPRJEtapas
Dim gobjEtapaIP As ClassPRJEtapaItensProd
Dim gobjTelaCamposCust As ClassTelaDadosCust

Dim sPRJAnt As String

Public iFrameAtual As Integer
Dim iFrameAtualEtapa As Integer
Public iAlterado As Integer
'Dim iFrameAtualTributacao As Integer

Public giClienteAlterado As Integer
Public giFilialAlterada As Integer
Public gdDesconto As Double
Public giValorFreteAlterado As Integer
Public giValorSeguroAlterado As Integer
Public giValorDescontoAlterado As Integer
Public giValorDespesasAlterado  As Integer
Public giDataReferenciaAlterada As Integer
Public giNaturezaOpAlterada As Integer
Public giValorDescontoManual As Integer
Public giPercAcresFinAlterado As Integer
Public gobjContrato As New ClassPRJContratos
Public giLinhaAnterior As Integer
'Dim giRecalculandoTributacao As Integer
'Dim gcolTiposTribICMS As New Collection
'Dim gcolTiposTribIPI As New Collection

Public gobjTribTab As New ClassTribTab
'Dim giISSAliquotaAlterada As Integer
'Dim giISSValorAlterado As Integer
'Dim giValorIRRFAlterado As Integer
'Dim giTipoTributacaoAlterado As Integer
'Dim giAliqIRAlterada As Integer
'Dim iPISRetidoAlterado As Integer
'Dim iISSRetidoAlterado As Integer
'Dim iCOFINSRetidoAlterado As Integer
'Dim iCSLLRetidoAlterado As Integer
'
'Dim giTrazendoTribTela As Integer
'Dim giTrazendoTribItemTela As Integer
'Dim giNatOpItemAlterado As Integer
'Dim giTipoTributacaoItemAlterado As Integer
'Dim giICMSBaseItemAlterado As Integer
'Dim giICMSPercRedBaseItemAlterado As Integer
'Dim giICMSAliquotaItemAlterado As Integer
'Dim giICMSValorItemAlterado As Integer
'Dim giICMSSubstBaseItemAlterado As Integer
'Dim giICMSSubstAliquotaItemAlterado As Integer
'Dim giICMSSubstValorItemAlterado As Integer
'Dim giIPIBaseItemAlterado As Integer
'Dim giIPIPercRedBaseItemAlterado As Integer
'Dim giIPIAliquotaItemAlterado As Integer
'Dim giIPIValorItemAlterado As Integer
Public gbCarregandoTela As Boolean

Public iLinhaAnt As Integer

Dim sProjetoAnt As String
Dim sNomeProjetoAnt As String

Public objGridItens As AdmGrid
Public iGrid_Etapa_Col As Integer
Public iGrid_DescEtapa_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_DataEntrega_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_Observacao_Col As Integer

Dim objGridEtapa As AdmGrid
Dim iGrid_EtapaGrid_Col As Integer
Dim iGrid_DescricaoGrid_Col As Integer
Dim iGrid_Imprimir_Col As Integer
Dim iGrid_ObservacaoGrid_Col As Integer

'Constantes públicas dos tabs
Private Const TAB_PRINCIPAL = 1
Private Const TAB_ETAPA = 2
Private Const TAB_ITENS = 3
Private Const TAB_TRIBUTACAO = 4
Private Const TAB_DADOSCUSTOMIZADOS = 5

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Contratos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ContratoPRJ"

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

Private Sub UserControl_Initialize()
    Set gobjTelaCamposCust = New ClassTelaDadosCust
    Set gobjTelaCamposCust.objUserControl = Me
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is NomeReduzidoPRJ Then
            Call LabelNomeRedPRJ_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is Etapa Then
            Call BotaoEtapas_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
'        ElseIf Me.ActiveControl Is NaturezaOpItem Then
'            Call NaturezaItemLabel_Click
'        ElseIf Me.ActiveControl Is TipoTributacaoItem Then
'            Call LblTipoTribItem_Click
'        ElseIf Me.ActiveControl Is TipoTributacao Then
'            Call LblTipoTrib_Click
'        ElseIf Me.ActiveControl Is NaturezaOp Then
'            Call NaturezaLabel_Click
        ElseIf Me.ActiveControl Is PRJ Then
            Call PRJLabel_Click
        ElseIf Me.ActiveControl Is Proposta Then
            Call LabelProposta_Click
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProjeto = Nothing
    Set objEventoPRJ = Nothing
    Set objEventoProposta = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEtapa = Nothing
    Set gobjTelaCamposCust = Nothing
    Set gobjProposta = Nothing
    
    Set objGridItens = Nothing
    Set objGridEtapa = Nothing

    'Encerra tributacao
    Set gobjTribTab = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187606)

    End Select

    Exit Sub

End Sub

Public Sub Novo_GOBJ(objTribAux As Object)
    Set objTribAux = New ClassPRJContratos
End Sub

Public Sub Atualiza_GOBJ(objTribAux As Object)
    Set gobjContrato = objTribAux
End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProjeto = New AdmEvento
    Set objEventoProposta = New AdmEvento
    Set objEventoPRJ = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEtapa = New AdmEvento
    Set objEventoNaturezaOp = New AdmEvento
    Set objEventoTiposDeTributacao = New AdmEvento

    Call gobjTelaCamposCust.Exibe_Campos_Customizados
    
    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    giDataReferenciaAlterada = 0
    
    Set objGridItens = New AdmGrid
    
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 187718
    
    Set objGridEtapa = New AdmGrid
    
    lErro = Inicializa_GridEtapa(objGridEtapa)
    If lErro <> SUCESSO Then gError 187719
    
    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 187720
    
    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError 189066
             
    lErro = Inicializa_Mascara_Projeto(PRJ)
    If lErro <> SUCESSO Then gError 189066
             
'    lErro = TributacaoPRJCTR_Reset()
'    If lErro <> SUCESSO Then gError 187721
'
'    Call BotaoGravarTribCarga
'
'    lErro = CarregaTiposTrib()
'    If lErro <> SUCESSO Then gError 187722

    Set gobjTribTab = New ClassTribTab
    lErro = gobjTribTab.Ativar(Me, , , gobjTribTab.TIPOTELA_PRJ_CTR)
    If lErro <> SUCESSO Then gError 187722
    
    Set gobjContrato = New ClassPRJContratos
    lErro = gobjTribTab.TributacaoNF_Reset(gobjContrato)
    If lErro <> SUCESSO Then gError 187722
    
    iListIndexDefault = Drive1.ListIndex
    
    If Len(Trim(CurDir)) > 0 Then
        Dir1.Path = CurDir
        Drive1.Drive = left(CurDir, 2)
    End If
    
    NomeDiretorio.Text = Dir1.Path

    Call ValorTotal_Calcula

    iAlterado = 0
    iFrameAtual = 1
    iFrameAtualEtapa = 1
    'iFrameAtualTributacao = 1
    
    iDescontoAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 187718 To 187722, 189066

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187607)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objContrato As ClassPRJContratos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objContrato Is Nothing) Then

        lErro = Traz_Contrato_Tela(objContrato)
        If lErro <> SUCESSO Then gError 187723

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 187723

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187608)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objContratoBD As New ClassPRJContratos

On Error GoTo Erro_Move_Tela_Memoria

    objContrato.sCodigo = Codigo.Text
    objContrato.iFilialEmpresa = giFilialEmpresa
    objContrato.dtData = StrParaDate(DataCriacao.Text)
       
    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 187724

        If lErro = SUCESSO Then
            objContrato.lCliente = objCliente.lCodigo
        End If
            
    End If

    objContrato.iFilialCliente = Codigo_Extrai(Filial.Text)
    objContrato.sObservacao = Obs.Text
    
    objContrato.dCustoCalculado = StrParaDbl(TotalCustoCalc.Caption)
    objContrato.dCustoInformado = StrParaDbl(TotalCustoInf.Caption)
    objContrato.dValorFrete = StrParaDbl(ValorFrete.Text)
    objContrato.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objContrato.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objContrato.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objContrato.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objContrato.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objContrato.sNaturezaOp = gobjTribTab.sNatOpInterna
    
    If Not (gobjProjeto Is Nothing) Then
        objContrato.lNumIntDocPRJ = gobjProjeto.lNumIntDoc
    End If
    
    'Lê o Projetos que está sendo Passado
    lErro = CF("PRJContratos_Le", objContratoBD)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189399

    'Se já estava gravado não considera se a proposta estiver preenchida
    If lErro = SUCESSO Then
        objContrato.lNumIntDocProposta = objContratoBD.lNumIntDocProposta
    Else
        If Not (gobjProposta Is Nothing) Then
            objContrato.lNumIntDocProposta = gobjProposta.lNumIntDoc
        End If
    End If
    
    If ExibirCustoCalc.Value = vbChecked Then
        objContrato.iExibirCustoCalc = MARCADO
    End If
    If ExibirCustoInf.Value = vbChecked Then
        objContrato.iExibirCustoInfo = MARCADO
    End If
    If ExibirPreco.Value = vbChecked Then
        objContrato.iExibirPreco = MARCADO
    End If
    If ExibirProdutos.Value = vbChecked Then
        objContrato.iExibirProdutos = MARCADO
    End If
    
    'Move Tributacao para objContrato
    Set objContrato.objTributacaoPRJCTR = gobjContrato.objTributacaoPRJCTR
    
    lErro = gobjTelaCamposCust.Move_Tela_Memoria(objContrato.objCamposCust, objContrato.objTiposCamposCust)
    If lErro <> SUCESSO Then gError 187725
    
    lErro = Move_Itens_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 187726
    
    lErro = Move_Etapa_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 187727
    
    objContrato.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objContrato.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
 
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 187724 To 187727

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187609)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objContrato As New ClassPRJContratos

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PRJContratos"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 187728

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDocPRJ", objContrato.lNumIntDocPRJ, 0, "NumIntDocPRJ"
    colCampoValor.Add "Codigo", objContrato.sCodigo, STRING_PRJ_CODIGO, "Codigo"
    'Filtros para o Sistema de Setas
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 187728

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187610)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objContrato As New ClassPRJContratos

On Error GoTo Erro_Tela_Preenche

    objContrato.lNumIntDocPRJ = colCampoValor.Item("NumIntDocPRJ").vValor
    objContrato.sCodigo = colCampoValor.Item("Codigo").vValor
    objContrato.iFilialEmpresa = giFilialEmpresa

    If objContrato.sCodigo <> "" Then
        lErro = Traz_Contrato_Tela(objContrato)
        If lErro <> SUCESSO Then gError 187729
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 187729

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187611)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objContrato As New ClassPRJContratos

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    Call Recolhe_Observacao(iLinhaAnt)

    If Len(Trim(Codigo.Text)) = 0 Then gError 187730
    If Len(Trim(Projeto.ClipText)) = 0 Then gError 187731
    If Len(Trim(DataCriacao.ClipText)) = 0 Then gError 187732
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 189376
    If Len(Trim(Filial.Text)) = 0 Then gError 189377

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 187733
    
    If objContrato.colItens.Count = 0 Then gError 187734
    
    lErro = gobjTribTab.Valida_Dados()
    If lErro <> SUCESSO Then gError 187735

    lErro = Trata_Alteracao(objContrato, objContrato.sCodigo, objContrato.lNumIntDocPRJ)
    If lErro <> SUCESSO Then gError 187736

    'Grava o/a Projetos no Banco de Dados
    lErro = CF("PRJContratos_Grava", objContrato)
    If lErro <> SUCESSO Then gError 187737

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 187730
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJCONTRATO_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus
            
        Case 187731
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Projeto.SetFocus
            
        Case 187732
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
'            DataCriacao.SetFocus
        
        Case 187733, 187735, 187736, 187737
        
        Case 187734
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)

        Case 189376
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
'            Cliente.SetFocus

        Case 189377
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)
'            Filial.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187612)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Contrato() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Contrato

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    sProjetoAnt = ""
    sNomeProjetoAnt = ""
    
    Set gobjProjeto = Nothing
    Set gobjEtapa = Nothing
    Set gobjEtapaIP = Nothing
    Set gobjProposta = Nothing
    
    TvwEtapas.Nodes.Clear
    
    Call Grid_Limpa(objGridItens)
    
    iLinhaAnt = 1
    
    Call Grid_Limpa(objGridEtapa)
    
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    ExibirProdutos.Value = vbUnchecked
    ExibirPreco.Value = vbUnchecked
    ExibirCustoInf.Value = vbUnchecked
    ExibirCustoCalc.Value = vbUnchecked
    
    CustoCalculado.Caption = ""
    DescricaoEtapa.Caption = ""
    TotalCustoInf.Caption = ""
    TotalPreco.Caption = ""
    TotalCustoCalc.Caption = ""
    
    If Len(Trim(CurDir)) > 0 Then
        Dir1.Path = CurDir
        Drive1.Drive = left(CurDir, 2)
    End If
    
    NomeDiretorio.Text = Dir1.Path
    
    Filial.Clear

'    'tab de tributacao resumo
'    'ISSIncluso.Value = 0
'    IPIBase.Caption = ""
'    IPIValor.Caption = ""
'    ISSBase.Caption = ""
'    DescTipoTrib.Caption = ""
'    IRBase.Caption = ""
'    ICMSBase.Caption = ""
'    ICMSValor.Caption = ""
'    ICMSSubstBase.Caption = ""
'    ICMSSubstValor.Caption = ""
'
'    'tab de tributacao itens
'    LabelValorFrete.Caption = ""
'    LabelValorDesconto.Caption = ""
'    LabelValorSeguro.Caption = ""
'    LabelValorOutrasDespesas.Caption = ""
'    ComboItensTrib.Clear
'    LabelValorItem.Caption = ""
'    LabelQtdeItem.Caption = ""
'    LabelUMItem.Caption = ""
'    LabelDescrNatOpItem.Caption = ""
'    DescTipoTribItem.Caption = ""
    
    ValorTotal.Caption = ""
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
    
'    'Resseta tributação
'    Call TributacaoPRJCTR_Reset
'
'    Call BotaoGravarTrib

    iAlterado = 0
    giValorDescontoAlterado = 0
    giClienteAlterado = 0
    giFilialAlterada = 0
    giDataReferenciaAlterada = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0
    giValorDescontoManual = 0

    Call gobjTribTab.Limpa_Tela

    Limpa_Tela_Contrato = SUCESSO

    Exit Function

Erro_Limpa_Tela_Contrato:

    Limpa_Tela_Contrato = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187613)

    End Select

    Exit Function

End Function

Function Traz_Contrato_Tela(ByVal objContrato As ClassPRJContratos, Optional ByVal bNaoLe As Boolean = False) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim objProposta As New ClassPRJPropostas
Dim sTextoProjeto As String
Dim sTextoProposta As String
Dim iIndice As Integer, dPercDesc As Double

On Error GoTo Erro_Traz_Contrato_Tela

    gbCarregandoTela = True

    Call Limpa_Tela_Contrato
    
    If Not bNaoLe Then
    
        'Lê o Projetos que está sendo Passado
        lErro = CF("PRJContratos_Le", objContrato, True, True)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187738

    End If
    
    Codigo.Text = objContrato.sCodigo

    If lErro = SUCESSO Then
        
        'lErro = TributacaoPRJCTR_Reset(objContrato)
        lErro = gobjTribTab.Traz_NFiscal_Tela(objContrato)
        If lErro <> SUCESSO Then gError 187739
        
        ValorTotal.Caption = Format(objContrato.dValorTotal, "Standard")
        ValorProdutos.Caption = Format(objContrato.dValorProdutos, "Standard")
        
        If objContrato.dtData <> DATA_NULA Then
            DataCriacao.PromptInclude = False
            DataCriacao.Text = Format(objContrato.dtData, "dd/mm/yy")
            DataCriacao.PromptInclude = True
        End If
       
        objProjeto.lNumIntDoc = objContrato.lNumIntDocPRJ
        
        'Lê o Projetos que está sendo Passado
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187740
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189124
        
        Call Projeto_Validate(bSGECancelDummy)
        
        lErro = Retorno_Projeto_Tela(PRJ, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189124
        
        Call PRJ_Validate(bSGECancelDummy)
        
        If objContrato.lNumIntDocProposta <> 0 Then
        
            objProposta.lNumIntDoc = objContrato.lNumIntDocProposta
        
            'Lê a Proposta que está sendo Passado
            lErro = CF("PRJPropostas_Le_NumIntDoc", objProposta)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189397
            
            Proposta.Text = objProposta.sCodigo
            Call Proposta_Validate(bSGECancelDummy)
            
        End If
    
        'Se existe um código para o Cliente
        If objContrato.lCliente <> 0 Then
        
            Call Cliente_Formata(objContrato.lCliente)
            Call Filial_Formata(Filial, objContrato.iFilialCliente)
        End If
        
        Obs.Text = objContrato.sObservacao
     
        objContrato.objCamposCust.iTipoNumIntDocOrigem = CAMPO_CUSTOMIZADO_TIPO_CONTRATO
        objContrato.objCamposCust.lNumIntDocOrigem = objContrato.lNumIntDoc

        lErro = gobjTelaCamposCust.Traz_CamposCustomizados_Tela(objContrato.objCamposCust)
        If lErro <> SUCESSO Then gError 187741
        
        If objContrato.iExibirProdutos Then
            ExibirProdutos.Value = vbChecked
        End If
        If objContrato.iExibirPreco Then
            ExibirPreco.Value = vbChecked
        End If
        If objContrato.iExibirCustoInfo Then
            ExibirCustoInf.Value = vbChecked
        End If
        If objContrato.iExibirCustoCalc Then
            ExibirCustoCalc.Value = vbChecked
        End If
               
'        NatOpEspelho.Caption = objContrato.sNaturezaOp
        
        ValorFrete.Text = Format(objContrato.dValorFrete, "Standard")
        ValorSeguro.Text = Format(objContrato.dValorSeguro, "Standard")
        ValorDesconto.Text = Format(objContrato.dValorDesconto, "Standard")
        ValorDespesas.Text = Format(objContrato.dValorOutrasDespesas, "Standard")
    
'        NaturezaOp.Text = objContrato.sNaturezaOp
'        Call NaturezaOp_Validate(bSGECancelDummy)
        
        giValorFreteAlterado = 0
        giValorSeguroAlterado = 0
        giValorDescontoAlterado = 0
        giValorDespesasAlterado = 0
        
        lErro = Traz_Itens_Tela(objContrato)
        If lErro <> SUCESSO Then gError 187742
        
        lErro = Traz_Etapa_Tela(objContrato)
        If lErro <> SUCESSO Then gError 187743
                
'        'Carrega o Tab de Tributação
'        lErro = Carrega_Tab_Tributacao(objContrato)
'        If lErro <> SUCESSO Then gError 187744
            
        ValorTotal.Caption = Format(objContrato.dValorTotal, "Standard")
       
    Else
    
        objProjeto.lNumIntDoc = objContrato.lNumIntDocPRJ
        
        'Lê o Projetos que está sendo Passado
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187740
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189124
        
        Call Projeto_Validate(bSGECancelDummy)
        
        lErro = Retorno_Projeto_Tela(PRJ, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189124
        
        Call PRJ_Validate(bSGECancelDummy)

    End If
    
    If objContrato.dValorItens = 0 Then
        For iIndice = 1 To objGridItens.iLinhasExistentes
            objContrato.dValorItens = objContrato.dValorItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            objContrato.dValorDescontoItens = objContrato.dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        Next
    End If
    
    ValorProdutos2.Caption = Format(objContrato.dValorItens, "Standard")
    ValorDescontoItens.Text = Format(objContrato.dValorDescontoItens, "Standard")
    If objContrato.dValorItens > 0 Then
        dPercDesc = objContrato.dValorDescontoItens / objContrato.dValorItens
    Else
        dPercDesc = 0
    End If
    PercDescontoItens.Text = Format(dPercDesc * 100, "Fixed")
    
    dValorDescontoItensAnt = objContrato.dValorDescontoItens
    dPercDescontoItensAnt = dPercDesc
    
    Call Calcula_Total_Tvw
       
    iAlterado = 0
    
    gbCarregandoTela = False

    Traz_Contrato_Tela = SUCESSO

    Exit Function

Erro_Traz_Contrato_Tela:

    Traz_Contrato_Tela = gErr

    Select Case gErr

        Case 187738 To 187744, 189124, 189397

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187614)

    End Select
    
    gbCarregandoTela = False

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 187745

    'Limpa Tela
    Call Limpa_Tela_Contrato

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 187745

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187615)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187616)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 187746

    Call Limpa_Tela_Contrato

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 187746

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187617)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objContrato As New ClassPRJContratos
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 187747
    If Len(Trim(Projeto.ClipText)) = 0 Then gError 187748
    
    objContrato.sCodigo = Codigo.Text
    objContrato.lNumIntDocPRJ = gobjProjeto.lNumIntDoc
    objContrato.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CONTRATO")

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("PRJContratos_Exclui", objContrato)
        If lErro <> SUCESSO Then gError 187749

        'Limpa Tela
        Call Limpa_Tela_Contrato

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 187747
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJCONTRATO_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus
            
        Case 187748
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Projeto.SetFocus
            
        Case 187749

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187618)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187619)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataCriacao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_DownClick

    DataCriacao.SetFocus

    If Len(DataCriacao.ClipText) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 187750
        
        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_DownClick:

    Select Case gErr

        Case 187750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187620)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCriacao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_UpClick

    DataCriacao.SetFocus

    If Len(Trim(DataCriacao.ClipText)) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 187751

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_UpClick:

    Select Case gErr

        Case 187751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187621)

    End Select

    Exit Sub

End Sub

Private Sub DataCriacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCriacao, iAlterado)
    
End Sub

Private Sub DataCriacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCriacao_Validate

    If Len(Trim(DataCriacao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataCriacao.Text)
        If lErro <> SUCESSO Then gError 187752

        If gobjContrato.dtData <> StrParaDate(DataCriacao.Text) Then
            
            gobjContrato.dtData = StrParaDate(DataCriacao.Text)
            
            Call ValorTotal_Calcula
            
        End If
    
    End If

    Exit Sub

Erro_DataCriacao_Validate:

    Cancel = True

    Select Case gErr

        Case 187752

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187622)

    End Select

    Exit Sub

End Sub

Private Sub DataCriacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Cliente_Validate

    If giClienteAlterado = 1 Then
    
        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then
    
            'Busca o Cliente no BD
            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 187753
            
            'gobjContrato.lCliente = objCliente.lCodigo
                       
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 187754
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
            
            If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
    
                If Not gbCarregandoTela Then
    
                    If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
    
                        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
    
                        'Seleciona filial na Combo Filial
                        Call CF("Filial_Seleciona", Filial, iCodFilial)
    
                    End If
                    
                End If
                
                'Se o Tipo estiver preenchido
                If objCliente.iTipo > 0 Then
                
                    objTipoCliente.iCodigo = objCliente.iTipo
                    
                    'Lê o Tipo de Cliente
                    lErro = CF("TipoCliente_Le", objTipoCliente)
                    If lErro <> SUCESSO And lErro <> 19062 Then gError 187755
                    
                End If

                giValorDescontoManual = 0
                
                'Guarda o valor do desconto do cliente
                If objCliente.dDesconto > 0 Then
                    
                    gdDesconto = objCliente.dDesconto
                
                ElseIf objTipoCliente.dDesconto > 0 Then
                    
                    gdDesconto = objTipoCliente.dDesconto
                
                Else
                    
                    gdDesconto = 0
                
                End If
    
                If Not gbCarregandoTela Then
    
                    Call DescontoGlobal_Recalcula
    
                    'ATualiza o total com o novo desconto
                    lErro = ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 187756
   
                End If
                
                giClienteAlterado = 0
    
        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then
    
                gobjContrato.lCliente = 0
                giValorDescontoManual = 0
                gdDesconto = 0
                                
                If Not gbCarregandoTela Then
                
                    Call DescontoGlobal_Recalcula
    
                    'ATualiza o total com o novo desconto
                    lErro = ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 187757
                    
                    objCliente.lCodigo = 0
                    
                    Filial.Clear
                            
                End If
    
        End If
    
        If Not gbCarregandoTela Then
    
    ''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
    '*** fim tributacao
        End If
    
    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 187753 To 187757

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187623)

    End Select

    Exit Sub

End Sub

Private Sub DescontoGlobal_Recalcula()

Dim dValorDesconto As Double
Dim dValorProdutos As Double

    If gbCarregandoTela Then Exit Sub
    
    PercDescontoItens.Text = Format(gdDesconto * 100, "FIXED")
    Call PercDescontoItens_Validate(bSGECancelDummy)
'
'    If Len(Trim(ValorProdutos.Caption)) <> 0 And IsNumeric(ValorProdutos.Caption) Then
'
'        'Se o cliente possui desconto e o campo desconto não foi alterado pelo usuário
'        If gdDesconto > 0 And giValorDescontoManual = 0 Then
'
'            Call Calcula_ValorProdutos(dValorProdutos)
'
'            'Calcula o valor do desconto para o cliente e coloca na tela
'            dValorDesconto = gdDesconto * dValorProdutos
'            ValorDesconto.Text = Format(dValorDesconto, "Standard")
'            giValorDescontoAlterado = 0
'
'            'Para tributação
'            gobjContrato.dValorDesconto = dValorDesconto
'
'        End If
'
'    End If

End Sub

Public Sub Calcula_ValorProdutos(dValorProdutos As Double)

Dim dValorTotal As Double
Dim dValor As Double
Dim iIndice As Integer

    dValor = 0

    For iIndice = 1 To objGridItens.iLinhasExistentes

        dValorTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

        dValor = dValor + dValorTotal

    Next

    dValorProdutos = dValor

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
    giClienteAlterado = 1
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 187758

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 187759

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 187760

        If lErro = 17660 Then

            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 187761

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objCliente.lCodigo

                gError 187762
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
            lErro = Trata_FilialCliente
            If lErro <> SUCESSO Then gError 187763
        
        Else
            
            objCliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0
            
        End If
        
    'Não encontrou a STRING
    ElseIf lErro = 6731 Then
        
        'trecho incluido por Leo em 17/04/02
        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 187764
        
        If lErro = SUCESSO Then gError 187765
        
    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 187758, 187760, 187763

        Case 187759
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 187761, 187764 'tratado na rotina chamada

        Case 187762
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 187765
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187624)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    'Se nenhuma filial foi selecionada, sai.
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento para a filial do cliente selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 187766

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 187766

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187625)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Function Trata_FilialCliente() As Long

Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim objTipoCliente As New ClassTipoCliente
Dim dValorTotal As Double
Dim dValorBase As Double
Dim objTransportadora As New ClassTransportadora
Dim dValorComissao As Double
Dim dValorEmissao As Double
Dim lErro As Long

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 187767
    If lErro = 17660 Then gError 187768

    gobjContrato.iFilialCliente = objFilialCliente.iCodFilial
    
    Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

    'Calula o valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 187769
    
    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 187767, 187769

        Case 187768
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187626)

    End Select

    Exit Function

End Function

Private Sub Obs_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Obs_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Obs_Validate

    'Verifica se Obs está preenchida
    If Len(Trim(Obs.Text)) <> 0 Then

       '#######################################
       'CRITICA Obs
       '#######################################

    End If

    Exit Sub

Erro_Obs_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187627)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objContrato As New ClassPRJContratos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objContrato.sCodigo = Codigo.Text

    End If

    Call Chama_Tela("PRJContratosLista", colSelecao, objContrato, objEventoCodigo, , "Código")

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187628)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objContrato As New ClassPRJContratos

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objContrato = obj1
    
    lErro = Traz_Contrato_Tela(objContrato)
    If lErro <> SUCESSO Then gError 187770
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 187770

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187629)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)


End Sub

Private Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de tributacao
        If Opcao.SelectedItem.Index = TAB_TRIBUTACAO Then
            
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError 187771

        End If
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr

        Case 187771

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187630)

    End Select

    Exit Sub

End Sub

Private Sub OpcaoEtapa_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If OpcaoEtapa.SelectedItem.Index <> iFrameAtualEtapa Then

        If TabStrip_PodeTrocarTab(iFrameAtualEtapa, OpcaoEtapa, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameE(OpcaoEtapa.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameE(iFrameAtualEtapa).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtualEtapa = OpcaoEtapa.SelectedItem.Index
        
    End If

End Sub

'##################################################################
'Tem que colocar o código para o modo de edição aqui
Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelProjeto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProjeto, Source, X, Y)
End Sub

Private Sub LabelProjeto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProjeto, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeRedPRJ_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeRedPRJ, Source, X, Y)
End Sub

Private Sub LabelNomeRedPRJ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeRedPRJ, Button, Shift, X, Y)
End Sub
'##################################################################

'##################################################################
'Tratamento dos Campos customizados
Private Sub Data_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Data_GotFocus(Index)
End Sub

Private Sub Data_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Data_Validate(Index, Cancel)
End Sub

Private Sub Numero_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Numero_GotFocus(Index)
End Sub

Private Sub Numero_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Numero_Validate(Index, Cancel)
End Sub

Private Sub Valor_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Valor_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Valor_GotFocus(Index)
End Sub

Private Sub Valor_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Valor_Validate(Index, Cancel)
End Sub

Private Sub Texto_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownData_DownClick(Index As Integer)
    Call gobjTelaCamposCust.UpDownData_DownClick(Index)
End Sub

Private Sub UpDownData_UpClick(Index As Integer)
    Call gobjTelaCamposCust.UpDownData_UpClick(Index)
End Sub

Private Sub BotaoDadosCustNovo_Click()
    Call gobjTelaCamposCust.BotaoDadosCustNovo_Click
End Sub

Private Sub BotaoDadosCustDel_Click()
    Call gobjTelaCamposCust.BotaoDadosCustDel_Click
End Sub
'##################################################################

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição Produto")
    objGridInt.colColuna.Add ("Etapa")
    objGridInt.colColuna.Add ("Descrição Etapa")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (Etapa.Name)
    objGridInt.colCampo.Add (DescEtapa.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_Etapa_Col = 3
    iGrid_DescEtapa_Col = 4
    iGrid_UnidadeMed_Col = 5
    iGrid_Quantidade_Col = 6
    iGrid_PrecoUnitario_Col = 7
    iGrid_PercDesc_Col = 8
    iGrid_Desconto_Col = 9
    iGrid_DataEntrega_Col = 10
    iGrid_PrecoTotal_Col = 11
    iGrid_PrecoTotalB_Col = 12
    iGrid_Observacao_Col = 13

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Sub Projeto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Projeto_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Then

        If Len(Trim(Projeto.ClipText)) > 0 Then
            
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 189087
        
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187772
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187773
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
            NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
            
        End If
        
        sProjetoAnt = Projeto.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 187774
        
    End If
   
    Exit Sub

Erro_Projeto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 187772, 187774, 189087
        
        Case 187773
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 187631)

    End Select

    Exit Sub

End Sub

Sub NomeReduzidoPrj_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long

On Error GoTo Erro_NomeReduzidoPrj_Validate

    'Se alterou o projeto
    If sNomeProjetoAnt <> NomeReduzidoPRJ.Text Then

        If Len(Trim(NomeReduzidoPRJ.Text)) > 0 Then
            
            objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le_NomeReduzido", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187775
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187776
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
            lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
            If lErro <> SUCESSO Then gError 189125
            
        End If
        
        sNomeProjetoAnt = NomeReduzidoPRJ.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 187777
        
    End If
    
    Exit Sub

Erro_NomeReduzidoPrj_Validate:

    Cancel = True

    Select Case gErr
    
        Case 187775, 187777, 189125
        
        Case 187776
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO3", gErr, objProjeto.sNomeReduzido, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 187632)

    End Select

    Exit Sub

End Sub

Sub LabelProjeto_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjeto_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Projeto.ClipText)) <> 0 Then

        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189088

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoProjeto, , "Código")

    Exit Sub

Erro_LabelProjeto_Click:

    Select Case gErr
    
        Case 189088

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187633)

    End Select

    Exit Sub
    
End Sub

Sub LabelNomeRedPRJ_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeRedPRJ_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(NomeReduzidoPRJ.Text)) <> 0 Then

        objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoProjeto, , "Nome Reduzido")

    Exit Sub

Erro_LabelNomeRedPRJ_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187634)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoProjeto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoProjeto_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189126
    
    NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
    
    Call Projeto_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoProjeto_evSelecao:

    Select Case gErr
    
        Case 189126

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187635)

    End Select

    Exit Sub

End Sub

Function Trata_Projeto(ByVal lNumIntDocPRJ As Long) As Long

Dim lErro As Long
Dim objProjeto As ClassProjetos
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_Trata_Projeto
   
    If lNumIntDocPRJ <> 0 Then
    
        Set objProjeto = New ClassProjetos

        objProjeto.lNumIntDoc = lNumIntDocPRJ
        
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187966
    
        lErro = CF("PRJEtapas_Le_Projeto", objProjeto, True)
        If lErro <> SUCESSO Then gError 187778
        
        objProjeto.objEscopo.lNumIntDoc = objProjeto.lNumIntDocEscopo
        lErro = CF("PRJEscopo_Le", objProjeto.objEscopo)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        For Each objEtapa In objProjeto.colEtapas
        
            lErro = objEtapa.Obtem_Custo
            If lErro <> SUCESSO Then gError 187781
            
        Next
       
    End If

    Set gobjProjeto = objProjeto

    lErro = Carrega_Arvore(objProjeto)
    If lErro <> SUCESSO Then gError 187779
    
    lErro = Preenche_Grid_Etapa(objProjeto)
    If lErro <> SUCESSO Then gError 187780
    
    sProjetoAnt = Projeto.Text
    sNomeProjetoAnt = NomeReduzidoPRJ.Text
    
    Trata_Projeto = SUCESSO

    Exit Function

Erro_Trata_Projeto:

    Trata_Projeto = gErr

    Select Case gErr
    
        Case 187778 To 187781, 187966, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187636)

    End Select

    Exit Function

End Function

Function Carrega_Arvore(ByVal objProjeto As ClassProjetos) As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim objNode As Node
Dim lErro As Long
Dim sChaveTvw As String
Dim iIndicePai As Integer
Dim sTexto As String
Dim objEtapa As ClassPRJEtapas
Dim objEtapaAux As ClassPRJEtapas
Dim iProxChave As Integer
Dim objEtapaIP As ClassPRJEtapaItensProd
Dim sTextoAux As String
Dim sProdutoMascarado As String

On Error GoTo Erro_Carrega_Arvore

    TvwEtapas.Nodes.Clear

    If Not (objProjeto Is Nothing) Then
    
        For Each objEtapa In objProjeto.colEtapas
    
            'Texto que identificará a nova Etapa que está sendo incluida
            sTexto = objEtapa.sNomeReduzido
            sTextoAux = ""
            
            If ExibirCustoCalc.Value = vbChecked Then
                sTextoAux = "Custo Calc: " & Format(objEtapa.dCustoCalcPrev, "STANDARD")
            End If
            
            If ExibirCustoInf.Value = vbChecked Then
                If Len(Trim(sTextoAux)) <> 0 Then
                    sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Custo Inf: " & Format(objEtapa.dCustoInfoPrev, "STANDARD")
                Else
                    sTextoAux = "Custo Inf: " & Format(objEtapa.dCustoInfoPrev, "STANDARD")
                End If
            End If
                       
            If ExibirPreco.Value = vbChecked Then
                If Len(Trim(sTextoAux)) <> 0 Then
                    sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Preço: " & Format(objEtapa.dPreco, "STANDARD")
                Else
                    sTextoAux = "Preço: " & Format(objEtapa.dPreco, "STANDARD")
                End If
            End If
            
            If Len(Trim(sTextoAux)) > 0 Then
                sTexto = sTexto & " (" & sTextoAux & ")"
            End If
            
            'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
            Call Calcula_Proxima_Chave(iProxChave)
    
            sChaveTvw = "X" & CStr(iProxChave)
    
            If objEtapa.lNumIntDocEtapaPaiOrg = 0 Then
    
                Set objNode = TvwEtapas.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)
    
            Else
    
                For Each objEtapaAux In objProjeto.colEtapas
                
                    If objEtapa.lNumIntDocEtapaPaiOrg = objEtapaAux.lNumIntDoc Then
                        iIndicePai = objEtapaAux.iIndiceTvw
                        Exit For
                    End If
    
                Next
    
                Set objNode = TvwEtapas.Nodes.Add(iIndicePai, tvwChild, sChaveTvw, sTexto)
    
            End If
                    
            TvwEtapas.Nodes.Item(objNode.Index).Expanded = True
            
            objEtapa.sCodigoAnt = objEtapa.sCodigo
            objEtapa.iIndiceTvw = objNode.Index
            objEtapa.sChaveTvw = sChaveTvw
            
            objNode.Checked = objEtapa.iTvwChecked
                       
            objNode.Tag = sChaveTvw
            
            If ExibirProdutos.Value = vbChecked Then
                        
                For Each objEtapaIP In objEtapa.colItensProduzidos
                
                    lErro = Mascara_RetornaProdutoTela(objEtapaIP.sProduto, sProdutoMascarado)
                    If lErro <> SUCESSO Then gError 187782
                
                    'Texto que identificará a nova Etapa que está sendo incluida
                    sTexto = sProdutoMascarado & SEPARADOR & objEtapaIP.sDescricao
                    sTextoAux = ""
                    
                    If ExibirCustoCalc.Value = vbChecked Then
                        sTextoAux = "Custo Calc: " & Format(0, "STANDARD")
                    End If
                    
                    If ExibirCustoInf.Value = vbChecked Then
                        If Len(Trim(sTextoAux)) <> 0 Then
                            sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Custo Inf: " & Format(objEtapaIP.dCustoInfo, "STANDARD")
                        Else
                            sTextoAux = "Custo Inf: " & Format(objEtapaIP.dCustoInfo, "STANDARD")
                        End If
                    End If
                               
                    If ExibirPreco.Value = vbChecked Then
                        If Len(Trim(sTextoAux)) <> 0 Then
                            sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Preço: " & Format(objEtapaIP.dPreco, "STANDARD")
                        Else
                            sTextoAux = "Preço: " & Format(objEtapaIP.dPreco, "STANDARD")
                        End If
                    End If
                    
                    If Len(Trim(sTextoAux)) > 0 Then
                        sTexto = sTexto & " (" & sTextoAux & ")"
                    End If
                    
                    'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
                    Call Calcula_Proxima_Chave(iProxChave)
            
                    sChaveTvw = "X" & CStr(iProxChave)

                    Set objNode = TvwEtapas.Nodes.Add(objEtapa.iIndiceTvw, tvwChild, sChaveTvw, sTexto)
                
                    objNode.Tag = objEtapa.sChaveTvw
                    
                    objNode.Checked = objEtapaIP.iTvwChecked
                
                    objEtapaIP.iIndiceTvw = objNode.Index
                    objEtapaIP.sChaveTvw = sChaveTvw
                
                Next
            
            End If
            
        Next
        
        Call Calcula_Total_Tvw
        
    End If
    
    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr
    
        Case 187782
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187637)

    End Select

    Exit Function

End Function

Function Preenche_Grid_Etapa(ByVal objProjeto As ClassProjetos) As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim lErro As Long
Dim objEtapa As ClassPRJEtapas
Dim sProdutoMascarado As String
Dim iLinha As Integer

On Error GoTo Erro_Preenche_Grid_Etapa

    If Not (objProjeto Is Nothing) Then
    
        For Each objEtapa In objProjeto.colEtapas
            
            iLinha = iLinha + 1
            GridEtapa.TextMatrix(iLinha, iGrid_Imprimir_Col) = MARCADO
            GridEtapa.TextMatrix(iLinha, iGrid_EtapaGrid_Col) = objEtapa.sCodigo
            GridEtapa.TextMatrix(iLinha, iGrid_DescricaoGrid_Col) = objEtapa.sDescricao
            GridEtapa.TextMatrix(iLinha, iGrid_ObservacaoGrid_Col) = ""
        Next
        
    End If
    
    Call Grid_Refresh_Checkbox(objGridEtapa)
    
    objGridEtapa.iLinhasExistentes = iLinha
    
    Preenche_Grid_Etapa = SUCESSO

    Exit Function

Erro_Preenche_Grid_Etapa:

    Preenche_Grid_Etapa = gErr

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187638)

    End Select

    Exit Function

End Function

Private Sub Calcula_Proxima_Chave(iProxChave As Integer)

Dim sChave As String
Dim objNode1 As Node
Dim iAtual As Integer
Dim lErro As Long

On Error GoTo Erro_Calcula_Proxima_Chave

    iProxChave = 0

    For Each objNode1 In TvwEtapas.Nodes

        iAtual = StrParaInt(right(objNode1.Key, Len(objNode1.Key) - 1))

        If iAtual > iProxChave Then iProxChave = iAtual

    Next

     iProxChave = iProxChave + 1

     Exit Sub

Erro_Calcula_Proxima_Chave:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187639)

    End Select

    Exit Sub

End Sub

Private Sub ExibirCustoCalc_Click()
    Call Carrega_Arvore(gobjProjeto)
End Sub

Private Sub ExibirCustoInf_Click()
    Call Carrega_Arvore(gobjProjeto)
End Sub

Private Sub ExibirPreco_Click()
    Call Carrega_Arvore(gobjProjeto)
End Sub

Private Sub ExibirProdutos_Click()
    Call Carrega_Arvore(gobjProjeto)
End Sub

Private Sub TvwEtapas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objEtapa As ClassPRJEtapas
Dim objEtapaIP As ClassPRJEtapaItensProd

On Error GoTo Erro_TvwEtapas_NodeClick

    For Each objEtapa In gobjProjeto.colEtapas
        If objEtapa.sChaveTvw = Node.Tag Then
            Exit For
        End If
    Next

    Set gobjEtapa = objEtapa
    
    If objEtapa.iIndiceTvw = Node.Index Then

        DescricaoEtapa.Caption = gobjEtapa.sDescricao
        Preco.Text = Format(gobjEtapa.dPreco, "STANDARD")
        CustoInformado.Text = Format(gobjEtapa.dCustoInfoPrev, "STANDARD")
        CustoCalculado.Caption = Format(gobjEtapa.dCustoCalcPrev, "STANDARD")
        
        Set gobjEtapaIP = Nothing
        
    Else
    
        For Each objEtapaIP In objEtapa.colItensProduzidos
            If objEtapaIP.iIndiceTvw = Node.Index Then
                Exit For
            End If
        Next
        
        Set gobjEtapaIP = objEtapaIP
    
        DescricaoEtapa.Caption = gobjEtapa.sDescricao & " ( " & objEtapaIP.sProduto & SEPARADOR & objEtapaIP.sDescricao & ")"
        Preco.Text = Format(objEtapaIP.dPreco, "STANDARD")
        CustoInformado.Text = Format(objEtapaIP.dCustoInfo, "STANDARD")
        CustoCalculado.Caption = Format(0, "STANDARD")
    
    End If
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_TvwEtapas_NodeClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187640)

    End Select

    Exit Sub

End Sub

Private Sub CustoInformado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Preco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CustoInformado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CustoInformado_Validate

    'Veifica se CustoInformado está preenchida
    If Len(Trim(CustoInformado.Text)) <> 0 Then

       'Critica a CustoInformado
       lErro = Valor_Positivo_Critica(CustoInformado.Text)
       If lErro <> SUCESSO Then gError 187784
        
    End If

    If Not (gobjEtapa Is Nothing) Then
        If gobjEtapaIP Is Nothing Then
            gobjEtapa.dCustoInfoPrev = StrParaDbl(CustoInformado.Text)
            Call Acerta_Texto_Node(gobjEtapa)
        Else
            gobjEtapaIP.dCustoInfo = StrParaDbl(CustoInformado.Text)
            Call Acerta_Texto_Node_IP(gobjEtapaIP)
        End If
    End If


    Call Calcula_Total_Tvw

    Exit Sub

Erro_CustoInformado_Validate:

    Cancel = True

    Select Case gErr

        Case 187784

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187641)

    End Select

    Exit Sub

End Sub

Private Sub Preco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Preco_Validate

    'Veifica se Preco está preenchida
    If Len(Trim(Preco.Text)) <> 0 Then

        'Critica a Preco
        lErro = Valor_Positivo_Critica(Preco.Text)
        If lErro <> SUCESSO Then gError 187785
       
    End If

    If Not (gobjEtapa Is Nothing) Then
        If gobjEtapaIP Is Nothing Then
            gobjEtapa.dPreco = StrParaDbl(Preco.Text)
            Call Acerta_Texto_Node(gobjEtapa)
        Else
            gobjEtapaIP.dPreco = StrParaDbl(Preco.Text)
            Call Acerta_Texto_Node_IP(gobjEtapaIP)
        End If
    End If
    
    Call Calcula_Total_Tvw

    Exit Sub

Erro_Preco_Validate:

    Cancel = True

    Select Case gErr

        Case 187785

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187642)

    End Select

    Exit Sub

End Sub

Private Sub Calcula_Total_Tvw()

Dim objNode As Node
Dim objEtapa As ClassPRJEtapas
Dim dCustoI As Double
Dim dCustoC As Double
Dim dPreco As Double
Dim objEtapaIP As ClassPRJEtapaItensProd

On Error GoTo Erro_Calcula_Total_Tvw

    For Each objEtapa In gobjProjeto.colEtapas
    
        Set objNode = TvwEtapas.Nodes.Item(objEtapa.iIndiceTvw)
    
        If objNode.Checked = True Then
            
            dCustoI = dCustoI + objEtapa.dCustoInfoPrev
            dCustoC = dCustoC + objEtapa.dCustoCalcPrev
            dPreco = dPreco + objEtapa.dPreco
        
        End If
    
        For Each objEtapaIP In objEtapa.colItensProduzidos
        
            If objEtapaIP.iIndiceTvw <> 0 Then
        
                Set objNode = TvwEtapas.Nodes.Item(objEtapaIP.iIndiceTvw)
            
                If objNode.Checked = True Then
                    
                    dCustoI = dCustoI + objEtapaIP.dCustoInfo
                    dPreco = dPreco + objEtapaIP.dPreco
                
                End If
                
            End If
        
        Next
    
    Next
    
    TotalCustoCalc.Caption = Format(dCustoC, "STANDARD")
    TotalCustoInf.Caption = Format(dCustoI, "STANDARD")
    TotalPreco.Caption = Format(dPreco, "STANDARD")

    Exit Sub

Erro_Calcula_Total_Tvw:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187643)

    End Select

    Exit Sub

End Sub

Private Sub TvwEtapas_NodeCheck(ByVal Node As MSComctlLib.Node)

Dim objEtapa As ClassPRJEtapas
Dim objEtapaIP As ClassPRJEtapaItensProd
Dim bAchou As Boolean

On Error GoTo Erro_TvwEtapas_NodeCheck

    bAchou = False
    For Each objEtapa In gobjProjeto.colEtapas

        If objEtapa.iIndiceTvw = Node.Index Then
            objEtapa.iTvwChecked = Node.Checked
            Exit For
        End If

        For Each objEtapaIP In objEtapa.colItensProduzidos
    
            If objEtapaIP.iIndiceTvw = Node.Index Then
                objEtapaIP.iTvwChecked = Node.Checked
                bAchou = True
                Exit For
            End If
    
        Next
        If bAchou Then Exit For

    Next
    
    Call Calcula_Total_Tvw

    Exit Sub

Erro_TvwEtapas_NodeCheck:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187644)

    End Select

    Exit Sub
    
End Sub

Private Sub Acerta_Texto_Node(ByVal objEtapa As ClassPRJEtapas)

Dim bAchou As Boolean
Dim sTexto As String
Dim sTextoAux As String
Dim objNode As Node

On Error GoTo Erro_Acerta_Texto_Node

    Set objNode = TvwEtapas.Nodes.Item(objEtapa.iIndiceTvw)

    sTexto = objEtapa.sNomeReduzido
    sTextoAux = ""
    
    If ExibirCustoCalc.Value = vbChecked Then
        sTextoAux = "Custo Calc: " & Format(objEtapa.dCustoCalcPrev, "STANDARD")
    End If
    
    If ExibirCustoInf.Value = vbChecked Then
        If Len(Trim(sTextoAux)) <> 0 Then
            sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Custo Inf: " & Format(objEtapa.dCustoInfoPrev, "STANDARD")
        Else
            sTextoAux = "Custo Inf: " & Format(objEtapa.dCustoInfoPrev, "STANDARD")
        End If
    End If
               
    If ExibirPreco.Value = vbChecked Then
        If Len(Trim(sTextoAux)) <> 0 Then
            sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Preço: " & Format(objEtapa.dPreco, "STANDARD")
        Else
            sTextoAux = "Preço: " & Format(objEtapa.dPreco, "STANDARD")
        End If
    End If
    
    If Len(Trim(sTextoAux)) > 0 Then
        sTexto = sTexto & " (" & sTextoAux & ")"
    End If
    
    objNode.Text = sTexto

    Exit Sub

Erro_Acerta_Texto_Node:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187645)

    End Select

    Exit Sub
    
End Sub

Private Sub Acerta_Texto_Node_IP(ByVal objEtapaIP As ClassPRJEtapaItensProd)

Dim lErro As Long
Dim bAchou As Boolean
Dim sTexto As String
Dim sTextoAux As String
Dim objNode As Node
Dim sProdutoMascarado As String

On Error GoTo Erro_Acerta_Texto_Node_IP

    Set objNode = TvwEtapas.Nodes.Item(objEtapaIP.iIndiceTvw)

    lErro = Mascara_RetornaProdutoTela(objEtapaIP.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 187786

    sTexto = sProdutoMascarado & SEPARADOR & objEtapaIP.sDescricao
    sTextoAux = ""
    
    If ExibirCustoCalc.Value = vbChecked Then
        sTextoAux = "Custo Calc: " & Format(0, "STANDARD")
    End If
    
    If ExibirCustoInf.Value = vbChecked Then
        If Len(Trim(sTextoAux)) <> 0 Then
            sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Custo Inf: " & Format(objEtapaIP.dCustoInfo, "STANDARD")
        Else
            sTextoAux = "Custo Inf: " & Format(objEtapaIP.dCustoInfo, "STANDARD")
        End If
    End If
               
    If ExibirPreco.Value = vbChecked Then
        If Len(Trim(sTextoAux)) <> 0 Then
            sTextoAux = sTextoAux & " " & SEPARADOR & " " & "Preço: " & Format(objEtapaIP.dPreco, "STANDARD")
        Else
            sTextoAux = "Preço: " & Format(objEtapaIP.dPreco, "STANDARD")
        End If
    End If
    
    If Len(Trim(sTextoAux)) > 0 Then
        sTexto = sTexto & " (" & sTextoAux & ")"
    End If
    
    objNode.Text = sTexto

    Exit Sub

Erro_Acerta_Texto_Node_IP:

    Select Case gErr
    
        Case 187786

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187646)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCondPag_Click()

Dim lErro As Long
Dim objPRJRecebPagto As New ClassPRJRecebPagto
Dim objContrato As New ClassPRJContratos

On Error GoTo Erro_BotaoCondPag_Click

    If Not (gobjProjeto Is Nothing) Then
        objPRJRecebPagto.lNumIntDocPRJ = gobjProjeto.lNumIntDoc
    End If
    
    If (Len(Trim(Codigo.ClipText)) > 0) And (Not (gobjProjeto Is Nothing)) Then
    
        objContrato.lNumIntDocPRJ = gobjProjeto.lNumIntDoc
        objContrato.sCodigo = Codigo.Text
    
        'Lê o Projetos que está sendo Passado
        lErro = CF("PRJContratos_Le", objContrato)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187787
    
    End If
    
    If objContrato.lNumIntDoc <> 0 Then
    
        objPRJRecebPagto.lNumIntDocContrato = objContrato.lNumIntDoc
        objPRJRecebPagto.iTipo = PRJ_TIPO_RECEB
        
        lErro = CF("PRJRecebPagto_Le_Contrato", objPRJRecebPagto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187788
        
    End If
    
    Call Chama_Tela("RecebimentoPRJ", objPRJRecebPagto)

    Exit Sub

Erro_BotaoCondPag_Click:

    Select Case gErr
    
        Case 187787, 187788

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187647)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoRefazer_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objEtapa As New ClassPRJEtapas
Dim objEtapaIP As New ClassPRJEtapaItensProd
Dim objNode As Node
Dim objContratoItem As ClassPRJContratoItem
Dim objContrato As New ClassPRJContratos

On Error GoTo Erro_BotaoRefazer_Click

    For Each objEtapa In gobjProjeto.colEtapas
    
        Set objNode = TvwEtapas.Nodes.Item(objEtapa.iIndiceTvw)
    
        If objNode.Checked = True Then
            
            Set objContratoItem = New ClassPRJContratoItem
            
            objContratoItem.dPrecoTotal = objEtapa.dPreco
            objContratoItem.dPrecoUnitario = objEtapa.dPreco
            objContratoItem.sCodEtapa = objEtapa.sCodigo
            objContratoItem.sDescEtapa = objEtapa.sDescricao
            objContratoItem.dtDataEntrega = objEtapa.dtDataFim
            
            objContrato.colItens.Add objContratoItem
        
        End If
        
        For Each objEtapaIP In objEtapa.colItensProduzidos
        
            If objEtapaIP.iIndiceTvw <> 0 Then
        
                Set objNode = TvwEtapas.Nodes.Item(objEtapaIP.iIndiceTvw)
            
                If objNode.Checked = True Then
                    
                    Set objContratoItem = New ClassPRJContratoItem
                    
                    objContratoItem.dPrecoTotal = objEtapaIP.dPreco
                    objContratoItem.dPrecoUnitario = objEtapaIP.dPreco / objEtapaIP.dQuantidade
                    objContratoItem.dQuantidade = objEtapaIP.dQuantidade
                    objContratoItem.sCodEtapa = objEtapa.sCodigo
                    objContratoItem.sDescEtapa = objEtapa.sDescricao
                    objContratoItem.sProduto = objEtapaIP.sProduto
                    objContratoItem.sDescProd = objEtapaIP.sDescricao
                    objContratoItem.sUM = objEtapaIP.sUM
                    objContratoItem.dtDataEntrega = DATA_NULA
                    
                    objContrato.colItens.Add objContratoItem
                
                End If
    
            End If
        Next
        
    Next
   
    For iIndice = gobjContrato.colItens.Count To 1 Step -1
        lErro = gobjTribTab.Exclusao_Item_Grid(iIndice)
        If lErro <> SUCESSO Then gError 187789
    Next
   
    lErro = Traz_Itens_Tela(objContrato)
    If lErro <> SUCESSO Then gError 187790
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objContrato.colItens.Item(iIndice).sProduto)
        If lErro <> SUCESSO Then gError 187791
    Next
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        Call PrecoTotal_Calcula(iIndice)
    Next
    
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 187792
    
    Exit Sub

Erro_BotaoRefazer_Click:

    Select Case gErr
    
        Case 187789 To 187792

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187648)

    End Select

    Exit Sub
    
End Sub

Function Traz_Itens_Tela(ByVal objContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim sProdutoMascarado As String
Dim objContratoItem As ClassPRJContratoItem
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_Traz_Itens_Tela

    Call Grid_Limpa(objGridItens)

    'Exibe os dados da coleção de Competencias na tela
    For Each objContratoItem In objContrato.colItens
        
        iLinha = iLinha + 1
        
        If objContratoItem.sProduto <> "" Then
        
            lErro = Mascara_RetornaProdutoTela(objContratoItem.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 187793
            
            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True
        
            GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text
        
        End If
        
        If objContratoItem.dtDataEntrega <> DATA_NULA Then
            GridItens.TextMatrix(iLinha, iGrid_DataEntrega_Col) = Format(objContratoItem.dtDataEntrega, "dd/mm/yyyy")
        Else
            GridItens.TextMatrix(iLinha, iGrid_DataEntrega_Col) = ""
        End If
        
        GridItens.TextMatrix(iLinha, iGrid_DescEtapa_Col) = objContratoItem.sDescEtapa
        
        If objContratoItem.dValorDesconto <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(objContratoItem.dValorDesconto, "STANDARD")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If
        
        GridItens.TextMatrix(iLinha, iGrid_DescProduto_Col) = objContratoItem.sDescProd
        
        If Len(Trim(objContratoItem.sCodEtapa)) > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Etapa_Col) = objContratoItem.sCodEtapa
        
        Else
                    
            Set objEtapa = New ClassPRJEtapas
        
            If objContratoItem.lNumIntDocEtapa <> 0 Then

                objEtapa.lNumIntDoc = objContratoItem.lNumIntDocEtapa
                
                lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187794
                
                If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187795
                
            End If
        
            GridItens.TextMatrix(iLinha, iGrid_Etapa_Col) = objEtapa.sCodigo
        
        End If
        
        GridItens.TextMatrix(iLinha, iGrid_Observacao_Col) = objContratoItem.sObservacao
                
        If objContratoItem.dValorDesconto <> 0 And objContratoItem.dPrecoTotal <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col) = Format(objContratoItem.dValorDesconto / (objContratoItem.dPrecoTotal + objContratoItem.dValorDesconto), "PERCENT")
        Else
            GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col) = ""
        End If
                
        If objContratoItem.dPrecoTotal <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(objContratoItem.dPrecoTotal, "STANDARD")
        Else
            GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = ""
        End If
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(objContratoItem.dPrecoTotal + objContratoItem.dValorDesconto, "Standard")
        
        If objContratoItem.dPrecoUnitario <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(objContratoItem.dPrecoUnitario, "STANDARD")
        Else
            GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = ""
        End If
        
        If objContratoItem.dQuantidade <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objContratoItem.dQuantidade)
        Else
            GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = ""
        End If
        
        GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = objContratoItem.sUM
    
    Next

    objGridItens.iLinhasExistentes = iLinha

    Traz_Itens_Tela = SUCESSO

    Exit Function

Erro_Traz_Itens_Tela:

    Traz_Itens_Tela = gErr

    Select Case gErr
    
        Case 187793, 187794
        
        Case 187795
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO", gErr, objEtapa.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187649)

    End Select
    
    Exit Function

End Function

Private Function Move_Itens_Memoria(ByVal objContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Itens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        lErro = Move_GridItem_Memoria(objContrato, iIndice)
        If lErro <> SUCESSO Then gError 187796
    
    Next

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr
    
        Case 187796

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187650)

    End Select

    Exit Function

End Function

Public Function Move_GridItem_Memoria(ByVal objContrato As ClassPRJContratos, ByVal iIndice As Integer, Optional ByVal sProduto As String = "") As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objContratoItem As New ClassPRJContratoItem
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_Move_GridItem_Memoria
    
    If Len(Trim(sProduto)) > 0 Then
        sProdutoFormatado = sProduto
    Else
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 187797
    End If
    
    objEtapa.sCodigo = GridItens.TextMatrix(iIndice, iGrid_Etapa_Col)
    objEtapa.lNumIntDocPRJ = objContrato.lNumIntDocPRJ
    
    lErro = CF("PRJEtapas_Le", objEtapa)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187798

    If lErro = SUCESSO Then
        objContratoItem.sCodEtapa = GridItens.TextMatrix(iIndice, iGrid_Etapa_Col)
        objContratoItem.sDescEtapa = GridItens.TextMatrix(iIndice, iGrid_DescEtapa_Col)
        objContratoItem.lNumIntDocEtapa = objEtapa.lNumIntDoc
    End If
        
    objContratoItem.iItem = iIndice
    objContratoItem.sProduto = sProdutoFormatado
    objContratoItem.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    objContratoItem.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
    objContratoItem.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
    objContratoItem.dtDataEntrega = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col))
    objContratoItem.dValorDescGlobal = StrParaDbl(ValorDesconto.Text)
    objContratoItem.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objContratoItem.iFilialEmpresa = giFilialEmpresa
    objContratoItem.sDescProd = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
    objContratoItem.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    objContratoItem.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)

    If gobjContrato.colItens.Count >= iIndice Then
        Set objContratoItem.objTributacaoPRJCTRItem = gobjContrato.colItens.Item(iIndice).objTributacaoPRJCTRItem
    Else
        Set objContratoItem.objTributacaoPRJCTRItem = Nothing
    End If

    objContrato.colItens.Add objContratoItem

    Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItem_Memoria:

    Move_GridItem_Memoria = gErr

    Select Case gErr
    
        Case 187797, 187798

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187651)

    End Select

    Exit Function

End Function

Function Move_Etapa_Memoria(ByVal objContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim objEtapa As ClassPRJEtapas
Dim objEtapaIP As ClassPRJEtapaItensProd
Dim objNode As Node
Dim objContratoEtapa As ClassPRJContratoEtapa
Dim iLinha As Integer

On Error GoTo Erro_Move_Etapa_Memoria

    If Not (gobjProjeto Is Nothing) Then

        For Each objEtapa In gobjProjeto.colEtapas
        
            iLinha = iLinha + 1
                    
            Set objContratoEtapa = New ClassPRJContratoEtapa
            
            objContratoEtapa.iImprimir = StrParaInt(GridEtapa.TextMatrix(iLinha, iGrid_Imprimir_Col))
            objContratoEtapa.sObservacao = GridEtapa.TextMatrix(iLinha, iGrid_ObservacaoGrid_Col)
            objContratoEtapa.sDescricao = GridEtapa.TextMatrix(iLinha, iGrid_DescricaoGrid_Col)

            objContrato.colEtapas.Add objContratoEtapa
        
            Set objNode = TvwEtapas.Nodes.Item(objEtapa.iIndiceTvw)
        
            If objNode.Checked = True Then
                objContratoEtapa.iSelecionado = MARCADO
            Else
                objContratoEtapa.iSelecionado = DESMARCADO
            End If
            
            objContratoEtapa.dCustoInformado = objEtapa.dCustoInfoPrev
            objContratoEtapa.dPreco = objEtapa.dPreco
            objContratoEtapa.lNumIntDocEtapa = objEtapa.lNumIntDoc
            
            For Each objEtapaIP In objEtapa.colItensProduzidos
            
                If objEtapaIP.iIndiceTvw <> 0 Then
            
                    Set objContratoEtapa = New ClassPRJContratoEtapa
                    
                    objContrato.colEtapas.Add objContratoEtapa
                
                    Set objNode = TvwEtapas.Nodes.Item(objEtapaIP.iIndiceTvw)
                
                    If objNode.Checked = True Then
                        objContratoEtapa.iSelecionado = MARCADO
                    Else
                        objContratoEtapa.iSelecionado = DESMARCADO
                    End If
                    
                    objContratoEtapa.dCustoInformado = objEtapaIP.dCustoInfo
                    objContratoEtapa.dPreco = objEtapaIP.dPreco
                    objContratoEtapa.lNumIntDocEtapa = objEtapa.lNumIntDoc
                    objContratoEtapa.lNumIntDocEtapaItemProd = objEtapaIP.lNumIntDoc
                    
                End If
            
            Next
            
        Next
        
    End If

    Move_Etapa_Memoria = SUCESSO

    Exit Function

Erro_Move_Etapa_Memoria:

    Move_Etapa_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187652)

    End Select

    Exit Function

End Function

Private Function Traz_Etapa_Tela(ByVal objContrato As ClassPRJContratos) As Long

Dim objNode As Node
Dim objEtapa As ClassPRJEtapas
Dim objEtapaIP As ClassPRJEtapaItensProd
Dim objContratoEtapa As ClassPRJContratoEtapa
Dim bAchou As Boolean
Dim iLinha As Integer

On Error GoTo Erro_Traz_Etapa_Tela

    Call Grid_Limpa(objGridEtapa)

    For Each objEtapa In gobjProjeto.colEtapas
    
        bAchou = False
        For Each objContratoEtapa In objContrato.colEtapas
            If objContratoEtapa.lNumIntDocEtapa = objEtapa.lNumIntDoc And objContratoEtapa.lNumIntDocEtapaItemProd = 0 Then
                bAchou = True
                Exit For
            End If
        Next
        
        If bAchou Then
        
            iLinha = iLinha + 1
            GridEtapa.TextMatrix(iLinha, iGrid_Imprimir_Col) = objContratoEtapa.iImprimir
            GridEtapa.TextMatrix(iLinha, iGrid_EtapaGrid_Col) = objEtapa.sCodigo
            GridEtapa.TextMatrix(iLinha, iGrid_DescricaoGrid_Col) = objContratoEtapa.sDescricao
            GridEtapa.TextMatrix(iLinha, iGrid_ObservacaoGrid_Col) = objContratoEtapa.sObservacao
                
            Set objNode = TvwEtapas.Nodes.Item(objEtapa.iIndiceTvw)
            
            objEtapa.dPreco = objContratoEtapa.dPreco
            objEtapa.dCustoInfoPrev = objContratoEtapa.dCustoInformado
            
            If objContratoEtapa.iSelecionado = MARCADO Then
                objNode.Checked = True
                objEtapa.iTvwChecked = objNode.Checked
            Else
                objNode.Checked = False
                objEtapa.iTvwChecked = objNode.Checked
            End If
            
            Call Acerta_Texto_Node(objEtapa)
        
            For Each objEtapaIP In objEtapa.colItensProduzidos
            
                bAchou = False
                For Each objContratoEtapa In objContrato.colEtapas
                    If objContratoEtapa.lNumIntDocEtapa = objEtapa.lNumIntDoc And objContratoEtapa.lNumIntDocEtapaItemProd = objEtapaIP.lNumIntDoc Then
                        bAchou = True
                        Exit For
                    End If
                Next
                
                If bAchou Then
           
                    If objEtapaIP.iIndiceTvw <> 0 Then
                
                        Set objNode = TvwEtapas.Nodes.Item(objEtapaIP.iIndiceTvw)
                                 
                        objEtapaIP.dPreco = objContratoEtapa.dPreco
                        objEtapaIP.dCustoInfo = objContratoEtapa.dCustoInformado
                         
                        If objContratoEtapa.iSelecionado = MARCADO Then
                            objNode.Checked = True
                            objEtapaIP.iTvwChecked = objNode.Checked
                        Else
                            objNode.Checked = False
                            objEtapaIP.iTvwChecked = objNode.Checked
                        End If
                         
                        Call Acerta_Texto_Node_IP(objEtapaIP)
                        
                    End If
                    
                End If
           
            Next
           
        End If
    
    Next
    
    objGridEtapa.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridEtapa)
    
    Traz_Etapa_Tela = SUCESSO

    Exit Function

Erro_Traz_Etapa_Tela:

    Traz_Etapa_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187653)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridEtapa(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Imprimir")
    objGrid.colColuna.Add ("Etapa")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Imprimir.Name)
    objGrid.colCampo.Add (EtapaGrid.Name)
    objGrid.colCampo.Add (DescricaoGrid.Name)
    objGrid.colCampo.Add (ObservacaoGrid.Name)

    'Colunas do Grid
    iGrid_Imprimir_Col = 1
    iGrid_EtapaGrid_Col = 2
    iGrid_DescricaoGrid_Col = 3
    iGrid_ObservacaoGrid_Col = 4

    objGrid.objGrid = GridEtapa

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridEtapa.ColWidth(0) = 300

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridEtapa = SUCESSO

End Function

Private Sub GridEtapa_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridEtapa, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEtapa, iAlterado)
    End If

End Sub

Private Sub GridEtapa_GotFocus()
    Call Grid_Recebe_Foco(objGridEtapa)
End Sub

Private Sub GridEtapa_EnterCell()
    Call Grid_Entrada_Celula(objGridEtapa, iAlterado)
End Sub

Private Sub GridEtapa_LeaveCell()
    Call Saida_Celula(objGridEtapa)
End Sub

Private Sub GridEtapa_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridEtapa, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEtapa, iAlterado)
    End If

End Sub

Private Sub GridEtapa_RowColChange()
    
    Call Grid_RowColChange(objGridEtapa)
    
    Call Recolhe_Observacao(iLinhaAnt)
    Call Traz_Observacao(GridEtapa.Row)
    
    iLinhaAnt = GridEtapa.Row
    
End Sub

Private Sub GridEtapa_Scroll()
    Call Grid_Scroll(objGridEtapa)
End Sub

Private Sub GridEtapa_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridEtapa)
End Sub

Private Sub GridEtapa_LostFocus()
    Call Grid_Libera_Foco(objGridEtapa)
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

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnterior As Integer
Dim iIndice As Integer
Dim dValorTotal As Double
Dim dValorTotalB As Double
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row
    
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        Call gobjTribTab.Exclusao_Item_Grid(iItemAtual)

        'Calcula a soma dos valores de produtos
        For iIndice = 1 To objGridItens.iLinhasExistentes
            dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
        Next
        
        If objGridItens.iLinhasExistentes <> 0 Then
            Call PrecoTotal_Calcula(objGridItens.iLinhasExistentes)
'        Else
'            If StrParaDbl(ValorDesconto.Text) <> 0 Then
'                'Avisa ao usuário
'                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)
'
'                'Limpa o valor de desconto
'                gdDesconto = 0
'                ValorDesconto.Text = ""
'                giValorDescontoAlterado = 0
'
'                Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)
'
'                'Para tributação
'                gobjContrato.dValorDesconto = 0
'
'            End If
        End If
        
        'Coloca valor total dos produtos na tela
        ValorProdutos.Caption = Format(dValorTotal, "Standard")
        ValorProdutos2.Caption = Format(dValorTotalB, "Standard")

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 187799

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 187799

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187654)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Etapa_Col

                    lErro = Saida_Celula_Etapa(objGridInt)
                    If lErro <> SUCESSO Then gError 187800

                Case iGrid_DescEtapa_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, DescEtapa)
                    If lErro <> SUCESSO Then gError 187801
                
                Case iGrid_Produto_Col
                
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 187802
                
                Case iGrid_DescProduto_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, DescProduto)
                    If lErro <> SUCESSO Then gError 187803
                
                Case iGrid_UnidadeMed_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, UM)
                    If lErro <> SUCESSO Then gError 187804

                Case iGrid_Quantidade_Col
                
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 187805
                
                Case iGrid_PrecoUnitario_Col
                
                    lErro = Saida_Celula_PrecoUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 187806
                
                Case iGrid_PercDesc_Col
                
                    lErro = Saida_Celula_PercentDesc(objGridInt)
                    If lErro <> SUCESSO Then gError 187807
                
                Case iGrid_DataEntrega_Col
                
                    lErro = Saida_Celula_DataEntrega(objGridInt)
                    If lErro <> SUCESSO Then gError 187808
                
                Case iGrid_Observacao_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Observacao)
                    If lErro <> SUCESSO Then gError 187809

            End Select

        ElseIf objGridInt.objGrid.Name = GridEtapa.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Imprimir_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Imprimir)
                    If lErro <> SUCESSO Then gError 187810

                Case iGrid_ObservacaoGrid_Col

                    lErro = Saida_Celula_Padrao(objGridInt, ObservacaoGrid)
                    If lErro <> SUCESSO Then gError 187811

            End Select
                         
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 187812

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 187800 To 187811
            'erros tratatos nas rotinas chamadas
        
        Case 187812
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187655)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable
        
    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 187813
        
    Select Case objControl.Name
    
        Case Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If
    
        Case UM.Name
            'guarda a um go grid nessa coluna
            sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

            UM.Enabled = True

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UM.Text

            UM.Clear

            If iProdutoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sProdutoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 187814

                If lErro = 28030 Then gError 187815

                objClasseUM.iClasse = objProduto.iClasseUM
                
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 187816
                
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UM.AddItem objUM.sSigla
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
                UM.Enabled = False
            End If

        Case DescProduto.Name, Quantidade.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case PrecoUnitario.Name, PercentDesc.Name, DataEntrega.Name, Observacao.Name
    
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Imprimir.Name
            objControl.Enabled = True
            
        Case Etapa.Name
        
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case DescEtapa.Name
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Etapa_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
    
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 187813, 187814, 187816
        
        Case 187815
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 187656)

    End Select

    Exit Sub

End Sub

Private Sub Imprimir_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Imprimir_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridEtapa)
End Sub

Private Sub Imprimir_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtapa)
End Sub

Private Sub Imprimir_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtapa.objControle = Imprimir
    lErro = Grid_Campo_Libera_Foco(objGridEtapa)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ObservacaoGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridEtapa)
End Sub

Private Sub ObservacaoGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtapa)
End Sub

Private Sub ObservacaoGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtapa.objControle = ObservacaoGrid
    lErro = Grid_Campo_Libera_Foco(objGridEtapa)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187817

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 187817
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187657)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Etapa(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_Saida_Celula_Etapa

    Set objGridInt.objControle = Etapa
    
    'Se o campo foi preenchido
    If Len(Trim(Etapa.Text)) > 0 Then
    
        If Len(Trim(Projeto.ClipText)) = 0 Then gError 187818

        objEtapa.sCodigo = Etapa.Text
        objEtapa.lNumIntDocPRJ = gobjProjeto.lNumIntDoc

        'Le a etapa
        lErro = CF("PRJEtapas_Le", objEtapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187819
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescEtapa_Col) = objEtapa.sDescricao
        
'        'Se o produto não está preenchido
'        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then
'            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = "un"
'            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(1)
'        End If
'
'        'verifica se precisa preencher o grid com uma nova linha
'        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
'            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
'
'            'permite que a tributacao reflita a inclusao de uma linha no grid
'            lErro = Tributacao_Inclusao_Item_Grid(GridItens.Row)
'            If lErro <> SUCESSO Then gError 189368
'
'        End If
    
    Else

        GridItens.TextMatrix(GridItens.Row, iGrid_DescEtapa_Col) = ""

'        'Se o produto não está preenchido
'        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then
'            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = ""
'            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = ""
'        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187820

    Saida_Celula_Etapa = SUCESSO

    Exit Function

Erro_Saida_Celula_Etapa:

    Saida_Celula_Etapa = gErr

    Select Case gErr

        Case 187818
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 187819, 187820, 189368
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187658)

    End Select

    Exit Function

End Function

Private Function Recolhe_Observacao(ByVal iLinha As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Recolhe_Observacao

    If iLinha <> 0 Then
        GridEtapa.TextMatrix(iLinha, iGrid_ObservacaoGrid_Col) = ObsEtapa.Text
    End If

    Exit Function

Erro_Recolhe_Observacao:

    Recolhe_Observacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187659)

    End Select

    Exit Function

End Function

Private Function Traz_Observacao(ByVal iLinha As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Observacao

    If iLinha <> 0 Then
        ObsEtapa.Text = GridEtapa.TextMatrix(iLinha, iGrid_ObservacaoGrid_Col)
    End If

    Exit Function

Erro_Traz_Observacao:

    Traz_Observacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187660)

    End Select

    Exit Function

End Function

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Observacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Etapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Etapa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Etapa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Etapa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Etapa
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub DescEtapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DescEtapa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub DescEtapa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub DescEtapa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescEtapa
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescProduto_Change()
    iAlterado = REGISTRO_ALTERADO
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

Private Sub PercentDesc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercentDesc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEntrega_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataEntrega
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

Private Sub PrecoUnitario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PrecoUnitario_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoTotal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PrecoTotal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula_DataEntrega(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Entrega que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataEntrega As Date
Dim dtDataEmissao As Date

On Error GoTo Erro_Saida_Celula_DataEntrega

    Set objGridInt.objControle = DataEntrega

    If Len(Trim(DataEntrega.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(DataEntrega.Text)
        If lErro <> SUCESSO Then gError 187821

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187822

    Saida_Celula_DataEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_DataEntrega:

    Saida_Celula_DataEntrega = gErr

    Select Case gErr

        Case 187821, 187822
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187661)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim sValorPercAnterior As String

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    If Len(PercentDesc.Text) > 0 Then
    
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 187823

        dPercentDesc = StrParaDbl(PercentDesc.Text)

        If Format(dPercentDesc, "#0.#0\%") <> GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) Then
            
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 187824

            PercentDesc.Text = Format(dPercentDesc, "Fixed")

        End If

    Else

        dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
        dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    sValorPercAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)
    If sValorPercAnterior = "" Then sValorPercAnterior = "0,00%"

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187825
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then
        iDescontoAlterado = REGISTRO_ALTERADO

        'Recalcula o preço total
        Call PrecoTotal_Calcula(GridItens.Row)

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 191251

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 187826

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 187823, 187825, 187826
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 187824
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187662)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitario

    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 187827

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
    
    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187828

    If Not bPrecoUnitarioIgual Then

        Call PrecoTotal_Calcula(GridItens.Row)

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 191251

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 187829

    End If

   Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case 187827, 187828, 187829
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187663)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto Data que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO And lErro <> 26658 Then gError 187830
        If lErro = 26658 Then gError 187831
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187832

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 187830 To 187832
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187664)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade que está deixando de ser a corrente

Dim lErro As Long
Dim bQuantidadeIgual As Boolean
Dim iIndice As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    bQuantidadeIgual = False

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 187833

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    If dQuantAnterior = StrParaDbl(Quantidade.Text) Then bQuantidadeIgual = True

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187834

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 And Not bQuantidadeIgual Then
    
        Call PrecoTotal_Calcula(GridItens.Row)
        
    End If
    
    If Not bQuantidadeIgual Then

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 187835

    End If
    
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 187835

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 187833 To 187835
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187665)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function
'
'Private Sub ComboICMSTipo_Click()
'
'    If ComboICMSTipo.ListIndex = -1 Then Exit Sub
'
'    If giTrazendoTribItemTela = 0 Then
'        Call BotaoGravarTribItem_Click
'    End If
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ComboIPITipo_Click()
'
'    If ComboIPITipo.ListIndex = -1 Then Exit Sub
'
'    If giTrazendoTribItemTela = 0 Then
'        Call BotaoGravarTribItem_Click
'    End If
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ICMSAliquotaItem_Change()
'
'    giICMSAliquotaItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ICMSAliquotaItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSAliquotaItem_Validate
'
'    If giICMSAliquotaItemAlterado Then
'
'        If Len(Trim(ICMSAliquotaItem.ClipText)) > 0 Then
'
'            lErro = Porcentagem_Critica(ICMSAliquotaItem.Text)
'            If lErro <> SUCESSO Then gError 187836
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSAliquotaItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSAliquotaItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187836
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187666)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ICMSBase_Change()
'
'    ICMSBase1.Caption = ICMSBase.Caption
'
'End Sub
'
'Private Sub ICMSBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSBaseItemAlterado = 1
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ICMSBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSBaseItem_Validate
'
'    If giICMSBaseItemAlterado Then
'
'        If Len(Trim(ICMSBaseItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSBaseItem.Text)
'            If lErro <> SUCESSO Then gError 187837
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187837
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187667)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ICMSPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSPercRedBaseItemAlterado = 1
'
'End Sub
'
'Private Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSPercRedBaseItem_Validate
'
'    If giICMSPercRedBaseItemAlterado Then
'
'        If Len(Trim(ICMSPercRedBaseItem.Text)) > 0 Then
'
'            lErro = Porcentagem_Critica(ICMSPercRedBaseItem.Text)
'            If lErro <> SUCESSO Then gError 187838
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSPercRedBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSPercRedBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187838
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187668)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ICMSSubstAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstAliquotaItemAlterado = 1
'
'End Sub
'
'Private Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSSubstAliquotaItem_Validate
'
'    If giICMSSubstAliquotaItemAlterado Then
'
'        If Len(Trim(ICMSSubstAliquotaItem.ClipText)) > 0 Then
'
'            lErro = Porcentagem_Critica(ICMSSubstAliquotaItem.Text)
'            If lErro <> SUCESSO Then gError 187839
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSSubstAliquotaItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSSubstAliquotaItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187839
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187669)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ICMSSubstBase_Change()
'
'    ICMSSubstBase1.Caption = ICMSSubstBase.Caption
'
'End Sub
'
'Private Sub ICMSSubstBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstBaseItemAlterado = 1
'
'End Sub
'
'Private Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSSubstBaseItem_Validate
'
'    If giICMSSubstBaseItemAlterado Then
'
'        If Len(Trim(ICMSSubstBaseItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSSubstBaseItem.Text)
'            If lErro <> SUCESSO Then gError 187840
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSSubstBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSSubstBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187840
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187670)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ICMSSubstValor_Change()
'
'    ICMSSubstValor1.Caption = ICMSSubstValor.Caption
'
'End Sub
'
'Private Sub ICMSSubstValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstValorItemAlterado = 1
'
'End Sub
'
'Private Sub ICMSSubstValorItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSSubstValorItem_Validate
'
'    If giICMSSubstValorItemAlterado Then
'
'        If Len(Trim(ICMSSubstValorItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSSubstValorItem.Text)
'            If lErro <> SUCESSO Then gError 187850
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSSubstValorItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSSubstValorItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187850
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187671)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ICMSValor_Change()
'
'    ICMSValor1.Caption = ICMSValor.Caption
'
'End Sub
'
'Private Sub ICMSValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSValorItemAlterado = 1
'
'End Sub
'
'Private Sub ICMSValorItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSValorItem_Validate
'
'    If giICMSValorItemAlterado Then
'
'        If Len(Trim(ICMSValorItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSValorItem.Text)
'            If lErro <> SUCESSO Then gError 187851
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSValorItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSValorItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187851
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187672)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub IPIAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIAliquotaItemAlterado = 1
'
'End Sub
'
'Private Sub IPIAliquotaItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIAliquotaItem_Validate
'
'    If giIPIAliquotaItemAlterado Then
'
'        If Len(Trim(IPIAliquotaItem.ClipText)) > 0 Then
'
'            lErro = Porcentagem_Critica(IPIAliquotaItem.Text)
'            If lErro <> SUCESSO Then gError 187852
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIAliquotaItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187852
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187673)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub IPIBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIBaseItemAlterado = 1
'
'End Sub
'
'Private Sub IPIBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIBaseItem_Validate
'
'    If giIPIBaseItemAlterado Then
'
'        If Len(Trim(IPIBaseItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(IPIBaseItem.Text)
'            If lErro <> SUCESSO Then gError 187853
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187853
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187674)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub IPIPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIPercRedBaseItemAlterado = 1
'
'End Sub
'
'Private Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIPercRedBaseItem_Validate
'
'    If giIPIPercRedBaseItemAlterado Then
'
'        If Len(Trim(IPIPercRedBaseItem.Text)) > 0 Then
'
'            lErro = Porcentagem_Critica(IPIPercRedBaseItem.Text)
'            If lErro <> SUCESSO Then gError 187854
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIPercRedBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIPercRedBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187854
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187675)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub IPIValor_Change()
'
'    IPIValor1.Caption = IPIValor.Caption
'
'End Sub
'
'Private Sub IPIValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIValorItemAlterado = 1
'
'End Sub
'
'Private Sub IPIValorItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIValorItem_Validate
'
'    If giIPIValorItemAlterado Then
'
'        If Len(Trim(IPIValorItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(IPIValorItem.Text)
'            If lErro <> SUCESSO Then gError 187855
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIValorItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIValorItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187855
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187676)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub IRAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giAliqIRAlterada = 1
'
'End Sub
'
'Private Sub IRAliquota_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dIRAliquota As Double, dIRValor As Double
'
'On Error GoTo Erro_IRAliquota_Validate
'
'    If giAliqIRAlterada = 0 Then Exit Sub
'
'    If Len(Trim(IRAliquota.ClipText)) > 0 Then
'
'        lErro = Porcentagem_Critica(IRAliquota.Text)
'        If lErro <> SUCESSO Then gError 187856
'
'    End If
'
'    Call BotaoGravarTrib
'
'    giAliqIRAlterada = 0
'
'    Exit Sub
'
'Erro_IRAliquota_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187856
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187677)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ISSAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giISSAliquotaAlterada = 1
'
'End Sub
'
'Private Sub ISSAliquota_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ISSAliquota_Validate
'
'    If giISSAliquotaAlterada = 0 Then Exit Sub
'
'    If Len(Trim(ISSAliquota.ClipText)) > 0 Then
'
'        lErro = Porcentagem_Critica(ISSAliquota.Text)
'        If lErro <> SUCESSO Then gError 187857
'
'    End If
'
'    Call BotaoGravarTrib
'
'    giISSAliquotaAlterada = 0
'
'    Exit Sub
'
'Erro_ISSAliquota_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187857
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187678)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ISSValor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giISSValorAlterado = 1
'
'End Sub
'
'Private Sub ISSValor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ISSValor_Validate
'
'    If giISSValorAlterado = 0 Then Exit Sub
'
'    If Len(Trim(ISSValor.ClipText)) > 0 Then
'
'        lErro = Valor_NaoNegativo_Critica(ISSValor.Text)
'        If lErro <> SUCESSO Then gError 187858
'
'    End If
'
'    Call BotaoGravarTrib
'
'    giISSValorAlterado = 0
'
'    Exit Sub
'
'Erro_ISSValor_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187858
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187679)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub LblTipoTrib_Click()
'
'Dim colSelecao As New Collection
'Dim objTipoTrib As New ClassTipoDeTributacaoMovto
'
'    'apenas tipos de saida
'    colSelecao.Add "0"
'    colSelecao.Add "0"
'
'    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTrib, objEventoTiposDeTributacao)
'
'End Sub
'
'Private Sub LblTipoTribItem_Click()
'
'    Call LblTipoTrib_Click
'
'End Sub
'
'Private Sub NaturezaItemLabel_Click()
'
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim colSelecao As New Collection
'Dim dtDataRef As Date, sSelecao As String
'
'    If Len(Trim(NaturezaOpItem.Text)) > 0 Then objNaturezaOp.sCodigo = NaturezaOpItem.Text
'
'    dtDataRef = StrParaDate(DataCriacao.Text)
'
'    sSelecao = "Codigo >= " & NATUREZA_SAIDA_COD_INICIAL & " AND Codigo <= " & NATUREZA_SAIDA_COD_FINAL & " AND {fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4")
'
'    Call Chama_Tela("NaturezaOperacaoLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, sSelecao)
'
'End Sub
'
'Private Sub NaturezaLabel_Click()
'
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim colSelecao As New Collection
'Dim dtDataRef As Date
'
'    'Se NaturezaOP estiver preenchida coloca no Obj
'    objNaturezaOp.sCodigo = NaturezaOp.Text
'
'    dtDataRef = DataCriacao.Text
'
'    'selecao p/obter apenas as nat de saida
'    colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
'    colSelecao.Add NATUREZA_SAIDA_COD_FINAL
'
'    'Chama a Tela de browse de NaturezaOp
'    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))
'
'End Sub
'
'Private Sub NaturezaOp_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giNaturezaOpAlterada = 1
'
'End Sub
'
'Private Sub NaturezaOp_GotFocus()
'
'Dim iNaturezaAux As Integer
'
'    iNaturezaAux = giNaturezaOpAlterada
'    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
'    giNaturezaOpAlterada = iNaturezaAux
'
'End Sub
'
'Private Sub NaturezaOp_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_NaturezaOp_Validate
'
'    'Se Natureza não está preenchida espelha no frame Tributação
'    If Len(Trim(NaturezaOp.ClipText)) = 0 Then
'
'        NatOpEspelho.Caption = ""
'        DescNatOp.Caption = ""
'
'    End If
'
'    'Verifica se a NaturezaOP foi informada
'    If Len(Trim(NaturezaOp.ClipText)) = 0 Or giNaturezaOpAlterada = 0 Then Exit Sub
'
'    objNaturezaOp.sCodigo = Trim(NaturezaOp.Text)
'
'    If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 94495
'
'    'Lê a NaturezaOp
'    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'    If lErro <> SUCESSO And lErro <> 17958 Then gError 187859
'
'    'Se não existir --> Erro
'    If lErro = 17958 Then gError 187860
'
'    'Espelha Natureza no frame de Tributação
'    NatOpEspelho.Caption = objNaturezaOp.sCodigo
'    DescNatOp.Caption = objNaturezaOp.sDescricao
'
'    If giTrazendoTribTela = 0 And gbCarregandoTela = False Then Call BotaoGravarTrib
'
'    giNaturezaOpAlterada = 0
'
'    Exit Sub
'
'Erro_NaturezaOp_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187859
'
'        Case 187860
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOp.Text)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
'            Else
'            End If
'
'        Case 94495
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187680)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub NaturezaOpItem_Change()
'
'    giNatOpItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub objEventoNaturezaOp_evSelecao(obj1 As Object)
'
'Dim objNaturezaOp As New ClassNaturezaOp
'
'    Set objNaturezaOp = obj1
'
'    If iFrameAtual = TAB_PRINCIPAL Then
'
'        'Preenche a natureza de Opereração do frame principal
'        NaturezaOp.Text = objNaturezaOp.sCodigo
'        Call NaturezaOp_Validate(bSGECancelDummy)
'
'    Else
'        'Preenche a NatOp do frame de tributação
'        NaturezaOpItem.Text = objNaturezaOp.sCodigo
'        Call NaturezaOpItem_Validate(bSGECancelDummy)
'
'    End If
'
'    Me.Show
'
'End Sub
'
'Private Sub objEventoTiposDeTributacao_evSelecao(obj1 As Object)
'
'Dim objTipoTrib As ClassTipoDeTributacaoMovto
'
'    Set objTipoTrib = obj1
'
'    If iFrameAtualTributacao = 1 Then
'
'        TipoTributacao.Text = objTipoTrib.iTipo
'        Call TipoTributacao_Validate(bSGECancelDummy)
'
'    Else
'
'        TipoTributacaoItem.Text = objTipoTrib.iTipo
'        Call TipoTributacaoItem_Validate(bSGECancelDummy)
'
'    End If
'
'    Me.Show
'
'    Exit Sub
'
'End Sub
'
'
'Private Sub OpcaoTributacao_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_OpcaoTributacao_Click
'
'    'Se frame selecionado não for o atual
'    If OpcaoTributacao.SelectedItem.Index <> iFrameAtualTributacao Then
'
'        If TabStrip_PodeTrocarTab(iFrameAtualTributacao, OpcaoTributacao, Me) <> SUCESSO Then Exit Sub
'
'        'Esconde o frame atual, mostra o novo
'        FrameTributacao(OpcaoTributacao.SelectedItem.Index).Visible = True
'        FrameTributacao(iFrameAtualTributacao).Visible = False
'        'Armazena novo valor de iFrameAtualTributacao
'        iFrameAtualTributacao = OpcaoTributacao.SelectedItem.Index
'
'        'se abriu o tab de detalhamento
'        If OpcaoTributacao.SelectedItem.Index = 2 Then
'            lErro = gobjTribTab.TabClick
'            If lErro <> SUCESSO Then gError 187861
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_OpcaoTributacao_Click:
'
'    Select Case gErr
'
'        Case 187861
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187681)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TipoTributacao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giTipoTributacaoAlterado = 1
'
'End Sub
'
'Private Sub TipoTributacao_GotFocus()
'
'Dim iTipoTributacaoAux As Integer
'
'    iTipoTributacaoAux = giTipoTributacaoAlterado
'    Call MaskEdBox_TrataGotFocus(TipoTributacao, iAlterado)
'    giTipoTributacaoAlterado = iTipoTributacaoAux
'
'End Sub
'
'Private Sub TipoTributacao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_TipoTributacao_Validate
'
'    If Len(Trim(TipoTributacao.Text)) = 0 Then
'        'Limpa o campo da descrição
'        DescTipoTrib.Caption = ""
'    End If
'
'    If (giTipoTributacaoAlterado = 1) Then
'
'        objTipoDeTributacao.iTipo = StrParaInt(TipoTributacao.Text)
'
'        If objTipoDeTributacao.iTipo <> 0 Then
'            lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
'            If lErro <> SUCESSO And lErro <> 27259 Then gError 187862
'
'            'Se não encontrou o Tipo da Tributação --> erro
'            If lErro = 27259 Then gError 187863
'        End If
'
'        DescTipoTrib.Caption = objTipoDeTributacao.sDescricao
'
'        Call BotaoGravarTrib
'
'        giTipoTributacaoAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_TipoTributacao_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 187862
'
'        Case 187863
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacao.Text)
'
'            If vbMsgRes = vbYes Then
'
'                Call Chama_Tela("TipoDeTributacao", objTipoDeTributacao)
'
'            Else
'            End If
'
'        Case Else
'            vbMsgRes = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187682)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
''Por Leo em 02/05/02
'Public Sub TipoTributacaoItem_Change()
'
'    giTipoTributacaoItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'
'End Sub
'
''Por Leo em 02/05/02
'Public Sub TipoTributacaoItem_GotFocus()
'
'Dim iTipoTributacaoItemAux As Integer
'
'    iTipoTributacaoItemAux = giTipoTributacaoItemAlterado
'
'    Call MaskEdBox_TrataGotFocus(TipoTributacaoItem, iAlterado)
'
'    giTipoTributacaoItemAlterado = iTipoTributacaoItemAux
'
'End Sub
'
''Por Leo em 02/05/02
'Public Sub TipoTributacaoItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_TipoTributacaoItem_Validate
'
'    'Se trocou o tipo de tributação
'    If giTipoTributacaoItemAlterado Then
'
'        objTributacaoTipo.iTipo = StrParaInt(TipoTributacaoItem.Text)
'        If objTributacaoTipo.iTipo <> 0 Then
'
'            lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'            If lErro <> SUCESSO And lErro <> 27259 Then gError 187864
'
'            'Se não encontrou o Tipo da Tributação --> erro
'            If lErro = 27259 Then gError 187865
'
'            DescTipoTribItem.Caption = objTributacaoTipo.sDescricao
'
'            Call BotaoGravarTribItem_Click
'
'        Else
'            'Limpa o campo
'            DescTipoTribItem.Caption = ""
'
'        End If
'
'        giTipoTributacaoItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_TipoTributacaoItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187864
'
'        Case 187865
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacaoItem.Text)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("TipoDeTributacao", objTributacaoTipo)
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187683)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function ValorTotal_Calcula() As Long
'Calcula o Valor Total do Pedido

'Dim dValorDespesas As Double
'Dim dValorProdutos As Double
Dim dValorTotal As Double
'Dim dValorFrete As Double
'Dim dValorSeguro As Double
'Dim dValorIPI As Double
'Dim dValorICMSSubst As Double
'Dim vbMsgRes As VbMsgBoxResult
'Dim dValorAposIR As Double
'Dim dValorIRRF As Double
Dim lErro As Long
'Dim dValorISS As Double

On Error GoTo Erro_ValorTotal_Calcula

'    If Not gbCarregandoTela Then
'        'Atualiza os valores de tributação
'        lErro =  gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 187866
'    End If
'
'    'Recolhe os valores da tela
'    If Len(Trim(ValorProdutos.Caption)) > 0 And IsNumeric(ValorProdutos.Caption) Then dValorProdutos = StrParaDbl(ValorProdutos.Caption)
'    If Len(Trim(ValorFrete.Text)) > 0 And IsNumeric(ValorFrete.Text) Then dValorFrete = StrParaDbl(ValorFrete.Text)
'    If Len(Trim(ValorIRRF.Text)) > 0 And IsNumeric(ValorIRRF.Text) Then dValorIRRF = StrParaDbl(ValorIRRF.Text)
'    If Len(Trim(ValorSeguro.Text)) > 0 And IsNumeric(ValorSeguro.Text) Then dValorSeguro = StrParaDbl(ValorSeguro.Text)
'    If Len(Trim(ValorDespesas.Text)) > 0 And IsNumeric(ValorDespesas.Text) Then dValorDespesas = StrParaDbl(ValorDespesas.Text)
'    If Len(Trim(ICMSSubstValor1.Caption)) > 0 And IsNumeric(ICMSSubstValor1.Caption) Then dValorICMSSubst = StrParaDbl(ICMSSubstValor1.Caption)
'    If Len(Trim(IPIValor1.Caption)) > 0 And IsNumeric(IPIValor1.Caption) Then dValorIPI = StrParaDbl(IPIValor1.Caption)
'    If Len(Trim(ISSValor.Text)) > 0 And IsNumeric(ISSValor.Text) And ISSIncluso.Value = vbUnchecked Then dValorISS = StrParaDbl(ISSValor.Text)
'
'    'Calcula o Valor Total
'    dValorTotal = dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas + dValorIPI + dValorICMSSubst + dValorISS
'
'    dValorAposIR = dValorTotal - dValorIRRF
'
'    If dValorTotal <> 0 And dValorAposIR < 0 And dValorIRRF > 0 Then
'
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_IR_FONTE_MAIOR_VALOR_TOTAL", dValorIRRF, dValorTotal)
'        ValorIRRF.Text = ""
'
'        Call ValorIRRF_Validate(bSGECancelDummy)
'
'        'Faz a atualização dos valores da tributação
'        lErro =  gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 187867
'
'    End If

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 187866

    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    Call ValorDescontoItens_Calcula

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 187866 ', 187867

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187684)

    End Select

    Exit Function

End Function

Public Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorDescontoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim iIndice As Integer

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o valor foi alterado
    If giValorDescontoAlterado = 0 Then Exit Sub

    'Vale o desconto que foi colocado aqui
    giValorDescontoManual = 1

    dValorDesconto = 0

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
        dValorProdutos = dValorProdutos + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    Next

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then

        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 187868

        dValorDesconto = StrParaDbl(ValorDesconto.Text)

        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        If dValorDesconto > dValorProdutos Then gError 187869

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    'Para tributação
    gobjContrato.dValorDesconto = dValorDesconto
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 187870

    giValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 187868, 187870

        Case 187869
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187685)

    End Select

    Exit Sub

End Sub

Public Sub ValorDespesas_Change()

    giValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDespesas_Validate(Cancel As Boolean)

Dim dValorDespesas As Double
Dim lErro As Long

On Error GoTo Erro_ValorDespesas_Validate

    If giValorDespesasAlterado = 0 Then Exit Sub

    'Se  estiver preenchido
    If Len(Trim(ValorDespesas.Text)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
        If lErro <> SUCESSO Then gError 187871

        dValorDespesas = StrParaDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

    'Para tributação
    gobjContrato.dValorOutrasDespesas = dValorDespesas
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
'*** fim tributacao
    
    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 187872

    giValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case 187871, 187872

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187686)

    End Select

    Exit Sub

End Sub

Public Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorFreteAlterado = 1

End Sub

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim dValorFrete As Double
Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    If giValorFreteAlterado = 0 Then Exit Sub

    dValorFrete = 0

    If Len(Trim(ValorFrete.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 187873

        dValorFrete = StrParaDbl(ValorFrete.Text)

        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

    'Para tributação
    gobjContrato.dValorFrete = dValorFrete
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 187874
    
    giValorFreteAlterado = 0

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case gErr

        Case 187873, 187874

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187687)

    End Select

    Exit Sub

End Sub
'
'Public Sub ValorIRRF_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorIRRF As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_ValorIRRF_Validate
'
'    If giValorIRRFAlterado = 0 Then Exit Sub
'
'    'Verifica se ValorIRRF foi preenchido
'    If Len(Trim(ValorIRRF.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
'        If lErro <> SUCESSO Then gError 187875
'
'        dValorIRRF = StrParaDbl(ValorIRRF.Text)
'
'        ValorIRRF.Text = Format(dValorIRRF, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = StrParaDbl(ValorTotal.Caption)
'
'        If dValorIRRF > dValorTotal Then gError 187876
'
'    End If
'
'    Call BotaoGravarTrib
'
'    giValorIRRFAlterado = 0
'
'    Exit Sub
'
'Erro_ValorIRRF_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187875
'
'        Case 187876
'            Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187688)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorSeguroAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim dValorSeguro As Double
Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    If giValorSeguroAlterado = 0 Then Exit Sub

    dValorSeguro = 0

    If Len(Trim(ValorSeguro.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 187877

        dValorSeguro = StrParaDbl(ValorSeguro.Text)

        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

    'Para tributação
    gobjContrato.dValorSeguro = dValorSeguro
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 187878

    giValorSeguroAlterado = 0

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True

    Select Case gErr

        Case 187877, 187878

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187689)

    End Select

    Exit Sub

End Sub

Function Produto_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult
Dim sCliente As String
Dim iFilialCli As Integer

On Error GoTo Erro_Produto_Saida_Celula

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 187879
       
    If lErro = 86295 Then gError 187880
        
    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 187881
           
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 187882

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

    End If

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(gdDesconto, "Percent")

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
               
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 187883
        
    Else
    
        gobjContrato.colItens(GridItens.Row).sProduto = objProduto.sCodigo
        gobjContrato.colItens(GridItens.Row).sDescProd = objProduto.sDescricao
        
    End If

    'Atualiza a checkbox do grid para exibir a figura marcada/desmarcada
    Call Grid_Refresh_Checkbox(objGridItens)

    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr
    
        Case 187879, 187883
        
        Case 187880
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 187881
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If
            
        Case 187882
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187690)

    End Select

    Exit Function

End Function

Public Sub PrecoTotal_Calcula(iLinha As Integer)

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double
Dim dValorTotalB As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objGridItens1 As Object

On Error GoTo Erro_PrecoTotal_Calcula

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    'Cálculo do desconto
    lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))
    If lTamanho > 0 Then
        dPercentDesc = StrParaDbl(Format(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade)

    'Se percentual for >0 tira o desconto
    dDesconto = dPercentDesc * dPrecoTotal
    dPrecoTotalReal = dPrecoTotal - dDesconto 'Inserido por Wagner
        
    GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")

    'Coloca preco total do ítem no grid
    GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")

    GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dPrecoTotal, "Standard")

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
        dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
        dValorDesconto = dValorDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    Next

'    If gdDesconto > 0 Then
'        dValorDesconto = gdDesconto * dValorTotal
'    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
'        dValorDesconto = StrParaDbl(ValorDesconto.Text)
'    End If
'    dValorTotal = dValorTotal - dValorDesconto
'
'    'Verifica se o valor de desconto é maior que o valor dos produtos
'    If dValorTotal < 0 And dValorDesconto > 0 Then
'
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDesconto, dValorTotal)
'
'        gdDesconto = 0
'        ValorDesconto.Text = ""
'        giValorDescontoAlterado = 0
'        dValorDesconto = 0
'
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
'
'        'Para tributação
'        gobjContrato.dValorDesconto = dValorDesconto
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 187884
'
'        'Calcula a soma dos valores de produtos
'        dValorTotal = 0
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'        Next
'
'    End If

    'Coloca valor total dos produtos na tela
    ValorProdutos.Caption = Format(dValorTotal, "Standard")
    ValorProdutos2.Caption = Format(dValorTotalB, "Standard")
    ValorDescontoItens.Text = Format(dValorDesconto, "Standard")
    
    'Call Tributacao_Alteracao_Item_Grid(iLinha)

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr

        Case 187884

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187691)

    End Select

    Exit Sub

End Sub
'
'Private Sub NaturezaOpItem_Validate(Cancel As Boolean)
'
'Dim sNatOp As String
'Dim lErro As Long
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_NaturezaOpItem_Validate
'
'    If giNatOpItemAlterado = 0 Then Exit Sub
'
'    sNatOp = Trim(NaturezaOpItem.Text)
'
'    If sNatOp <> "" Then
'
'        objNaturezaOp.sCodigo = sNatOp
'
'        If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 187885
'
'        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'        If lErro <> SUCESSO And lErro <> 17958 Then gError 187886
'
'        'Se não achou a Natureza de Operação --> erro
'        If lErro <> SUCESSO Then gError 187887
'
'        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
'
'        Call BotaoGravarTribItem_Click
'
'    Else
'
'        'Limpa a descrição
'        LabelDescrNatOpItem.Caption = ""
'
'    End If
'
'    giNatOpItemAlterado = 0
'
'    Exit Sub
'
'Erro_NaturezaOpItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187885
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
'
'        Case 187886
'
'        Case 187887
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOpItem.Text)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187692)
'
'    End Select
'
'End Sub
'
'Private Sub NaturezaOpItem_GotFocus()
'
'Dim iNaturezaOpAux As Integer
'
'    iNaturezaOpAux = giNatOpItemAlterado
'
'    Call MaskEdBox_TrataGotFocus(NaturezaOpItem, iAlterado)
'
'    giNatOpItemAlterado = iNaturezaOpAux
'
'End Sub
'
'Private Sub TributacaoRecalcular_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_TributacaoRecalcular_Click
'
'    giRecalculandoTributacao = 1
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 187888
'
'    giRecalculandoTributacao = 0
'
'    Exit Sub
'
'Erro_TributacaoRecalcular_Click:
'
'    Select Case gErr
'
'        Case 187888
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187693)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ValorIRRF_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giValorIRRFAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ISSIncluso_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_ISSIncluso_Click
'
'    iAlterado = REGISTRO_ALTERADO
'
'    If giTrazendoTribTela = 0 And gbCarregandoTela = False Then Call BotaoGravarTrib
'
'    Exit Sub
'
'Erro_ISSIncluso_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187694)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ComboItensTrib_Click()
'
'Dim iIndice As Integer, objContratoItem As ClassPRJContratoItem
'
'    iIndice = ComboItensTrib.ListIndex
'
'    If iIndice <> -1 Then
'
'        'preenche os campos da tela em funcao do item selecionado
'
'        Set objContratoItem = gobjContrato.colItens.Item(iIndice + 1)
'
'        LabelValorItem.Caption = Format(objContratoItem.dPrecoTotal, "Standard")
'        LabelQtdeItem.Caption = CStr(objContratoItem.dQuantidade)
'        LabelUMItem.Caption = objContratoItem.sUM
'
'        Call TributacaoItem_TrazerTela(objContratoItem.objTributacaoPRJCTRItem)
'
'    End If
'
'End Sub
'
'Private Sub TribSobreDesconto_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'
'    End If
'
'    Call TributacaoItem_TrazerTela(gobjContrato.objTributacaoPRJCTR.objTributacaoDesconto)
'
'End Sub
'
'Private Sub TribSobreOutrasDesp_Click()
'
'   'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    Call TributacaoItem_TrazerTela(gobjContrato.objTributacaoPRJCTR.objTributacaoOutras)
'
'
'End Sub
'
'Private Sub TribSobreSeguro_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'
'    End If
'
'    Call TributacaoItem_TrazerTela(gobjContrato.objTributacaoPRJCTR.objTributacaoSeguro)
'
'End Sub
'
'Private Sub TribSobreFrete_Click()
'
'    'exibir o frame de "outros"
'    FrameOutrosTrib.Visible = True
'    FrameItensTrib.Visible = False
'
'    Call TributacaoItem_TrazerTela(gobjContrato.objTributacaoPRJCTR.objTributacaoFrete)
'
'End Sub
'
'Private Sub TribSobreItem_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    'se houver itens na combo
'    If gobjContrato.colItens.Count <> 0 Then
'
'        'mostra o frame de itens e esconde o de outros
'        FrameItensTrib.Visible = True
'        FrameOutrosTrib.Visible = False
'
'        'selecionar o 1o item
'        ComboItensTrib.ListIndex = 0
'
'        Call ComboItensTrib_Click
'
'    Else
'
'        'senao houver itens na combo selecionar Frete
'        TribSobreFrete.Value = True
'
'        Call TribSobreFrete_Click
'
'    End If
'
'End Sub
'
'Private Function TributacaoPRJCTR_Reset(Optional objContrato As ClassPRJContratos) As Long
''cria ou atualiza gobjContrato, com dados correspondentes a objContrato (se este for passado) ou com dados "padrao"
'
'Dim lErro As Long
'Dim objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoPRJCTR_Reset
'
'    'se gobjContrato já foi inicializado
'    If Not (gobjContrato Is Nothing) Then
'
'        Set objTributoDoc = gobjContrato
'
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then gError 187889
'
'        Set gobjContrato = Nothing
'
'    End If
'
'    'se o pedido de venda veio preenchido
'    If Not (objContrato Is Nothing) Then
'
'        Set gobjContrato = objContrato
'
'    Else
'
'        Set gobjContrato = New ClassPRJContratos
'        gobjContrato.dtData = gdtDataAtual
'
'    End If
'
'    Set objTributoDoc = gobjContrato
'    lErro = objTributoDoc.Ativar
'    If lErro <> SUCESSO Then gError 187890
'
'    giNaturezaOpAlterada = 0
'    giISSAliquotaAlterada = 0
'    giISSValorAlterado = 0
'    giValorIRRFAlterado = 0
'    giTipoTributacaoAlterado = 0
'    giAliqIRAlterada = 0
'    iPISRetidoAlterado = 0
'    iISSRetidoAlterado = 0
'    iCOFINSRetidoAlterado = 0
'    iCSLLRetidoAlterado = 0
'
'    giNatOpItemAlterado = 0
'    giTipoTributacaoItemAlterado = 0
'    giICMSBaseItemAlterado = 0
'    giICMSPercRedBaseItemAlterado = 0
'    giICMSAliquotaItemAlterado = 0
'    giICMSValorItemAlterado = 0
'    giICMSSubstBaseItemAlterado = 0
'    giICMSSubstAliquotaItemAlterado = 0
'    giICMSSubstValorItemAlterado = 0
'    giIPIBaseItemAlterado = 0
'    giIPIPercRedBaseItemAlterado = 0
'    giIPIAliquotaItemAlterado = 0
'    giIPIValorItemAlterado = 0
'
'    TributacaoPRJCTR_Reset = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPRJCTR_Reset:
'
'    TributacaoPRJCTR_Reset = gErr
'
'    Select Case gErr
'
'        Case 187889, 187890
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187695)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoGravarTrib()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoGravarTrib
'
'    lErro = Tributacao_GravarTela()
'    If lErro <> SUCESSO Then gError 187891
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 187892
'
'    lErro = Carrega_Tab_Tributacao(gobjContrato)
'    If lErro <> SUCESSO Then gError 187893
'
'    Exit Sub
'
'Erro_BotaoGravarTrib:
'
'    Select Case gErr
'
'        Case 187891 To 187893
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187696)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Tributacao_GravarTela() As Long
''transfere dados de tributacao da tela para gobjContrato
''os dados que estiverem diferentes devem ser marcados como "manuais"
'
'Dim lErro As Long
'Dim iIndice As Integer, iTemp As Integer, dTemp As Double, objTributacaoPRJCTR As ClassTributacaoPRJCTR
'
'On Error GoTo Erro_Tributacao_GravarTela
'
'    Set objTributacaoPRJCTR = gobjContrato.objTributacaoPRJCTR
'
'    If gobjContrato.sNaturezaOp <> NaturezaOp.Text Then
'
'        gobjContrato.sNaturezaOp = NaturezaOp.Text
'        gobjContrato.iNaturezaOpManual = VAR_PREENCH_MANUAL
'
'    End If
'
'    iTemp = StrParaInt(TipoTributacao.Text)
'    If iTemp <> objTributacaoPRJCTR.iTipoTributacao Then
'        objTributacaoPRJCTR.iTipoTributacao = iTemp
'        objTributacaoPRJCTR.iTipoTributacaoManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados de ISS
'    iTemp = ISSIncluso.Value
'    If iTemp <> objTributacaoPRJCTR.iISSIncluso Then
'        objTributacaoPRJCTR.iISSIncluso = iTemp
'        objTributacaoPRJCTR.iISSInclusoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ISSAliquota.Text <> CStr(objTributacaoPRJCTR.dISSAliquota * 100) Then
'        dTemp = StrParaDbl(ISSAliquota.Text) / 100
'        If objTributacaoPRJCTR.dISSAliquota <> dTemp Then
'            objTributacaoPRJCTR.dISSAliquota = dTemp
'            objTributacaoPRJCTR.iISSAliquotaManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If ISSValor.Text <> CStr(objTributacaoPRJCTR.dISSValor) Then
'        dTemp = StrParaDbl(ISSValor.Text)
'        If objTributacaoPRJCTR.dISSValor <> dTemp Then
'            objTributacaoPRJCTR.dISSValor = dTemp
'            objTributacaoPRJCTR.iISSValorManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    'setar dados de IR
'    If IRAliquota.Text <> CStr(objTributacaoPRJCTR.dIRRFAliquota * 100) Then
'        dTemp = StrParaDbl(IRAliquota.Text) / 100
'        If objTributacaoPRJCTR.dIRRFAliquota <> dTemp Then
'            objTributacaoPRJCTR.dIRRFAliquota = dTemp
'            objTributacaoPRJCTR.iIRRFAliquotaManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If ValorIRRF.Text <> CStr(objTributacaoPRJCTR.dIRRFValor) Then
'        dTemp = StrParaDbl(ValorIRRF.Text)
'        If objTributacaoPRJCTR.dIRRFValor <> dTemp Then
'            objTributacaoPRJCTR.dIRRFValor = dTemp
'            objTributacaoPRJCTR.iIRRFValorManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If PISRetido.Text <> CStr(objTributacaoPRJCTR.dPISRetido) Then
'        dTemp = StrParaDbl(PISRetido.Text)
'        If objTributacaoPRJCTR.dPISRetido <> dTemp Then
'            objTributacaoPRJCTR.dPISRetido = dTemp
'            objTributacaoPRJCTR.iPISRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If ISSRetido.Text <> CStr(objTributacaoPRJCTR.dISSRetido) Then
'        dTemp = StrParaDbl(ISSRetido.Text)
'        If objTributacaoPRJCTR.dISSRetido <> dTemp Then
'            objTributacaoPRJCTR.dISSRetido = dTemp
'            objTributacaoPRJCTR.iISSRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If COFINSRetido.Text <> CStr(objTributacaoPRJCTR.dCOFINSRetido) Then
'        dTemp = StrParaDbl(COFINSRetido.Text)
'        If objTributacaoPRJCTR.dCOFINSRetido <> dTemp Then
'            objTributacaoPRJCTR.dCOFINSRetido = dTemp
'            objTributacaoPRJCTR.iCOFINSRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If CSLLRetido.Text <> CStr(objTributacaoPRJCTR.dCSLLRetido) Then
'        dTemp = StrParaDbl(CSLLRetido.Text)
'        If objTributacaoPRJCTR.dCSLLRetido <> dTemp Then
'            objTributacaoPRJCTR.dCSLLRetido = dTemp
'            objTributacaoPRJCTR.iCSLLRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    Tributacao_GravarTela = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_GravarTela:
'
'    Tributacao_GravarTela = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187697)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Carrega_Tab_Tributacao(objContrato As ClassPRJContratos) As Long
'
'Dim lErro As Long
'Dim objTributacaoPRJCTR As ClassTributacaoPRJCTR
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'
'On Error GoTo Erro_Carrega_Tab_Tributacao
'
'    giTrazendoTribTela = 1
'
'    Set objTributacaoPRJCTR = objContrato.objTributacaoPRJCTR
'
'    If NaturezaOp.Text <> objContrato.sNaturezaOp Then
'
'        NaturezaOp.Text = objContrato.sNaturezaOp
'        Call NaturezaOp_Validate(bSGECancelDummy)
'
'    End If
'
'    'no frame de "resumo"
'    objTributacaoTipo.iTipo = objTributacaoPRJCTR.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        TipoTributacao.Text = CStr(objTributacaoPRJCTR.iTipoTributacao)
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then gError 187894
'
'        DescTipoTrib.Caption = objTributacaoTipo.sDescricao
'
'        'se nao incide ISS
'        If objTributacaoTipo.iISSIncide = 0 Then
'            ISSValor.Enabled = False
'            ISSAliquota.Enabled = False
'            ISSIncluso.Enabled = False
'        Else
'            ISSValor.Enabled = True
'            ISSAliquota.Enabled = True
'            ISSIncluso.Enabled = True
'        End If
'
'        'se nao incide IR
'        If objTributacaoTipo.iIRIncide = 0 Then
'            ValorIRRF.Enabled = False
'            IRAliquota.Enabled = False
'        Else
'            ValorIRRF.Enabled = True
'            IRAliquota.Enabled = True
'        End If
'
'        'se nao retem PIS
'        If objTributacaoTipo.iPISRetencao = 0 Then
'            PISRetido.Enabled = False
'        Else
'            PISRetido.Enabled = True
'        End If
'
'        'se nao retem ISS
'        If objTributacaoTipo.iISSRetencao = 0 Then
'            ISSRetido.Enabled = False
'        Else
'            ISSRetido.Enabled = True
'        End If
'
'        'se nao retem COFINS
'        If objTributacaoTipo.iCOFINSRetencao = 0 Then
'            COFINSRetido.Enabled = False
'        Else
'            COFINSRetido.Enabled = True
'        End If
'
'        'se nao retem CSLL
'        If objTributacaoTipo.iCSLLRetencao = 0 Then
'            CSLLRetido.Enabled = False
'        Else
'            CSLLRetido.Enabled = True
'        End If
'
'    Else
'
'        TipoTributacao.Text = ""
'        DescTipoTrib.Caption = ""
'
'    End If
'
'    IPIBase.Caption = Format(objTributacaoPRJCTR.dIPIBase, "Standard")
'    IPIValor.Caption = Format(objTributacaoPRJCTR.dIPIValor, "Standard")
'    ISSBase.Caption = Format(objTributacaoPRJCTR.dISSBase, "Standard")
'    ISSAliquota.Text = CStr(objTributacaoPRJCTR.dISSAliquota * 100)
'    ISSValor.Text = CStr(objTributacaoPRJCTR.dISSValor)
'    ISSIncluso.Value = objTributacaoPRJCTR.iISSIncluso
'    IRBase.Caption = Format(objTributacaoPRJCTR.dIRRFBase, "Standard")
'    IRAliquota.Text = CStr(objTributacaoPRJCTR.dIRRFAliquota * 100)
'    ValorIRRF.Text = CStr(objTributacaoPRJCTR.dIRRFValor)
'    ICMSBase.Caption = Format(objTributacaoPRJCTR.dICMSBase, "Standard")
'    ICMSValor.Caption = Format(objTributacaoPRJCTR.dICMSValor, "Standard")
'    ICMSSubstBase.Caption = Format(objTributacaoPRJCTR.dICMSSubstBase, "Standard")
'    ICMSSubstValor.Caption = Format(objTributacaoPRJCTR.dICMSSubstValor, "Standard")
'    PISRetido.Text = CStr(objTributacaoPRJCTR.dPISRetido)
'    ISSRetido.Text = CStr(objTributacaoPRJCTR.dISSRetido)
'    COFINSRetido.Text = CStr(objTributacaoPRJCTR.dCOFINSRetido)
'    CSLLRetido.Text = CStr(objTributacaoPRJCTR.dCSLLRetido)
'
'    'o frame de "detalhamento" vou deixar p/carregar qdo o usuario entrar nele
'
'    giISSAliquotaAlterada = 0
'    giISSValorAlterado = 0
'    giValorIRRFAlterado = 0
'    giTipoTributacaoAlterado = 0
'    giAliqIRAlterada = 0
'    iPISRetidoAlterado = 0
'    iISSRetidoAlterado = 0
'    iCOFINSRetidoAlterado = 0
'    iCSLLRetidoAlterado = 0
'
'    giTrazendoTribTela = 0
'
'    Carrega_Tab_Tributacao = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Tab_Tributacao:
'
'    giTrazendoTribTela = 0
'
'    Carrega_Tab_Tributacao = gErr
'
'    Select Case gErr
'
'        Case 187894
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187698)
'
'    End Select
'
'End Function
'
'Private Sub BotaoGravarTribItem_Click()
'
'Dim lErro As Long, objTributacaoPRJCTRItem As ClassTributacaoPRJCTRItem, iIndice As Integer
'
'On Error GoTo Erro_BotaoGravarTribItem_Click
'
'    'atualizar dados da colecao p/o item ou complemento corrente
'
'    'se um item estiver selecionado
'    If TribSobreItem.Value = True Then
'        iIndice = ComboItensTrib.ListIndex
'        If iIndice <> -1 Then
'            Set objTributacaoPRJCTRItem = gobjContrato.colItens.Item(iIndice + 1).objTributacaoPRJCTRItem
'        Else
'            gError 187895
'        End If
'    Else
'        If TribSobreDesconto.Value = True Then
'            Set objTributacaoPRJCTRItem = gobjContrato.objTributacaoPRJCTR.objTributacaoDesconto
'        Else
'            If TribSobreFrete.Value = True Then
'                Set objTributacaoPRJCTRItem = gobjContrato.objTributacaoPRJCTR.objTributacaoFrete
'            Else
'                If TribSobreSeguro.Value = True Then
'                    Set objTributacaoPRJCTRItem = gobjContrato.objTributacaoPRJCTR.objTributacaoSeguro
'                Else
'                    If TribSobreOutrasDesp.Value = True Then
'                        Set objTributacaoPRJCTRItem = gobjContrato.objTributacaoPRJCTR.objTributacaoOutras
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    lErro = TributacaoItem_GravarTela(objTributacaoPRJCTRItem)
'    If lErro <> SUCESSO Then gError 187896
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 187897
'
'    lErro = TributacaoItem_TrazerTela(objTributacaoPRJCTRItem)
'    If lErro <> SUCESSO Then gError 187898
'
'    Exit Sub
'
'Erro_BotaoGravarTribItem_Click:
'
'    Select Case gErr
'
'        Case 187895
'            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEM_TRIB_SEL", gErr, Error)
'
'        Case 187896 To 187898
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187699)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function TributacaoItem_GravarTela(objTributacaoPRJCTRItem As ClassTributacaoPRJCTRItem) As Long
''transfere dados de tributacao de um item da tela para objTributacaoPRJCTRItem
''os dados que estiverem diferentes devem ser marcados como "manuais"
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iTemp As Integer
'Dim dTemp As Double
'Dim sTemp As String
'
'On Error GoTo Erro_TributacaoItem_GravarTela
'
'    sTemp = Trim(NaturezaOpItem.Text)
'    If Trim(objTributacaoPRJCTRItem.sNaturezaOp) <> sTemp Then
'        objTributacaoPRJCTRItem.sNaturezaOp = sTemp
'        objTributacaoPRJCTRItem.iNaturezaOpManual = VAR_PREENCH_MANUAL
'    End If
'
'    iTemp = StrParaInt(TipoTributacaoItem.Text)
'    If iTemp <> objTributacaoPRJCTRItem.iTipoTributacao Then
'        objTributacaoPRJCTRItem.iTipoTributacao = iTemp
'        objTributacaoPRJCTRItem.iTipoTributacaoManual = VAR_PREENCH_MANUAL
'    End If
'
'    'Setar dados de ICMS
'
'    iTemp = ComboICMSTipo.ItemData(ComboICMSTipo.ListIndex)
'    If iTemp <> objTributacaoPRJCTRItem.iICMSTipo Then
'        objTributacaoPRJCTRItem.iICMSTipo = iTemp
'        objTributacaoPRJCTRItem.iICMSTipoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSBaseItem.Text <> CStr(objTributacaoPRJCTRItem.dICMSBase) Then
'        dTemp = StrParaDbl(ICMSBaseItem.Text)
'        objTributacaoPRJCTRItem.dICMSBase = dTemp
'        objTributacaoPRJCTRItem.iICMSBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSPercRedBaseItem.Text <> CStr(objTributacaoPRJCTRItem.dICMSPercRedBase * 100) Then
'        dTemp = StrParaDbl(ICMSPercRedBaseItem.Text) / 100
'        objTributacaoPRJCTRItem.dICMSPercRedBase = dTemp
'        objTributacaoPRJCTRItem.iICMSPercRedBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSAliquotaItem.Text <> CStr(objTributacaoPRJCTRItem.dICMSAliquota * 100) Then
'        dTemp = StrParaDbl(ICMSAliquotaItem.Text) / 100
'        objTributacaoPRJCTRItem.dICMSAliquota = dTemp
'        objTributacaoPRJCTRItem.iICMSAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSValorItem.Text <> CStr(objTributacaoPRJCTRItem.dICMSValor) Then
'        dTemp = StrParaDbl(ICMSValorItem.Text)
'        objTributacaoPRJCTRItem.dICMSValor = dTemp
'        objTributacaoPRJCTRItem.iICMSValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados ICMS Substituicao
'
'    If ICMSSubstBaseItem.Text <> CStr(objTributacaoPRJCTRItem.dICMSSubstBase) Then
'        dTemp = StrParaDbl(ICMSSubstBaseItem.Text)
'        objTributacaoPRJCTRItem.dICMSSubstBase = dTemp
'        objTributacaoPRJCTRItem.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSSubstAliquotaItem.Text <> CStr(objTributacaoPRJCTRItem.dICMSSubstAliquota * 100) Then
'        dTemp = StrParaDbl(ICMSSubstAliquotaItem.Text) / 100
'        objTributacaoPRJCTRItem.dICMSSubstAliquota = dTemp
'        objTributacaoPRJCTRItem.iICMSSubstAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSSubstValorItem.Text <> CStr(objTributacaoPRJCTRItem.dICMSSubstValor) Then
'        dTemp = StrParaDbl(ICMSSubstValorItem.Text)
'        objTributacaoPRJCTRItem.dICMSSubstValor = dTemp
'        objTributacaoPRJCTRItem.iICMSSubstValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados de IPI
'    iTemp = ComboIPITipo.ItemData(ComboIPITipo.ListIndex)
'    If iTemp <> objTributacaoPRJCTRItem.iIPITipo Then
'        objTributacaoPRJCTRItem.iIPITipo = iTemp
'        objTributacaoPRJCTRItem.iIPITipoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIBaseItem.Text <> CStr(objTributacaoPRJCTRItem.dIPIBaseCalculo) Then
'        dTemp = StrParaDbl(IPIBaseItem.Text)
'        objTributacaoPRJCTRItem.dIPIBaseCalculo = dTemp
'        objTributacaoPRJCTRItem.iIPIBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIPercRedBaseItem.Text <> CStr(objTributacaoPRJCTRItem.dIPIPercRedBase * 100) Then
'        dTemp = StrParaDbl(IPIPercRedBaseItem.Text) / 100
'        objTributacaoPRJCTRItem.dIPIPercRedBase = dTemp
'        objTributacaoPRJCTRItem.iIPIPercRedBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIAliquotaItem.Text <> CStr(objTributacaoPRJCTRItem.dIPIAliquota * 100) Then
'        dTemp = StrParaDbl(IPIAliquotaItem.Text) / 100
'        objTributacaoPRJCTRItem.dIPIAliquota = dTemp
'        objTributacaoPRJCTRItem.iIPIAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIValorItem.Text <> CStr(objTributacaoPRJCTRItem.dIPIValor) Then
'        dTemp = StrParaDbl(IPIValorItem.Text)
'        objTributacaoPRJCTRItem.dIPIValor = dTemp
'        objTributacaoPRJCTRItem.iIPIValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    TributacaoItem_GravarTela = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_GravarTela:
'
'    TributacaoItem_GravarTela = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187700)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TributacaoItem_TrazerTela(ByVal objTributacaoPRJCTRItem As Object) As Long
''Traz para a tela dados de tributacao de um item
'
'Dim iIndice As Integer
'Dim objContratoItem As ClassPRJContratoItem
'Dim lErro As Long
'Dim objTipoTribIPI As New ClassTipoTribIPI
'Dim objTipoTribICMS As New ClassTipoTribICMS
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim sNatOp As String
'
'On Error GoTo Erro_TributacaoItem_TrazerTela
'
'    giTrazendoTribItemTela = 1
'
'    NaturezaOpItem.Text = objTributacaoPRJCTRItem.sNaturezaOp
'
'    sNatOp = Trim(NaturezaOpItem.Text)
'
'    If sNatOp <> "" Then
'
'        objNaturezaOp.sCodigo = sNatOp
'        'Lê a Natureza de Operação
'        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'        If lErro <> SUCESSO And lErro <> 17958 Then gError 187899
'
'        'Se não achou a Natureza de Operação --> erro
'        If lErro <> SUCESSO Then gError 187900
'
'        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
'    Else
'        LabelDescrNatOpItem.Caption = ""
'    End If
'
'    objTributacaoTipo.iTipo = objTributacaoPRJCTRItem.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then gError 187901
'
'        TipoTributacaoItem.Text = CStr(objTributacaoPRJCTRItem.iTipoTributacao)
'        DescTipoTribItem.Caption = objTributacaoTipo.sDescricao
'
'        'Se não incide IPI
'        If objTributacaoTipo.iIPIIncide = 0 Then
'
'            ComboIPITipo.Enabled = False
'            IPIBaseItem.Enabled = False
'        Else
'
'            ComboIPITipo.Enabled = True
'            IPIBaseItem.Enabled = True
'
'        End If
'
'        'Se não incide ICMS
'        If objTributacaoTipo.iICMSIncide = 0 Then
'
'            ComboICMSTipo.Enabled = False
'            ICMSBaseItem.Enabled = False
'        Else
'
'            ComboICMSTipo.Enabled = True
'            ICMSBaseItem.Enabled = True
'
'        End If
'
'    Else
'
'        TipoTributacaoItem.Text = ""
'        DescTipoTribItem.Caption = ""
'
'    End If
'
'    'Setar dados de ICMS
'    Call Combo_Seleciona_ItemData(ComboICMSTipo, objTributacaoPRJCTRItem.iICMSTipo)
'
'    ICMSBaseItem.Text = CStr(objTributacaoPRJCTRItem.dICMSBase)
'    ICMSPercRedBaseItem.Text = CStr(objTributacaoPRJCTRItem.dICMSPercRedBase * 100)
'    ICMSAliquotaItem.Text = CStr(objTributacaoPRJCTRItem.dICMSAliquota * 100)
'    ICMSValorItem.Text = CStr(objTributacaoPRJCTRItem.dICMSValor)
'
'    'setar dados ICMS Substituicao
'    ICMSSubstBaseItem.Text = CStr(objTributacaoPRJCTRItem.dICMSSubstBase)
'    ICMSSubstAliquotaItem.Text = CStr(objTributacaoPRJCTRItem.dICMSSubstAliquota * 100)
'    ICMSSubstValorItem.Text = CStr(objTributacaoPRJCTRItem.dICMSSubstValor)
'
'    For Each objTipoTribICMS In gcolTiposTribICMS
'        If objTipoTribICMS.iTipo = objTributacaoPRJCTRItem.iICMSTipo Then Exit For
'    Next
'
'    'Se permite redução de base habilitar este campo
'    If objTipoTribICMS.iPermiteReducaoBase Then
'        ICMSPercRedBaseItem.Enabled = True
'    Else
'        'Desabilita-lo e limpa-lo em caso contrário
'        ICMSPercRedBaseItem.Enabled = False
'    End If
'
'    'Se permite aliquota habilitar este campo e valor.
'    If objTipoTribICMS.iPermiteAliquota Then
'
'        ICMSAliquotaItem.Enabled = True
'        ICMSValorItem.Enabled = True
'
'    Else
'
'        'Desabilitar os dois campos e coloca-los com zero
'        ICMSAliquotaItem.Enabled = False
'        ICMSValorItem.Enabled = False
'
'    End If
'
'    'Se permite margem de lucro habilitar campos do frame de substituicao
'    If objTipoTribICMS.iPermiteMargLucro Then
'
'        ICMSSubstBaseItem.Enabled = True
'        ICMSSubstAliquotaItem.Enabled = True
'        ICMSSubstValorItem.Enabled = True
'
'    Else
'
'        'Limpa-los e desabilita-los
'        ICMSSubstBaseItem.Enabled = False
'        ICMSSubstAliquotaItem.Enabled = False
'        ICMSSubstValorItem.Enabled = False
'
'    End If
'
'    'Setar dados de IPI
'    Call Combo_Seleciona_ItemData(ComboIPITipo, objTributacaoPRJCTRItem.iIPITipo)
'
'    IPIBaseItem.Text = CStr(objTributacaoPRJCTRItem.dIPIBaseCalculo)
'    IPIPercRedBaseItem.Text = CStr(objTributacaoPRJCTRItem.dIPIPercRedBase * 100)
'    IPIAliquotaItem.Text = CStr(objTributacaoPRJCTRItem.dIPIAliquota * 100)
'    IPIValorItem.Text = CStr(objTributacaoPRJCTRItem.dIPIValor)
'
'    For Each objTipoTribIPI In gcolTiposTribIPI
'        If objTipoTribIPI.iTipo = objTributacaoPRJCTRItem.iIPITipo Then Exit For
'    Next
'
'    'Se permite redução de base habilitar este campo
'    If objTipoTribIPI.iPermiteReducaoBase Then 'leo voltar aqui
'        IPIPercRedBaseItem.Enabled = True
'    Else
'
'        'desabilita-lo e limpa-lo em caso contrário
'        IPIPercRedBaseItem.Enabled = False
'
'    End If
'
'    'Se permite alíquota habilitar este campo e valor.
'    If objTipoTribIPI.iPermiteAliquota Then
'
'        IPIAliquotaItem.Enabled = True
'        IPIValorItem.Enabled = True
'
'    Else
'        'Desabilitar os dois campos e coloca-los com zero
'        IPIAliquotaItem.Enabled = False
'        IPIValorItem.Enabled = False
'
'    End If
'
'    giTrazendoTribItemTela = 0
'    giNatOpItemAlterado = 0
'    giTipoTributacaoItemAlterado = 0
'    giICMSBaseItemAlterado = 0
'    giICMSPercRedBaseItemAlterado = 0
'    giICMSAliquotaItemAlterado = 0
'    giICMSValorItemAlterado = 0
'    giICMSSubstBaseItemAlterado = 0
'    giICMSSubstAliquotaItemAlterado = 0
'    giICMSSubstValorItemAlterado = 0
'    giIPIBaseItemAlterado = 0
'    giIPIPercRedBaseItemAlterado = 0
'    giIPIAliquotaItemAlterado = 0
'    giIPIValorItemAlterado = 0
'
'    TributacaoItem_TrazerTela = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_TrazerTela:
'
'    giTrazendoTribItemTela = 0
'
'    TributacaoItem_TrazerTela = gErr
'
'    Select Case gErr
'
'        Case 187899, 187901
'
'        Case 187900
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", objNaturezaOp.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187701)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function TributacaoItem_InicializaTab() As Long
''deve ser chamada na entrada do tab de detalhamento dentro do tab de tributacao
'Dim lErro As Long
'Dim objContratoItem As ClassPRJContratoItem
'Dim sItem As String
'
'On Error GoTo Erro_TributacaoItem_InicializaTab
'
'    'preencher o valor de frete, seguro, descontos e outras desp no frameOutros
'    LabelValorFrete.Caption = Format(gobjContrato.dValorFrete, "Standard")
'    LabelValorDesconto.Caption = Format(gobjContrato.dValorDesconto, "Standard")
'    LabelValorSeguro.Caption = Format(gobjContrato.dValorSeguro, "Standard")
'    LabelValorOutrasDespesas.Caption = Format(gobjContrato.dValorOutrasDespesas, "Standard")
'
'    'esvaziar a combo de itens
'    ComboItensTrib.Clear
'
'    'preencher a combo de itens: com "codigo do produto - descricao"
'    For Each objContratoItem In gobjContrato.colItens
'
'        sItem = ""
'
'        If Len(Trim(objContratoItem.sProduto)) > 0 Then
'
'            lErro = Mascara_MascararProduto(objContratoItem.sProduto, sItem)
'            If lErro <> SUCESSO Then gError 187902
'
'        End If
'
'        sItem = sItem & " - " & objContratoItem.sDescProd & " (" & objContratoItem.sCodEtapa & SEPARADOR & objContratoItem.sDescEtapa & ")"
'
'        ComboItensTrib.AddItem sItem
'
'    Next
'
'    TribSobreItem.Value = True
'    Call TribSobreItem_Click
'
'    TributacaoItem_InicializaTab = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_InicializaTab:
'
'    TributacaoItem_InicializaTab = gErr
'
'    Select Case gErr
'
'        Case 187902
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187702)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoGravarTribCarga()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoGravarTribCarga
'
'    lErro = Tributacao_GravarTela()
'    If lErro <> SUCESSO Then gError 187903
'
'    'Atualiza os valores de tributação
'    lErro =  gobjTribTab.AtualizarTributacao()
'    If lErro <> SUCESSO Then gError 187904
'
'    Exit Sub
'
'Erro_BotaoGravarTribCarga:
'
'    Select Case gErr
'
'        Case 187903, 187904
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187703)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function TributacaoPRJCTR_Terminar() As Long
'
'Dim lErro As Long, objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoPRJCTR_Terminar
'
'    If Not (gobjContrato Is Nothing) Then
'        Set objTributoDoc = gobjContrato
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then gError 187905
'        Set gobjContrato = Nothing
'    End If
'
'    TributacaoPRJCTR_Terminar = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPRJCTR_Terminar:
'
'    TributacaoPRJCTR_Terminar = gErr
'
'    Select Case gErr
'
'        Case 187905
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187704)
'
'    End Select
'
'End Function
'
'Private Function CarregaTiposTrib() As Long
'
'Dim lErro As Long, sCodigo As String
'Dim objTipoTribICMS As ClassTipoTribICMS
'Dim objTipoTribIPI As ClassTipoTribIPI
'
'On Error GoTo Erro_CarregaTiposTrib
'
'    lErro = CF("TiposTribICMS_Le_Todos", gcolTiposTribICMS)
'    If lErro <> SUCESSO Then gError 187907
'
'    'Preenche ComboICMSTipo
'    For Each objTipoTribICMS In gcolTiposTribICMS
'
'        sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribICMS.iTipo)))
'        sCodigo = sCodigo & CStr(objTipoTribICMS.iTipo) & SEPARADOR & objTipoTribICMS.sDescricao
'        ComboICMSTipo.AddItem (sCodigo)
'        ComboICMSTipo.ItemData(ComboICMSTipo.NewIndex) = objTipoTribICMS.iTipo
'
'    Next
'
'    lErro = CF("TiposTribIPI_Le_Todos", gcolTiposTribIPI)
'    If lErro <> SUCESSO Then gError 187908
'
'    'Preenche ComboIPITipo
'    For Each objTipoTribIPI In gcolTiposTribIPI
'
'        sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribIPI.iTipo)))
'        sCodigo = sCodigo & CStr(objTipoTribIPI.iTipo) & SEPARADOR & objTipoTribIPI.sDescricao
'        ComboIPITipo.AddItem (sCodigo)
'        ComboIPITipo.ItemData(ComboIPITipo.NewIndex) = objTipoTribIPI.iTipo
'
'    Next
'
'    CarregaTiposTrib = SUCESSO
'
'    Exit Function
'
'Erro_CarregaTiposTrib:
'
'    CarregaTiposTrib = gErr
'
'    Select Case gErr
'
'        Case 187907, 187908
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187705)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function AtualizarTributacao() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_AtualizarTributacao
'
'    If Not (gobjContrato Is Nothing) Then
'
'        'Atualiza os impostos
'        lErro = gobjTributacao.AtualizaImpostos(gobjContrato, giRecalculandoTributacao)
'        If lErro <> SUCESSO Then gError 187909
'
'        'joga dados do obj atualizado p/a tela
'        lErro = Carrega_Tab_Tributacao(gobjContrato)
'        If lErro <> SUCESSO Then gError 187910
'
'    End If
'
'    AtualizarTributacao = SUCESSO
'
'    Exit Function
'
'Erro_AtualizarTributacao:
'
'    AtualizarTributacao = gErr
'
'    Select Case gErr
'
'        Case 187909, 187910
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187706)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Tributacao_Inclusao_Item_Grid(ByVal iLinha As Integer) As Long
'
''trata a inclusao de uma linha de item no grid
'Dim lErro As Long
'Dim objTributoDocItem As ClassTributoDocItem
'Dim objContratoItem As ClassPRJContratoItem
'On Error GoTo Erro_Tributacao_Inclusao_Item_Grid
'
'    lErro = Move_GridItem_Memoria(gobjContrato, iLinha)
'    If lErro <> SUCESSO Then gError 187911
'
'    Set objContratoItem = gobjContrato.colItens.Item(iLinha)
'    Set objTributoDocItem = objContratoItem
'
'    lErro = objTributoDocItem.Ativar(gobjContrato)
'    If lErro <> SUCESSO Then gError 187912
'
'    Tributacao_Inclusao_Item_Grid = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Inclusao_Item_Grid:
'
'    Tributacao_Inclusao_Item_Grid = gErr
'
'    Select Case gErr
'
'        Case 187911, 187912
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187707)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Tributacao_Remover_Item_Grid(iLinha As Integer) As Long
''trata a exclusao de uma linha de item no grid
'Dim objContratoItem As ClassPRJContratoItem, objTributoDocItem As ClassTributoDocItem
'
'        Set objContratoItem = gobjContrato.colItens(iLinha)
'        Set objTributoDocItem = objContratoItem
'        Call objTributoDocItem.Desativar
'        Call gobjContrato.RemoverItem(iLinha)
'
'End Function
'
'Function Tributacao_Alteracao_Item_Grid(iIndice As Integer) As Long
''trata a alteracao de uma linha de item no grid
'
'Dim lErro As Long, sProduto As String, iPreenchido As Integer
'Dim objContratoItem As ClassPRJContratoItem
'
'On Error GoTo Erro_Tributacao_Alteracao_Item_Grid
'
'    Set objContratoItem = gobjContrato.colItens.Item(iIndice)
'
'    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then
'
'        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 187913
'
'        objContratoItem.sProduto = sProduto
'
'    End If
'
'    objContratoItem.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
'
'    objContratoItem.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'    objContratoItem.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'    objContratoItem.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
'
'    Tributacao_Alteracao_Item_Grid = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Alteracao_Item_Grid:
'
'    Tributacao_Alteracao_Item_Grid = gErr
'
'    Select Case gErr
'
'        Case 187913
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187708)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Valida_Tributacao_Gravacao() As Long
'
'Dim lErro As Long
'Dim objContratoItem As ClassPRJContratoItem
'Dim iIndice As Integer, dtDataRef As Date
'Dim vbResult As VbMsgBoxResult
'
'On Error GoTo Erro_Valida_Tributacao_Gravacao
'
'    If gobjContrato.objTributacaoPRJCTR.iTipoTributacao = 0 Then gError 187914
'
'    If gobjContrato.lCliente = 0 Then
'        vbResult = Rotina_Aviso(vbYesNo, "AVISO_CLIENTE_CONTRATO_NAO_PREENCHIDO")
'        If vbResult = vbNo Then gError 187915
'    End If
'
'    dtDataRef = gobjContrato.dtData
'
'    iIndice = 0
'
'    For Each objContratoItem In gobjContrato.colItens
'
'        iIndice = iIndice + 1
'        If Len(Trim(objContratoItem.objTributacaoPRJCTRItem.sNaturezaOp)) = 0 Then gError 187916
'        If objContratoItem.objTributacaoPRJCTRItem.iTipoTributacao = 0 Then gError 187917
'        If Natop_ErroTamanho(dtDataRef, objContratoItem.objTributacaoPRJCTRItem.sNaturezaOp) Then gError 187927
'
'    Next
'
'    If Len(Trim(gobjContrato.objTributacaoPRJCTR.objTributacaoDesconto.sNaturezaOp)) = 0 Then gError 187918
'    If gobjContrato.objTributacaoPRJCTR.objTributacaoDesconto.iTipoTributacao = 0 Then gError 187919
'
'    If Len(Trim(gobjContrato.objTributacaoPRJCTR.objTributacaoFrete.sNaturezaOp)) = 0 Then gError 187920
'    If gobjContrato.objTributacaoPRJCTR.objTributacaoFrete.iTipoTributacao = 0 Then gError 187921
'
'    If Len(Trim(gobjContrato.objTributacaoPRJCTR.objTributacaoOutras.sNaturezaOp)) = 0 Then gError 187922
'    If gobjContrato.objTributacaoPRJCTR.objTributacaoOutras.iTipoTributacao = 0 Then gError 187923
'
'    If Len(Trim(gobjContrato.objTributacaoPRJCTR.objTributacaoSeguro.sNaturezaOp)) = 0 Then gError 187924
'    If gobjContrato.objTributacaoPRJCTR.objTributacaoSeguro.iTipoTributacao = 0 Then gError 187925
'
'    If Natop_ErroTamanho(dtDataRef, gobjContrato.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjContrato.objTributacaoPRJCTR.objTributacaoDesconto.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjContrato.objTributacaoPRJCTR.objTributacaoFrete.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjContrato.objTributacaoPRJCTR.objTributacaoOutras.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjContrato.objTributacaoPRJCTR.objTributacaoSeguro.sNaturezaOp) Then gError 187926
'
'    Valida_Tributacao_Gravacao = SUCESSO
'
'    Exit Function
'
'Erro_Valida_Tributacao_Gravacao:
'
'    Valida_Tributacao_Gravacao = gErr
'
'    Select Case gErr
'
'        Case 187914
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)
'
'        Case 187915
'
'        Case 187916
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ITEM_TRIBUTACAO_NAO_PREENCHIDA", iIndice)
'
'        Case 187917
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
'
'        Case 187918
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESCONTO_NAO_PRENCHIDA", gErr)
'
'        Case 187919
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESCONTO_NAO_PREENCHIDO", gErr)
'
'        Case 187920
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_FRETE_NAO_PRENCHIDA", gErr)
'
'        Case 187921
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_FRETE_NAO_PREENCHIDO", gErr)
'
'        Case 187922
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESPESAS_NAO_PRENCHIDA", gErr)
'
'        Case 187923
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESPESAS_NAO_PREENCHIDO", gErr)
'
'        Case 187924
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SEGURO_NAO_PRENCHIDA", gErr)
'
'        Case 187925
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_SEGURO_NAO_PREENCHIDO", gErr)
'
'        Case 187926, 187927
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_TAMANHO_INCORRETO", Err)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187709)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub PISRetido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iPISRetidoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PISRetido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValor As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_PISRetido_Validate
'
'    If iPISRetidoAlterado = 0 Then Exit Sub
'
'    'Verifica se foi preenchido
'    If Len(Trim(PISRetido.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(PISRetido.Text)
'        If lErro <> SUCESSO Then gError 187928
'
'        dValor = StrParaDbl(PISRetido.Text)
'
'        PISRetido.Text = Format(dValor, "Standard")
'
'        dValorTotal = StrParaDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 187929
'
'    End If
'
'    Call BotaoGravarTrib
'
'    iPISRetidoAlterado = 0
'
'    Exit Sub
'
'Erro_PISRetido_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187928
'
'        Case 187929
'            Call Rotina_Erro(vbOKOnly, "ERRO_PIS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187710)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub COFINSRetido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iCOFINSRetidoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub COFINSRetido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValor As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_COFINSRetido_Validate
'
'    If iCOFINSRetidoAlterado = 0 Then Exit Sub
'
'    'Verifica se foi preenchido
'    If Len(Trim(COFINSRetido.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(COFINSRetido.Text)
'        If lErro <> SUCESSO Then gError 187930
'
'        dValor = StrParaDbl(COFINSRetido.Text)
'
'        COFINSRetido.Text = Format(dValor, "Standard")
'
'        dValorTotal = StrParaDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 187931
'
'    End If
'
'    Call BotaoGravarTrib
'
'    iCOFINSRetidoAlterado = 0
'
'    Exit Sub
'
'Erro_COFINSRetido_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187930
'
'        Case 187931
'            Call Rotina_Erro(vbOKOnly, "ERRO_COFINS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187711)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub CSLLRetido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iCSLLRetidoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub CSLLRetido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValor As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_CSLLRetido_Validate
'
'    If iCSLLRetidoAlterado = 0 Then Exit Sub
'
'    'Verifica se foi preenchido
'    If Len(Trim(CSLLRetido.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(CSLLRetido.Text)
'        If lErro <> SUCESSO Then gError 187932
'
'        dValor = StrParaDbl(CSLLRetido.Text)
'
'        CSLLRetido.Text = Format(dValor, "Standard")
'
'        dValorTotal = StrParaDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 187933
'
'    End If
'
'    Call BotaoGravarTrib
'
'    iCSLLRetidoAlterado = 0
'
'    Exit Sub
'
'Erro_CSLLRetido_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187932
'
'        Case 187933
'            Call Rotina_Erro(vbOKOnly, "ERRO_CSLL_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187712)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ISSRetido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iISSRetidoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ISSRetido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValor As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_ISSRetido_Validate
'
'    If iISSRetidoAlterado = 0 Then Exit Sub
'
'    'Verifica se foi preenchido
'    If Len(Trim(ISSRetido.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(ISSRetido.Text)
'        If lErro <> SUCESSO Then gError 187934
'
'        dValor = StrParaDbl(ISSRetido.Text)
'
'        ISSRetido.Text = Format(dValor, "Standard")
'
'        dValorTotal = StrParaDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 187935
'
'    End If
'
'    Call BotaoGravarTrib
'
'    iISSRetidoAlterado = 0
'
'    Exit Sub
'
'Erro_ISSRetido_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 187934
'
'        Case 187935
'            Call Rotina_Erro(vbOKOnly, "ERRO_ISS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187713)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 187936

    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

    'Faz o Tratamento do produto
    lErro = Produto_Saida_Celula()
    If lErro <> SUCESSO Then
    
        If Not (Me.ActiveControl Is Produto) Then
        
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""
        
        End If
    
        gError 187937
        
    End If
    
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
            
        Case 187936
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 187937

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187714)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then

        sProduto1 = Produto.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 187938

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 187939

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 187938
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 187939
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187715)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEtapa_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_objEventoEtapa_evSelecao

    Set objEtapa = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    GridItens.TextMatrix(GridItens.Row, iGrid_Etapa_Col) = objEtapa.sCodigo
    
    If Not (Me.ActiveControl Is Etapa) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_DescEtapa_Col) = objEtapa.sDescricao
        
        'Se o produto não está preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = "un"
            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(1)
        End If
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError 189369
        
        End If
        
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoEtapa_evSelecao:

    Select Case gErr
    
        Case 189369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187716)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEtapas_Click()

Dim objEtapa As New ClassPRJEtapas
Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoEtapas_Click

    If gobjProjeto Is Nothing Then gError 187940

    If Me.ActiveControl Is Etapa Then

        objEtapa.sCodigo = Etapa.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 187941

        objEtapa.sCodigo = GridItens.TextMatrix(GridItens.Row, iGrid_Etapa_Col)

    End If
    
    colSelecao.Add gobjProjeto.lNumIntDoc

    Call Chama_Tela("PRJEtapasLista", colSelecao, objEtapa, objEventoEtapa, , "Código", "NumIntDocPRJ = ?")

    Exit Sub

Erro_BotaoEtapas_Click:

    Select Case gErr
    
        Case 187940
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)

        Case 187941
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187717)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then gError 187942

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 187943

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)

    'Se o Tipo estiver preenchido
    If objCliente.iTipo > 0 Then
        objTipoCliente.iCodigo = objCliente.iTipo
        'Lê o Tipo de Cliente
        lErro = CF("TipoCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 187944
    End If

    'Guarda o valor do desconto do cliente
    If objCliente.dDesconto > 0 Then
        gdDesconto = objCliente.dDesconto
    ElseIf objTipoCliente.dDesconto > 0 Then
        gdDesconto = objTipoCliente.dDesconto
    Else
        gdDesconto = 0
    End If

    'para fazer valer o que veio do bd
    giValorDescontoManual = 1

    giClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case 187942, 187943, 187944

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187945)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 187946

    If lErro = 17660 Then gError 187947

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 187946

        Case 187947
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187948)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerarArq_Click()

Dim lErro As Long
Dim objContrato As New ClassPRJContratos
Dim objContratoBD As New ClassPRJContratos
Dim objCliente As New ClassCliente
Dim objEndereco As New ClassEndereco
Dim objReceb As New ClassPRJRecebPagto
Dim colEnderecos As New ColEndereco
Dim objTela As Object
Dim sDiretorio As String

On Error GoTo Erro_BotaoGerarArq_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 189005
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 189006
    If Len(Trim(Modelo.Text)) = 0 Then gError 189007
    
    If InStr(1, NomeArquivo.Text, ".") = 0 Then
        NomeArquivo.Text = NomeArquivo.Text & ".doc"
    End If
    
    If right(NomeDiretorio.Text, 1) = "\" Or right(NomeDiretorio.Text, 1) = "/" Then
        sDiretorio = NomeDiretorio.Text & NomeArquivo.Text
    Else
        sDiretorio = NomeDiretorio.Text & "\" & NomeArquivo.Text
    End If
    
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 187950
    
    objContratoBD.lNumIntDocPRJ = gobjProjeto.lNumIntDoc
    objContratoBD.sCodigo = objContrato.sCodigo

    'Lê o Projetos que está sendo Passado
    lErro = CF("PRJContratos_Le", objContratoBD)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187969
    
    If objContratoBD.lNumIntDoc <> 0 Then
    
        objReceb.lNumIntDocContrato = objContratoBD.lNumIntDoc
        objReceb.iTipo = PRJ_TIPO_RECEB
        objReceb.lNumIntDocPRJ = gobjProjeto.lNumIntDoc
        
        lErro = CF("PRJRecebPagto_Le_Contrato", objReceb)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187970
        
        lErro = CF("PRJRecebPagtoRegras_Le", objReceb)
        If lErro <> SUCESSO Then gError 187994
        
        Set objTela = gobjProjeto
    
        lErro = CF("RecebPagto_Calcula_Regras", objTela, objReceb)
        If lErro <> SUCESSO Then gError 187990
        
    End If
    
    If objContrato.lCliente <> 0 Then
    
        objCliente.lCodigo = objContrato.lCliente
    
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 187967

        'Lê os dados dos tres tipos de enderecos
        lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
        If lErro <> SUCESSO Then gError 187968
        
        Set objEndereco = colEnderecos.Item(1)
    
    End If

    lErro = Gera_Arquivo_Doc(sDiretorio, objCliente, gobjProjeto, gobjProposta, objContrato, objEndereco, objReceb)
    If lErro <> SUCESSO Then gError 189020
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGerarArq_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 187950, 187967, 187968, 187969, 187970, 187990, 187994, 189020
        
        Case 189005
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
            NomeDiretorio.SetFocus
        
        Case 189006
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
            NomeArquivo.SetFocus
        
        Case 189007
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_INFORMADO", gErr)
            Modelo.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187949)

    End Select

    Exit Sub

End Sub

Private Function Gera_Arquivo_Doc(ByVal sDiretorio As String, ByVal objCliente As ClassCliente, ByVal objProjeto As ClassProjetos, ByVal objProposta As ClassPRJPropostas, ByVal objContrato As ClassPRJContratos, ByVal objEndereco As ClassEndereco, ByVal objReceb As ClassPRJRecebPagto)

Dim lErro As Long
'Dim objWord As Object 'Word.Application
'Dim objDoc As Object 'Word.Document
'Dim objCampoForm As Object 'Word.FormField
Dim objMnemonicoMala As ClassMnemonicoMalaDireta
Dim vValor As Variant
Dim objEtapa As ClassPRJEtapas
Dim objContratoEtapa As ClassPRJContratoEtapa
Dim objContratoItem As ClassPRJContratoItem
Dim bExibe As Boolean
Dim iContador As Integer
Dim objEtapaIP As ClassPRJEtapaItensProd
Dim bAchou As Boolean
Dim iColunas As Integer
Dim objCondPagto As ClassCondicaoPagto
Dim objRecebRegra As New ClassPRJRecebPagtoRegras
Dim sMesNome As String
Dim objTelaProjeto As Object
Dim objTelaGrafico As New ClassTelaGrafico
Dim objTelaCronograma As Object
Dim sNomeFigura As String
Dim objFSO As New FileSystemObject
Dim objProjetoAux As ClassProjetos
Dim iIndice As Integer ', dVersaoWord As Double
Dim objWordApp As New ClassWordApp, lNumFormFields As Long, lIndiceFF As Long, sNomeFF As String

On Error GoTo Erro_Gera_Arquivo_Doc

    'Set objWord = CreateObject("Word.Application")
    
    lErro = objWordApp.Abrir
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    dVersaoWord = 0
'    If IsNumeric(objWord.Version) Then
'        dVersaoWord = StrParaDbl(Replace(objWord.Version, ".", ","))
'    End If
'
'    If dVersaoWord < 15 Then
'        Set objDoc = objWord.Documents.Open(Modelo.Text, , True)
'    Else
'        Set objDoc = objWord.Documents.Open(Modelo.Text)
'    End If

    lErro = objWordApp.Abrir_Doc(Modelo.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'For Each objCampoForm In objDoc.FormFields
    
    lNumFormFields = objWordApp.Qtde_FormFields()
    
    For lIndiceFF = lNumFormFields To 1 Step -1
    
'        Call objCampoForm.Select
    
        'lErro = objWordApp.FormField_Seleciona(lIndiceFF)
        'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
        vValor = ""
        Set objMnemonicoMala = New ClassMnemonicoMalaDireta
    
        lErro = objWordApp.FormField_Obtem_Nome(lIndiceFF, sNomeFF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objMnemonicoMala.sMnemonico = sNomeFF 'objCampoForm.Name
        objMnemonicoMala.iTipo = MNEMONICO_MALADIRETA_TIPO_CONTRATO
        
        lErro = CF("MnemonicoMalaDireta_Le", objMnemonicoMala)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187956
        
        If lErro <> SUCESSO Then
        
            objMnemonicoMala.iTipo = MNEMONICO_MALADIRETA_TIPO_PROPOSTA
            
            lErro = CF("MnemonicoMalaDireta_Le", objMnemonicoMala)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187956

        End If
        
        If lErro <> SUCESSO Then gError 187957
    
        Select Case objMnemonicoMala.iTipoObj
        
            Case MNEMONICO_MALADIRETA_TIPOOBJ_CLIENTE
            
                lErro = Critica_ObjetoAtributo(objCliente, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError 187958

            Case MNEMONICO_MALADIRETA_TIPOOBJ_PROJETO

                lErro = Critica_ObjetoAtributo(objProjeto, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError 187959

            Case MNEMONICO_MALADIRETA_TIPOOBJ_CONTRATO

                lErro = Critica_ObjetoAtributo(objContrato, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError 187960

            Case MNEMONICO_MALADIRETA_TIPOOBJ_ESCOPO

                lErro = Critica_ObjetoAtributo(objProjeto.objEscopo, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError 187961

            Case MNEMONICO_MALADIRETA_TIPOOBJ_ENDERECO_CLIENTE

                lErro = Critica_ObjetoAtributo(objEndereco, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError 187962
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_RECEBIMENTO

                lErro = Critica_ObjetoAtributo(objReceb, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError 187963

            Case MNEMONICO_MALADIRETA_TIPOOBJ_PROPOSTA

                lErro = Critica_ObjetoAtributo(objProposta, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError 187963

            Case MNEMONICO_MALADIRETA_TIPOOBJ_OUTROS
            
                Select Case objMnemonicoMala.sMnemonico
                
                    Case "Dia_Agora"
                        vValor = Format(Day(Now), "00")
                        
                    Case "Mes_Agora"
                        vValor = Format(Month(Now), "00")
                
                    Case "Ano_Agora"
                        vValor = Format(Year(Now), "0000")
                
                    Case "Hora_Agora"
                        vValor = Format(Now, "HH:MM:SS")
                
                    Case "Data_Agora"
                        vValor = Format(Now, "DD/MM/YYYY")
                
                    Case "Mes_Agora_Nome"
                        Call MesNome(Month(Now), sMesNome)
                        vValor = sMesNome
                
                    Case "Lista_Etapas_PRJ"
                                                            
'                        Call DOC_Cria_Tabela(objDoc, objWord, objProjeto.colEtapas.Count + 1, 2)
'                        Call DOC_Insere_Cabec_Tabela(objWord, "Etapa", "Descrição")
'
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, objProjeto.colEtapas.Count + 1, 2)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Etapa", "Descrição")
                
                        For Each objEtapa In objProjeto.colEtapas
                            'Call DOC_Insere_Valores_Tabela(objWord, objEtapa.sCodigo, objEtapa.sDescricao)
                            Call objWordApp.DOC_Insere_Valores_Tabela(objEtapa.sCodigo, objEtapa.sDescricao)
                        Next
                    
                    Case "Lista_Etapas_PRJ_Dat"
                    
                        'Call DOC_Cria_Tabela(objDoc, objWord, objProjeto.colEtapas.Count + 1, 4)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Etapa", "Descrição", "Início", "Fim")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, objProjeto.colEtapas.Count + 1, 4)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Etapa", "Descrição", "Início", "Fim")
                
                        For Each objEtapa In objProjeto.colEtapas
                            'Call DOC_Insere_Valores_Tabela(objWord, objEtapa.sCodigo, objEtapa.sDescricao, Formata_Data(objEtapa.dtDataInicio), Formata_Data(objEtapa.dtDataFim))
                            Call objWordApp.DOC_Insere_Valores_Tabela(objEtapa.sCodigo, objEtapa.sDescricao, Formata_Data(objEtapa.dtDataInicio), Formata_Data(objEtapa.dtDataFim))
                        Next
                
                    Case "Lista_Etapas_Imp"
                        
                        iContador = 0
                        For Each objEtapa In objProjeto.colEtapas
                            bAchou = False
                            For Each objContratoEtapa In objContrato.colEtapas
                                If objContratoEtapa.lNumIntDocEtapa = objEtapa.lNumIntDoc Then
                                    bAchou = True
                                    Exit For
                                End If
                            Next
                            If bAchou Then
                                If objContratoEtapa.iImprimir = MARCADO Then
                                    iContador = iContador + 1
                                End If
                            End If
                        Next
                        
                        'Call DOC_Cria_Tabela(objDoc, objWord, iContador + 1, 3)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Etapa", "Descrição", "Observação")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, iContador + 1, 3)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Etapa", "Descrição", "Observação")
                
                        For Each objEtapa In objProjeto.colEtapas
                            bAchou = False
                            For Each objContratoEtapa In objContrato.colEtapas
                                If objContratoEtapa.lNumIntDocEtapa = objEtapa.lNumIntDoc Then
                                    bAchou = True
                                    Exit For
                                End If
                            Next
                        
                            If bAchou Then
                                If objContratoEtapa.iImprimir = MARCADO Then
                                    'Call DOC_Insere_Valores_Tabela(objWord, objEtapa.sCodigo, objContratoEtapa.sDescricao, objContratoEtapa.sObservacao)
                                    Call objWordApp.DOC_Insere_Valores_Tabela(objEtapa.sCodigo, objContratoEtapa.sDescricao, objContratoEtapa.sObservacao)
                                End If
                            End If
                        Next
                
                    Case "Lista_Etapas_Prop", "Lista_Etap_Prop_Marc"
                    
                        iContador = 0
                        For Each objContratoEtapa In objContrato.colEtapas
                            If objMnemonicoMala.sMnemonico = "Lista_Etap_Prop_Marc" Then
                                If objContratoEtapa.iSelecionado = MARCADO Then
                                    iContador = iContador + 1
                                End If
                            Else
                                iContador = iContador + 1
                            End If
                        Next
                        
                        iColunas = 2
                        If ExibirCustoCalc.Value = vbChecked Then
                            iColunas = iColunas + 1
                        End If
                        If ExibirCustoInf.Value = vbChecked Then
                            iColunas = iColunas + 1
                        End If
                        If ExibirPreco.Value = vbChecked Then
                            iColunas = iColunas + 1
                        End If
                        
'                        Call DOC_Cria_Tabela(objDoc, objWord, iContador + 1, iColunas)
'                        Call DOC_Insere_Cabec_Tabela(objWord, "Etapa", "Descrição")

                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, iContador + 1, iColunas)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Etapa", "Descrição")
                        
                        If ExibirCustoCalc.Value = vbChecked Then
                            'Call DOC_Insere_Cabec_Tabela(objWord, "Custo Calc.")
                            Call objWordApp.DOC_Insere_Cabec_Tabela("Custo Calc.")
                        End If
                        If ExibirCustoInf.Value = vbChecked Then
                            'Call DOC_Insere_Cabec_Tabela(objWord, "Custo Inf.")
                            Call objWordApp.DOC_Insere_Cabec_Tabela("Custo Inf.")
                        End If
                        If ExibirPreco.Value = vbChecked Then
                            'Call DOC_Insere_Cabec_Tabela(objWord, "Preço")
                            Call objWordApp.DOC_Insere_Cabec_Tabela("Preço")
                        End If
                                    
                        For Each objContratoEtapa In objContrato.colEtapas
                            For Each objEtapa In objProjeto.colEtapas
                                If objContratoEtapa.lNumIntDocEtapa = objEtapa.lNumIntDoc Then
                                    If objContratoEtapa.lNumIntDocEtapaItemProd <> 0 Then
                                        For Each objEtapaIP In objEtapa.colItensProduzidos
                                            If objContratoEtapa.lNumIntDocEtapaItemProd = objEtapaIP.lNumIntDoc Then
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    Exit For
                                End If
                            Next
                                                   
                            If objMnemonicoMala.sMnemonico = "Lista_Etap_Prop_Marc" Then
                                If objContratoEtapa.iSelecionado = MARCADO Then
                                    bExibe = True
                                Else
                                    bExibe = False
                                End If
                            Else
                                bExibe = True
                            End If
                            
                            If bExibe Then
                                'Call DOC_Insere_Valores_Tabela(objWord, objEtapa.sCodigo, objEtapa.sDescricao)
                                Call objWordApp.DOC_Insere_Valores_Tabela(objEtapa.sCodigo, objEtapa.sDescricao)
                                If ExibirCustoCalc.Value = vbChecked Then
                                    If objContratoEtapa.lNumIntDocEtapaItemProd = 0 Then
                                        'Call DOC_Insere_Valores_Tabela2(objWord, Format(objEtapa.dCustoCalcPrev, "STANDARD"))
                                        Call objWordApp.DOC_Insere_Valores_Tabela2(Format(objEtapa.dCustoCalcPrev, "STANDARD"))
                                    Else
                                        'Call DOC_Insere_Valores_Tabela2(objWord, Format(0, "STANDARD"))
                                        Call objWordApp.DOC_Insere_Valores_Tabela2(Format(0, "STANDARD"))
                                    End If
                                End If
                                If ExibirCustoInf.Value = vbChecked Then
                                    'Call DOC_Insere_Valores_Tabela2(objWord, Format(objContratoEtapa.dCustoInformado, "STANDARD"))
                                    Call objWordApp.DOC_Insere_Valores_Tabela2(Format(objContratoEtapa.dCustoInformado, "STANDARD"))
                                End If
                                If ExibirPreco.Value = vbChecked Then
                                    'Call DOC_Insere_Valores_Tabela2(objWord, Format(objContratoEtapa.dPreco, "STANDARD"))
                                    Call objWordApp.DOC_Insere_Valores_Tabela2(Format(objContratoEtapa.dPreco, "STANDARD"))
                                End If
                                    
                            End If
                        
                        Next
                    
                    Case "Condicao_Pagto"

                        'Call DOC_Cria_Tabela(objDoc, objWord, objReceb.colRegras.Count + 1, 5)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Data Prevista", "Percentual", "Valor", "Cond. Pagto", "Observação")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, objReceb.colRegras.Count + 1, 5)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Data Prevista", "Percentual", "Valor", "Cond. Pagto", "Observação")
                
                        For Each objRecebRegra In objReceb.colRegras
                        
                            Set objCondPagto = New ClassCondicaoPagto
                            
                            objCondPagto.iCodigo = objRecebRegra.iCondPagto
                        
                            'Tenta ler CondicaoPagto com esse código no BD
                            lErro = CF("CondicaoPagto_Le", objCondPagto)
                            If lErro <> SUCESSO And lErro <> 19205 Then gError 187993
                        
                            'Call DOC_Insere_Valores_Tabela(objWord, Formata_Data(objRecebRegra.dtRegraValor), Format(objRecebRegra.dPercentual, "PERCENT"), Format(objRecebRegra.dPercentual * objReceb.dValor, "STANDARD"), objRecebRegra.iCondPagto & SEPARADOR & objCondPagto.sDescReduzida, objRecebRegra.sObservacao)
                            Call objWordApp.DOC_Insere_Valores_Tabela(Formata_Data(objRecebRegra.dtRegraValor), Format(objRecebRegra.dPercentual, "PERCENT"), Format(objRecebRegra.dPercentual * objReceb.dValor, "STANDARD"), objRecebRegra.iCondPagto & SEPARADOR & objCondPagto.sDescReduzida, objRecebRegra.sObservacao)
                        
                        Next
                        
                    Case "Lista_Itens_Prop"
                    
                        'Call DOC_Cria_Tabela(objDoc, objWord, objContrato.colItens.Count + 1, 9)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Etapa", "Descrição Etapa", "Produto", "Descrição Produto", "UM", "Quantidade", "Preço Unitário", "Desconto", "Preço Total")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, objContrato.colItens.Count + 1, 9)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Etapa", "Descrição Etapa", "Produto", "Descrição Produto", "UM", "Quantidade", "Preço Unitário", "Desconto", "Preço Total")
                
                        For Each objContratoItem In objContrato.colItens
                            'Call DOC_Insere_Valores_Tabela(objWord, objContratoItem.sCodEtapa, objContratoItem.sDescEtapa, objContratoItem.sProduto, objContratoItem.sDescProd, objContratoItem.sUM, Formata_Estoque(objContratoItem.dQuantidade), Format(objContratoItem.dPrecoUnitario, "STANDARD"), Format(objContratoItem.dValorDesconto, "STANDARD"), Format(objContratoItem.dPrecoTotal, "STANDARD"))
                            Call objWordApp.DOC_Insere_Valores_Tabela(objContratoItem.sCodEtapa, objContratoItem.sDescEtapa, objContratoItem.sProduto, objContratoItem.sDescProd, objContratoItem.sUM, Formata_Estoque(objContratoItem.dQuantidade), Format(objContratoItem.dPrecoUnitario, "STANDARD"), Format(objContratoItem.dValorDesconto, "STANDARD"), Format(objContratoItem.dPrecoTotal, "STANDARD"))
                        Next
                        
                    Case "Cronograma_Grafico"
                    
                        Set objTelaProjeto = CreateObject("TelasPRJ.Projetos")
                        Set objTelaCronograma = CreateObject("TelasPCP.TelaGrafico")
                        
                        'sNomeFigura =  objDoc.Path & "\" & gsUsuario & Format(Now, "YYYYMMDDHHMMSS") & ".bmp"
                        sNomeFigura = objFSO.GetParentFolderName(Modelo.Text) & "\" & gsUsuario & Format(Now, "YYYYMMDDHHMMSS") & ".bmp"
                        
                        Set objProjetoAux = New ClassProjetos
                        
                        objProjetoAux.lNumIntDoc = gobjProjeto.lNumIntDoc
                        objProjetoAux.sCodigo = gobjProjeto.sCodigo
                        objProjetoAux.iFilialEmpresa = gobjProjeto.iFilialEmpresa
                        
                        Call objTelaProjeto.Form_Load
                        Call objTelaProjeto.Trata_Parametros(objProjetoAux)
                        objTelaGrafico.iZOOM = ZOOM_50
                        Call objTelaProjeto.Atualiza_Cronograma(objTelaGrafico, False, sNomeFigura)
                        
                        Call objTelaCronograma.Trata_Parametros(objTelaGrafico)
                        Call objTelaCronograma.BotaoImprimir_Click
                        
                        'objWord.selection.GoToNext wdGoToField
                        For iIndice = 1 To objTelaGrafico.iNumFiguras
                            If iIndice = 1 Then
                                'objWord.selection.InlineShapes.AddPicture sNomeFigura, False, True
                                Call objWordApp.DOC_Insere_Figura(sNomeFigura)
                                Call objFSO.DeleteFile(sNomeFigura)
                            Else
                                'objWord.selection.InlineShapes.AddPicture left(sNomeFigura, Len(sNomeFigura) - 4) & SEPARADOR & CStr(iIndice) & ".bmp", False, True
                                Call objWordApp.DOC_Insere_Figura(left(sNomeFigura, Len(sNomeFigura) - 4) & SEPARADOR & CStr(iIndice) & ".bmp")
                                Call objFSO.DeleteFile(left(sNomeFigura, Len(sNomeFigura) - 4) & SEPARADOR & CStr(iIndice) & ".bmp")
                            End If
                        Next
                                                
                    Case Else
                        gError 187964
                End Select

            Case Else
                gError 187965
                
        End Select
        
        Select Case UCase(TypeName(vValor))
        
            Case "DATE"
                vValor = Formata_Data(vValor)
        
            Case "DOUBLE"
                vValor = Format(vValor, "STANDARD")
        
        End Select
        
        'objCampoForm.Range.Text = vValor
        
        lErro = objWordApp.FormField_Preenche_Valor(lIndiceFF, vValor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next
    
    'objDoc.SaveAs sDiretorio
    
    'Temporário
    'objWord.Visible = True
    
    lErro = objWordApp.Salvar(sDiretorio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If UCase(right(sDiretorio, 3)) = "PDF" Or UCase(right(sDiretorio, 4)) = "PDF""" Then
        lErro = objWordApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Call ShellExecute(hWnd, "open", sDiretorio, vbNullString, vbNullString, 1)
    Else
   
        lErro = objWordApp.Mudar_Visibilidade(True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Gera_Arquivo_Doc = SUCESSO

    Exit Function

Erro_Gera_Arquivo_Doc:

    Gera_Arquivo_Doc = False

    'Call objDoc.Close(False)
    Call objWordApp.Fechar

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 187956, 187993

        Case 187957
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALADIRETA_NAO_CADASTRADO", gErr, objMnemonicoMala.sMnemonico, objMnemonicoMala.iTipo)
        
        Case 187958 To 187963
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALA_ATRIBUTO_INVALIDO", gErr, objMnemonicoMala.sNomeCampoObj, objMnemonicoMala.iTipo)

        Case 187964
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALA_NAO_TRATADO", gErr, objMnemonicoMala.sNomeCampoObj)

        Case 187965
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALA_TIPO_INVALIDO", gErr, objMnemonicoMala.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187955)

    End Select

    Exit Function
    
End Function

Private Sub BotaoMnemonicos_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objMnemonicoMala As New ClassMnemonicoMalaDireta

On Error GoTo Erro_BotaoMnemonicos_Click

    colSelecao.Add MNEMONICO_MALADIRETA_TIPO_PROPOSTA
    colSelecao.Add MNEMONICO_MALADIRETA_TIPO_CONTRATO

    Call Chama_Tela("MnemonicoMalaDiretaLista", colSelecao, objMnemonicoMala, Nothing, "Tipo = ? OR Tipo = ?")

    Exit Sub

Erro_BotaoMnemonicos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187955)

    End Select

    Exit Sub
    
End Sub

Private Function Critica_ObjetoAtributo(ByVal objObj As Object, ByVal sAtributo As String, vValor As Variant) As Long

On Error GoTo Erro_Critica_ObjetoAtributo
    
    vValor = CallByName(objObj, sAtributo, VbGet)

    Critica_ObjetoAtributo = SUCESSO

    Exit Function

Erro_Critica_ObjetoAtributo:

    Critica_ObjetoAtributo = gErr

    Exit Function

End Function

Private Function Formata_Data(ByVal dtData As Date) As String
    
    If dtData = DATA_NULA Then
        Formata_Data = ""
    Else
        Formata_Data = Format(dtData, "dd/mm/yyyy")
    End If

    Exit Function

End Function

'Private Sub DOC_Cria_Tabela(ByVal objDoc As Object, ByVal objWord As Object, ByVal iNumLinhas As Integer, ByVal iNumColunas As Integer)
'
'    'objWord.selection.GoToNext wdGoToField
'    objWord.selection.MoveRight wdCharacter, 1
'
'    objDoc.Tables.Add objWord.selection.Range, iNumLinhas, iNumColunas, wdWord9TableBehavior, wdAutoFitContent
'    objWord.selection.Tables(1).ApplyStyleHeadingRows = True
'    objWord.selection.Tables(1).ApplyStyleLastRow = True
'    objWord.selection.Tables(1).ApplyStyleFirstColumn = True
'    objWord.selection.Tables(1).ApplyStyleLastColumn = True
'
'    Exit Sub
'
'End Sub
'
'Private Sub DOC_Insere_Cabec_Tabela(ByVal objWord As Object, ParamArray avParams())
'
'Dim iIndice As Integer
'Dim bBold As Boolean
'
'    bBold = objWord.selection.Font.Bold
'    For iIndice = 0 To UBound(avParams)
'
'        objWord.selection.Font.Bold = True
'        objWord.selection.TypeText avParams(iIndice)
'
'        'If iIndice <> UBound(avParams) Then
'            objWord.selection.MoveRight wdCharacter, 1
'        'End If
'
'    Next
'    objWord.selection.Font.Bold = bBold
'
'    Exit Sub
'
'End Sub
'
'Private Sub DOC_Insere_Valores_Tabela(ByVal objWord As Object, ParamArray avParams())
'
'Dim iIndice As Integer
'
'    objWord.selection.MoveRight wdCell
'    'objWord.selection.GoToNext wdGoToLine
'
'    For iIndice = 0 To UBound(avParams)
'
'        objWord.selection.TypeText avParams(iIndice)
'
'        'If iIndice <> UBound(avParams) Then
'            objWord.selection.MoveRight wdCharacter, 1
'        'End If
'
'    Next
'
'    Exit Sub
'
'End Sub
'
'Private Sub DOC_Insere_Valores_Tabela2(ByVal objWord As Object, ParamArray avParams())
'
'Dim iIndice As Integer
'
'    For iIndice = 0 To UBound(avParams)
'
'        objWord.selection.TypeText avParams(iIndice)
'
'        'If iIndice <> UBound(avParams) Then
'            objWord.selection.MoveRight wdCharacter, 1
'        'End If
'
'    Next
'
'    Exit Sub
'
'End Sub

Private Sub BotaoProcurar_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoProcurar_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Word Files" & _
    "(*.doc)|*.doc"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    Modelo.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoProcurar_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub Dir1_Change()

     NomeDiretorio.Text = Dir1.Path

End Sub

Private Sub Dir1_Click()

On Error GoTo Erro_Dir1_Click

    Exit Sub
    
Erro_Dir1_Click:

    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189911)
    
    Exit Sub

End Sub

Private Sub Drive1_Change()

On Error GoTo Erro_Drive1_Change

    Dir1.Path = Drive1.Drive
       
    Exit Sub

Erro_Drive1_Change:

    Select Case gErr
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189910)

    End Select

    Drive1.ListIndex = iListIndexDefault
    
    Exit Sub
    
End Sub

Private Sub Drive1_GotFocus()
    
    iListIndexDefault = Drive1.ListIndex

End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 189908

    Drive1.Drive = Mid(NomeDiretorio.Text, 1, 2)

    Dir1.Path = NomeDiretorio.Text

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 189908, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189909)

    End Select

    Exit Sub

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

Private Sub TvwEtapas_DblClick()

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_TvwEtapas_DblClick

    objEtapa.lNumIntDocPRJ = gobjEtapa.lNumIntDocPRJ
    objEtapa.lNumIntDoc = gobjEtapa.lNumIntDoc
    objEtapa.sCodigo = gobjEtapa.sCodigo

    Call Chama_Tela("EtapaPRJ", objEtapa)

    Exit Sub

Erro_TvwEtapas_DblClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189375)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelProposta_Click()

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas
Dim colSelecao As New Collection
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProposta_Click

    If Len(Trim(PRJ.ClipText)) = 0 Then gError 187546

    lErro = Projeto_Formata(PRJ.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189082

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187547
    
    'Se não encontrou => Erro
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187548
    
    colSelecao.Add objProjeto.lNumIntDoc

    'Verifica se o Proposta foi preenchido
    If Len(Trim(Proposta.Text)) <> 0 Then

        objProposta.sCodigo = Proposta.Text

    End If

    Call Chama_Tela("PRJPropostasLista", colSelecao, objProposta, objEventoProposta, "NumIntDocPRJ = ?", "Código")

    Exit Sub

Erro_LabelProposta_Click:

    Select Case gErr
    
        Case 187546
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            PRJ.SetFocus
            
        Case 187547, 189082

        Case 187548
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187549)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProposta_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas

On Error GoTo Erro_objEventoProposta_evSelecao

    Set objProposta = obj1
    
    Proposta.Text = objProposta.sCodigo
    Call Proposta_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoProposta_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187550)

    End Select

    Exit Sub

End Sub

Private Sub Proposta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Proposta_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Proposta_Validate

    If Len(Trim(Proposta.Text)) > 0 Then
    
        If Len(Trim(PRJ.ClipText)) = 0 Then gError 187551
    
        lErro = Projeto_Formata(PRJ.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189083
    
        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        'Le
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187552
        
        'Se não encontrou => Erro
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187553
        
        objProposta.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objProposta.sCodigo = Proposta.Text
        
        'Lê a proposta que está sendo Passado
        lErro = CF("PRJPropostas_Le", objProposta)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187554
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187555
        
        Set gobjProposta = objProposta
    
    Else
    
        Set gobjProposta = Nothing
            
    End If
    
    Exit Sub

Erro_Proposta_Validate:

    Cancel = True

    Select Case gErr

        Case 187551
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus
            
        Case 187552, 187554, 189083

        Case 187553
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case 187555
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJPROPOSTAS_NAO_CADASTRADO", gErr, Proposta.Text, objProjeto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187556)

    End Select

    Exit Sub
    
End Sub

Private Sub PRJLabel_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_PRJLabel_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(PRJ.Text)) <> 0 Then

        lErro = Projeto_Formata(PRJ.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189103

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela_Modal("ProjetosLista", colSelecao, objProjeto, objEventoPRJ, , "Código")

    Exit Sub

Erro_PRJLabel_Click:

    Select Case gErr

        Case 189103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoPRJ_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoPRJ_evSelecao

    Set objProjeto = obj1
    
    lErro = Retorno_Projeto_Tela(PRJ, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189107

    Call PRJ_Validate(bSGECancelDummy)

    Exit Sub

Erro_objEventoPRJ_evSelecao:

    Select Case gErr
    
        Case 189107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181629)

    End Select

    Exit Sub

End Sub

Private Sub PRJ_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_PRJ_Validate

    If Len(Trim(PRJ.ClipText)) > 0 Then

        lErro = Projeto_Formata(PRJ.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189104

        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        'Le o almoxarifado pelo código ou pelo nome reduzido e joga o nome reduzido em Almoxarifado.Text
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181669
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181670
        
        If sPRJAnt <> PRJ.Text Then
            sPRJAnt = PRJ.Text
            Proposta.Text = ""
            Set gobjProposta = Nothing
        End If
        
    End If
    
    Exit Sub

Erro_PRJ_Validate:

    Cancel = True

    Select Case gErr
    
        Case 181669, 189104
        
        Case 181670
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181671)

    End Select

    Exit Sub
    
End Sub

Private Sub PRJ_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoVerProposta_Click()

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas

On Error GoTo Erro_BotaoVerProposta_Click

    If gobjProposta Is Nothing Then gError 189378

    objProposta.lNumIntDoc = gobjProposta.lNumIntDoc
    objProposta.lNumIntDocPRJ = gobjProposta.lNumIntDocPRJ
    objProposta.sCodigo = gobjProposta.sCodigo
    
    'Chama  a tela de Pedido de Venda passando o pedido de venda da tela
    Call Chama_Tela("PropostaPRJ", objProposta)
    
    Exit Sub

Erro_BotaoVerProposta_Click:

    Select Case gErr
    
        Case 189378
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJPROPOSTA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189379)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProposta_Click()

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas

On Error GoTo Erro_BotaoProposta_Click

    If gobjProposta Is Nothing Then gError 189380

    objProposta.lNumIntDoc = gobjProposta.lNumIntDoc
    objProposta.lNumIntDocPRJ = gobjProposta.lNumIntDocPRJ
    objProposta.sCodigo = gobjProposta.sCodigo

    'Lê a Proposta que está sendo Passado
    lErro = CF("PRJPropostas_Le", objProposta, True, True)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187782
    
    If objProposta.lNumIntDocContrato <> 0 Then gError 189395

    'Traz os dados da proposta para a tela
    lErro = Traz_Dados_Proposta_Tela(objProposta)
    If lErro <> SUCESSO Then gError 189396

    Call ValorTotal_Calcula
    
    Exit Sub

Erro_BotaoProposta_Click:

    Select Case gErr
    
        Case 189380
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJPROPOSTA_NAO_PREENCHIDO", gErr)

        Case 187782, 189396
        
        Case 189395
            Call Rotina_Erro(vbOKOnly, "ERRO_PROPOSTA_VINCULADA_CONTRATO", gErr, objProposta.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189381)

    End Select

    Exit Sub

End Sub

Public Function Traz_Dados_Proposta_Tela(objProposta As ClassPRJPropostas) As Long

Dim lErro As Long
Dim objContrato As New ClassPRJContratos
Dim objCliente As New ClassCliente

On Error GoTo Erro_Traz_Dados_Proposta_Tela
    
    'Transfere os dados
    Call Transfere_Dados_Proposta_Contrato(objProposta, objContrato)
    
    'Carrega os dados
    lErro = Traz_Contrato_Tela(objContrato, True)
    If lErro <> SUCESSO Then gError 189396
               
    Traz_Dados_Proposta_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Proposta_Tela:

    Traz_Dados_Proposta_Tela = gErr

    Select Case gErr
    
        Case 189396

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189397)

    End Select

    Exit Function

End Function

Public Sub Transfere_Dados_Proposta_Contrato(objProposta As ClassPRJPropostas, objContrato As ClassPRJContratos)
'Transfere os dados do objProposta para objPedidoDeVenda

Dim objItemProposta As ClassPRJPropostaItem
Dim objItemContrato As ClassPRJContratoItem
Dim objPropostaEtapa As ClassPRJPropostaEtapa
Dim objContratoEtapa As ClassPRJContratoEtapa

On Error GoTo Erro_Transfere_Dados_Proposta_Contrato

    With objContrato
        
        .dCustoCalculado = objProposta.dCustoCalculado
        .dCustoInformado = objProposta.dCustoInformado
        .dtData = objProposta.dtData
        .dValorDesconto = objProposta.dValorDesconto
        .dValorFrete = objProposta.dValorFrete
        .dValorOutrasDespesas = objProposta.dValorOutrasDespesas
        .dValorProdutos = objProposta.dValorProdutos
        .dValorSeguro = objProposta.dValorSeguro
        .dValorTotal = objProposta.dValorTotal
        .iExibirCustoCalc = objProposta.iExibirCustoCalc
        .iExibirCustoInfo = objProposta.iExibirCustoInfo
        .iExibirPreco = objProposta.iExibirPreco
        .iExibirProdutos = objProposta.iExibirProdutos
        .iFilialCliente = objProposta.iFilialCliente
        .iFilialEmpresa = objProposta.iFilialEmpresa
        .iNaturezaOpManual = objProposta.iNaturezaOpManual
        .lCliente = objProposta.lCliente
        .lNumIntDocPRJ = objProposta.lNumIntDocPRJ
        .lNumIntDocProposta = objProposta.lNumIntDoc
        .sNaturezaOp = objProposta.sNaturezaOp
        .sObservacao = objProposta.sObservacao
        .dValorDescontoItens = objProposta.dValorDescontoItens
        .dValorItens = objProposta.dValorItens
        
    End With
    
    For Each objPropostaEtapa In objProposta.colEtapas
    
        Set objContratoEtapa = New ClassPRJContratoEtapa
    
        With objContratoEtapa
        
            .dCustoInformado = objPropostaEtapa.dCustoInformado
            .dPreco = objPropostaEtapa.dPreco
            .iImprimir = objPropostaEtapa.iImprimir
            .iSelecionado = objPropostaEtapa.iSelecionado
            .lNumIntDocEtapa = objPropostaEtapa.lNumIntDocEtapa
            .lNumIntDocEtapaItemProd = objPropostaEtapa.lNumIntDocEtapaItemProd
            .sDescricao = objPropostaEtapa.sDescricao
            .sObservacao = objPropostaEtapa.sObservacao
        
        End With
    
        objContrato.colEtapas.Add objContratoEtapa
    
    Next

    'Para cada item
    For Each objItemProposta In objProposta.colItens
        
        Set objItemContrato = New ClassPRJContratoItem
        
        Call ItemContrato_ObterTrib_ItemProposta(objItemContrato, objItemProposta)
        
        With objItemContrato
        
            .dPrecoTotal = objItemProposta.dPrecoTotal
            .dPrecoUnitario = objItemProposta.dPrecoUnitario
            .dQuantidade = objItemProposta.dQuantidade
            .dtDataEntrega = objItemProposta.dtDataEntrega
            .dValorDescGlobal = objItemProposta.dValorDescGlobal
            .dValorDesconto = objItemProposta.dValorDesconto
            .iFilialEmpresa = objItemProposta.iFilialEmpresa
            .iItem = objItemProposta.iItem
            .lNumIntDocEtapa = objItemProposta.lNumIntDocEtapa
            .sCodEtapa = objItemProposta.sCodEtapa
            .sDescEtapa = objItemProposta.sDescEtapa
            .sDescProd = objItemProposta.sDescProd
            .sObservacao = objItemProposta.sObservacao
            .sProduto = objItemProposta.sProduto
            .sUM = objItemProposta.sUM
                        
        End With
        
        objContrato.colItens.Add objItemContrato

    Next
    
    Call Contrato_ObterTrib_Proposta(objContrato, objProposta)
    
    Exit Sub

Erro_Transfere_Dados_Proposta_Contrato:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189398)

    End Select

End Sub

'Private Sub ComplCRT_ObterTrib_ComplProp(objTribCompCTR As ClassTributacaoItemPV, objTribComplProp As ClassTributacaoItemPV)
'
'    With objTribCompCTR
'
'        .dICMSAliquota = objTribComplProp.dICMSAliquota
'        .dICMSBase = objTribComplProp.dICMSBase
'        .dICMSPercRedBase = objTribComplProp.dICMSPercRedBase
'        .dICMSSubstAliquota = objTribComplProp.dICMSSubstAliquota
'        .dICMSSubstBase = objTribComplProp.dICMSSubstBase
'        .dICMSSubstValor = objTribComplProp.dICMSSubstValor
'        .dICMSValor = objTribComplProp.dICMSValor
'        .dIPIAliquota = objTribComplProp.dIPIAliquota
'        .dIPIBaseCalculo = objTribComplProp.dIPIBaseCalculo
'        .dIPIPercRedBase = objTribComplProp.dIPIPercRedBase
'        .dIPIValor = objTribComplProp.dIPIValor
'        .iICMSTipo = objTribComplProp.iICMSTipo
'        .iIPITipo = objTribComplProp.iIPITipo
'        .iTipoTributacao = objTribComplProp.iTipoTributacao
'        .lNumIntDoc = objTribComplProp.lNumIntDoc
'        .sNaturezaOp = objTribComplProp.sNaturezaOp
'
'        .iNaturezaOpManual = objTribComplProp.iNaturezaOpManual
'        .iTipoTributacaoManual = objTribComplProp.iTipoTributacaoManual
'        .iIPITipoManual = objTribComplProp.iIPITipoManual
'        .iIPIBaseManual = objTribComplProp.iIPIBaseManual
'        .iIPIPercRedBaseManual = objTribComplProp.iIPIPercRedBaseManual
'        .iIPIAliquotaManual = objTribComplProp.iIPIAliquotaManual
'        .iIPIValorManual = objTribComplProp.iIPIValorManual
'        .iICMSTipoManual = objTribComplProp.iICMSTipoManual
'        .iICMSBaseManual = objTribComplProp.iICMSBaseManual
'        .iICMSPercRedBaseManual = objTribComplProp.iICMSPercRedBaseManual
'        .iICMSAliquotaManual = objTribComplProp.iICMSAliquotaManual
'        .iICMSValorManual = objTribComplProp.iICMSValorManual
'        .iICMSSubstBaseManual = objTribComplProp.iICMSSubstBaseManual
'        .iICMSSubstAliquotaManual = objTribComplProp.iICMSSubstAliquotaManual
'        .iICMSSubstValorManual = objTribComplProp.iICMSSubstValorManual
'
'    End With
'
'End Sub

Sub Contrato_ObterTrib_Proposta(objContrato As ClassPRJContratos, objProposta As ClassPRJPropostas)

    Call objContrato.objTributacaoPRJCTR.Copia(objProposta.objTributacaoPRJProp)

'    Set objContrato.objTributacaoPRJCTR = New ClassTributacaoPRJCTR
'
'    With objContrato.objTributacaoPRJCTR
'
'        .dICMSBase = objProposta.objTributacaoPRJProp.dICMSBase
'        .dICMSSubstBase = objProposta.objTributacaoPRJProp.dICMSSubstBase
'        .dICMSSubstValor = objProposta.objTributacaoPRJProp.dICMSSubstValor
'        .dICMSValor = objProposta.objTributacaoPRJProp.dICMSValor
'        .dIPIBase = objProposta.objTributacaoPRJProp.dIPIBase
'        .dIPIValor = objProposta.objTributacaoPRJProp.dIPIValor
'        .dIRRFAliquota = objProposta.objTributacaoPRJProp.dIRRFAliquota
'        .dIRRFBase = objProposta.objTributacaoPRJProp.dIRRFBase
'        .dIRRFValor = objProposta.objTributacaoPRJProp.dIRRFValor
'        .dISSBase = objProposta.objTributacaoPRJProp.dISSBase
'        .dISSValor = objProposta.objTributacaoPRJProp.dISSValor
'        .iISSIncluso = objProposta.objTributacaoPRJProp.iISSIncluso
'        .iTipoTributacao = objProposta.objTributacaoPRJProp.iTipoTributacao
'        .dPISRetido = objProposta.objTributacaoPRJProp.dPISRetido
'        .dISSRetido = objProposta.objTributacaoPRJProp.dISSRetido
'        .dCOFINSRetido = objProposta.objTributacaoPRJProp.dCOFINSRetido
'        .dCSLLRetido = objProposta.objTributacaoPRJProp.dCSLLRetido
'
'        .iTipoTributacaoManual = objProposta.objTributacaoPRJProp.iTipoTributacaoManual
'        .iICMSBaseManual = objProposta.objTributacaoPRJProp.iICMSBaseManual
'        .iICMSValorManual = objProposta.objTributacaoPRJProp.iICMSValorManual
'        .iICMSSubstBaseManual = objProposta.objTributacaoPRJProp.iICMSSubstBaseManual
'        .iICMSSubstValorManual = objProposta.objTributacaoPRJProp.iICMSSubstValorManual
'        .iIPIBaseManual = objProposta.objTributacaoPRJProp.iIPIBaseManual
'        .iIPIValorManual = objProposta.objTributacaoPRJProp.iIPIValorManual
'        .iIRRFAliquotaManual = objProposta.objTributacaoPRJProp.iIRRFAliquotaManual
'        .iIRRFValorManual = objProposta.objTributacaoPRJProp.iIRRFValorManual
'        .iISSAliquotaManual = objProposta.objTributacaoPRJProp.iISSAliquotaManual
'        .iISSInclusoManual = objProposta.objTributacaoPRJProp.iISSInclusoManual
'        .iISSValorManual = objProposta.objTributacaoPRJProp.iISSValorManual
'        .iPISRetidoManual = objProposta.objTributacaoPRJProp.iPISRetidoManual
'        .iISSRetidoManual = objProposta.objTributacaoPRJProp.iISSRetidoManual
'        .iCOFINSRetidoManual = objProposta.objTributacaoPRJProp.iCOFINSRetidoManual
'        .iCSLLRetidoManual = objProposta.objTributacaoPRJProp.iCSLLRetidoManual
'
'        Set .objTributacaoDesconto = New ClassTributacaoItemPV
'        Set .objTributacaoFrete = New ClassTributacaoItemPV
'        Set .objTributacaoOutras = New ClassTributacaoItemPV
'        Set .objTributacaoSeguro = New ClassTributacaoItemPV
'
'        Call ComplCRT_ObterTrib_ComplProp(.objTributacaoDesconto, objProposta.objTributacaoPRJProp.objTributacaoDesconto)
'        Call ComplCRT_ObterTrib_ComplProp(.objTributacaoFrete, objProposta.objTributacaoPRJProp.objTributacaoFrete)
'        Call ComplCRT_ObterTrib_ComplProp(.objTributacaoOutras, objProposta.objTributacaoPRJProp.objTributacaoOutras)
'        Call ComplCRT_ObterTrib_ComplProp(.objTributacaoSeguro, objProposta.objTributacaoPRJProp.objTributacaoSeguro)
'
'    End With

End Sub

Sub ItemContrato_ObterTrib_ItemProposta(objItemContrato As ClassPRJContratoItem, objItemProposta As ClassPRJPropostaItem)

    Call objItemContrato.objTributacaoPRJCTRItem.Copia(objItemProposta.objTributacaoPRJPropItem)

'    Set objItemContrato = New ClassPRJContratoItem
'
'    With objItemProposta
'
'        objItemContrato.objTributacaoPRJCTRItem.dICMSAliquota = .objTributacaoPRJPropItem.dICMSAliquota
'        objItemContrato.objTributacaoPRJCTRItem.dICMSBase = .objTributacaoPRJPropItem.dICMSBase
'        objItemContrato.objTributacaoPRJCTRItem.dICMSPercRedBase = .objTributacaoPRJPropItem.dICMSPercRedBase
'        objItemContrato.objTributacaoPRJCTRItem.dICMSSubstAliquota = .objTributacaoPRJPropItem.dICMSSubstAliquota
'        objItemContrato.objTributacaoPRJCTRItem.dICMSSubstBase = .objTributacaoPRJPropItem.dICMSSubstBase
'        objItemContrato.objTributacaoPRJCTRItem.dICMSSubstValor = .objTributacaoPRJPropItem.dICMSSubstValor
'        objItemContrato.objTributacaoPRJCTRItem.dICMSValor = .objTributacaoPRJPropItem.dICMSValor
'        objItemContrato.objTributacaoPRJCTRItem.dIPIAliquota = .objTributacaoPRJPropItem.dIPIAliquota
'        objItemContrato.objTributacaoPRJCTRItem.dIPIBaseCalculo = .objTributacaoPRJPropItem.dIPIBaseCalculo
'        objItemContrato.objTributacaoPRJCTRItem.dIPIPercRedBase = .objTributacaoPRJPropItem.dIPIPercRedBase
'        objItemContrato.objTributacaoPRJCTRItem.dIPIValor = .objTributacaoPRJPropItem.dIPIValor
'        objItemContrato.objTributacaoPRJCTRItem.iICMSTipo = .objTributacaoPRJPropItem.iICMSTipo
'        objItemContrato.objTributacaoPRJCTRItem.iIPITipo = .objTributacaoPRJPropItem.iIPITipo
'        objItemContrato.objTributacaoPRJCTRItem.iTipoTributacao = .objTributacaoPRJPropItem.iTipoTributacao
'        objItemContrato.objTributacaoPRJCTRItem.sNaturezaOp = .objTributacaoPRJPropItem.sNaturezaOp
'
'        objItemContrato.objTributacaoPRJCTRItem.iNaturezaOpManual = .objTributacaoPRJPropItem.iNaturezaOpManual
'        objItemContrato.objTributacaoPRJCTRItem.iTipoTributacaoManual = .objTributacaoPRJPropItem.iTipoTributacaoManual
'        objItemContrato.objTributacaoPRJCTRItem.iIPITipoManual = .objTributacaoPRJPropItem.iIPITipoManual
'        objItemContrato.objTributacaoPRJCTRItem.iIPIBaseManual = .objTributacaoPRJPropItem.iIPIBaseManual
'        objItemContrato.objTributacaoPRJCTRItem.iIPIPercRedBaseManual = .objTributacaoPRJPropItem.iIPIPercRedBaseManual
'        objItemContrato.objTributacaoPRJCTRItem.iIPIAliquotaManual = .objTributacaoPRJPropItem.iIPIAliquotaManual
'        objItemContrato.objTributacaoPRJCTRItem.iIPIValorManual = .objTributacaoPRJPropItem.iIPIValorManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSTipoManual = .objTributacaoPRJPropItem.iICMSTipoManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSBaseManual = .objTributacaoPRJPropItem.iICMSBaseManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSPercRedBaseManual = .objTributacaoPRJPropItem.iICMSPercRedBaseManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSAliquotaManual = .objTributacaoPRJPropItem.iICMSAliquotaManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSValorManual = .objTributacaoPRJPropItem.iICMSValorManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSSubstBaseManual = .objTributacaoPRJPropItem.iICMSSubstBaseManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSSubstAliquotaManual = .objTributacaoPRJPropItem.iICMSSubstAliquotaManual
'        objItemContrato.objTributacaoPRJCTRItem.iICMSSubstValorManual = .objTributacaoPRJPropItem.iICMSSubstValorManual
'
'    End With
    
End Sub

Public Property Get DataEmissao() As Object
     Set DataEmissao = DataCriacao
End Property

Public Sub ValorDescontoItens_Change()
    'Seta iComissoesAlterada
    'iComissoesAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercDescontoItens_Change()
    'Seta iComissoesAlterada
    'iComissoesAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorDescontoItens_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim vbMsg As VbMsgBoxResult, iIndice As Integer
Dim dDescontoItens As Double, dFator As Double

On Error GoTo Erro_ValorDescontoItens_Validate

    dValorDesconto = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDescontoItens.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDescontoItens.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorDesconto = CDbl(ValorDescontoItens.Text)
        
        'Coloca o Valor formatado na tela
        ValorDescontoItens.Text = Format(dValorDesconto, "Standard")
        
    End If
    
    'Se houve alguma alteração
    If Abs(dValorDescontoItensAnt - dValorDesconto) > DELTA_VALORMONETARIO Then
        
        'Se o desconto foi alterado nos itens pegunta se quer que o sistema recalcule
        If iDescontoAlterado = REGISTRO_ALTERADO Then
        
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_DESCONTO_ITENS_ALTERADO")
            If vbMsg = vbNo Then
                dValorDesconto = dValorDescontoItensAnt
                ValorDescontoItens.Text = Format(dValorDesconto, "Standard")
                gError ERRO_SEM_MENSAGEM
            End If
            iDescontoAlterado = 0
               
        End If
        
        lErro = ValorDescontoItens_Aplica
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dValorDescontoItensAnt = dValorDesconto
    
    End If
    
    If StrParaDbl(ValorProdutos2.Caption) > 0 Then
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDesconto / StrParaDbl(ValorProdutos2.Caption), 4)
    Else
        dPercDescontoItensAnt = 0
    End If
    PercDescontoItens.Text = Format(dPercDescontoItensAnt * 100, "FIXED")
    
    Exit Sub

Erro_ValorDescontoItens_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Public Sub PercDescontoItens_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercDesconto As Double
Dim vbMsg As VbMsgBoxResult, iIndice As Integer
Dim dDescontoItens As Double, dFator As Double

On Error GoTo Erro_PercDescontoItens_Validate

    dPercDesconto = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(PercDescontoItens.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Porcentagem_Critica(PercDescontoItens.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dPercDesconto = StrParaDbl(PercDescontoItens.Text) / 100

    End If
    
    'Se houve alguma alteração
    If Abs(dPercDescontoItensAnt - dPercDesconto) > DELTA_VALORMONETARIO2 Then
        
        'Se o desconto foi alterado nos itens pegunta se quer que o sistema recalcule
        If iDescontoAlterado = REGISTRO_ALTERADO Then
        
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_DESCONTO_ITENS_ALTERADO")
            If vbMsg = vbNo Then
                dPercDesconto = dPercDescontoItensAnt
                PercDescontoItens.Text = Format(dPercDesconto * 100, "FIXED")
                gError ERRO_SEM_MENSAGEM
            End If
            iDescontoAlterado = 0
                
        End If
        
        ValorDescontoItens.Text = Format(Arredonda_Moeda(dPercDesconto * StrParaDbl(ValorProdutos2.Caption)), "Standard")
        Call ValorDescontoItens_Validate(bSGECancelDummy)
    
    End If
        
    Exit Sub

Erro_PercDescontoItens_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Function ValorDescontoItens_Aplica() As Long

Dim lErro As Long, dTotal As Double, dValorItem As Double, dDescontoItem As Double
Dim dDesconto As Double, dFator As Double, dDescontoAplicado As Double, dDiferenca As Double
Dim dPercDesc As Double, iIndice As Integer, dValorTotal As Double

On Error GoTo Erro_ValorDescontoItens_Aplica

    If objGridItens.iLinhasExistentes > 0 Then
    
        dTotal = StrParaDbl(ValorProdutos2.Caption)
        dDesconto = StrParaDbl(ValorDescontoItens.Text)
        dFator = dDesconto / dTotal
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))) > 0 Then
                dValorItem = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
                dDescontoItem = Arredonda_Moeda(dValorItem * dFator)
                dDescontoAplicado = dDescontoAplicado + dDescontoItem
                GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(dDescontoItem, "STANDARD")
            End If
        Next
        If Abs(dDesconto - dDescontoAplicado) > DELTA_VALORMONETARIO Then
            GridItens.TextMatrix(1, iGrid_Desconto_Col) = Format(StrParaDbl(GridItens.TextMatrix(1, iGrid_Desconto_Col)) + dDescontoAplicado - dDesconto, "STANDARD")
        End If
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
            dPercDesc = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
            Call PrecoTotal_Calcula(iIndice)
            
            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        Next
    
    End If
        
    'Coloca valor total dos produtos na tela
    ValorProdutos.Caption = Format(dValorTotal, "Standard")

    'Calcula o valor total da nota
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    ValorDescontoItens_Aplica = SUCESSO

    Exit Function

Erro_ValorDescontoItens_Aplica:

    ValorDescontoItens_Aplica = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function

Function ValorDescontoItens_Calcula() As Long

Dim lErro As Long, iIndice As Integer
Dim dDesconto As Double, dPercDesc As Double

On Error GoTo Erro_ValorDescontoItens_Calcula

    dDesconto = 0
    dPercDesc = 0
    If Not (objGridItens Is Nothing) Then
        If objGridItens.iLinhasExistentes > 0 Then
            For iIndice = 1 To objGridItens.iLinhasExistentes
                dDesconto = dDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
            Next
            If StrParaDbl(ValorProdutos2.Caption) > 0 Then dPercDesc = Arredonda_Moeda(dDesconto / StrParaDbl(ValorProdutos2.Caption), 4)
            ValorDescontoItens.Text = Format(dDesconto, "Standard")
            PercDescontoItens.Text = Format(dPercDesc * 100, "FIXED")
        Else
            ValorDescontoItens.Text = Format(0, "Standard")
            PercDescontoItens.Text = Format(0, "FIXED")
        End If
        
        dValorDescontoItensAnt = dDesconto
        dPercDescontoItensAnt = dPercDesc
    End If

    ValorDescontoItens_Calcula = SUCESSO

    Exit Function

Erro_ValorDescontoItens_Calcula:

    ValorDescontoItens_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function


