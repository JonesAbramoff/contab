VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrcamentoVenda 
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4830
      Index           =   4
      Left            =   30
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   9375
      Begin TelasFATCro.TabTributacaoFat TabTrib 
         Height          =   4605
         Left            =   255
         TabIndex        =   142
         Top             =   90
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   8123
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   2
      Left            =   120
      TabIndex        =   42
      Top             =   1095
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   2685
         Index           =   3
         Left            =   225
         TabIndex        =   60
         Top             =   0
         Width           =   8865
         Begin VB.ComboBox CondPagtoItem 
            Height          =   315
            Left            =   2625
            TabIndex        =   110
            Top             =   1410
            Width           =   1395
         End
         Begin VB.CheckBox Escolhido 
            Caption         =   "Escolhido"
            Height          =   255
            Left            =   4530
            TabIndex        =   109
            Top             =   1575
            Value           =   1  'Checked
            Width           =   1005
         End
         Begin VB.TextBox Concorrente 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3990
            MaxLength       =   50
            TabIndex        =   108
            Top             =   1950
            Width           =   2280
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoVendaCro.ctx":0000
            Left            =   1575
            List            =   "OrcamentoVendaCro.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   240
            Width           =   720
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4005
            MaxLength       =   250
            TabIndex        =   67
            Top             =   765
            Width           =   2490
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3330
            MaxLength       =   50
            TabIndex        =   63
            Top             =   570
            Width           =   1305
         End
         Begin VB.ComboBox MotivoPerdaItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   15
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   465
            Width           =   1905
         End
         Begin VB.ComboBox StatusItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   0
            Width           =   1920
         End
         Begin MSMask.MaskEdBox VersaoKit 
            Height          =   225
            Left            =   5220
            TabIndex        =   64
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
            TabIndex        =   65
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
            Left            =   330
            TabIndex        =   66
            Top             =   360
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
            Left            =   2640
            TabIndex        =   69
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
            TabIndex        =   70
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
            TabIndex        =   71
            Top             =   675
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
            TabIndex        =   72
            Top             =   360
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
            Left            =   2580
            TabIndex        =   73
            Top             =   315
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
            Left            =   5670
            TabIndex        =   74
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
            Left            =   180
            TabIndex        =   75
            Top             =   225
            Width           =   8565
            _ExtentX        =   15108
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
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   1290
         Index           =   4
         Left            =   225
         TabIndex        =   47
         Top             =   2745
         Width           =   8865
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1695
            TabIndex        =   48
            Top             =   975
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   450
            TabIndex        =   49
            Top             =   450
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   4320
            TabIndex        =   50
            Top             =   975
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   2985
            TabIndex        =   51
            Top             =   975
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   450
            TabIndex        =   141
            Top             =   975
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ISS               Frete             Seguro              Despesas               IPI                Total"
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
            Index           =   63
            Left            =   825
            TabIndex        =   140
            Top             =   780
            Width           =   7230
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7320
            TabIndex        =   59
            Top             =   975
            Width           =   1125
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5985
            TabIndex        =   58
            Top             =   975
            Width           =   1125
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7305
            TabIndex        =   57
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1710
            TabIndex        =   56
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3015
            TabIndex        =   55
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4320
            TabIndex        =   54
            Top             =   450
            Width           =   1500
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   53
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Desconto        Base ICMS          ICMS         Base ICMS Subst    ICMS Subst       Produtos"
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
            Left            =   495
            TabIndex        =   52
            Top             =   225
            Width           =   7695
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
         Left            =   7350
         TabIndex        =   46
         Top             =   4140
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
         Left            =   225
         TabIndex        =   45
         Top             =   4140
         Width           =   1725
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
         Left            =   4975
         TabIndex        =   44
         Top             =   4140
         Width           =   1725
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
         Left            =   2600
         TabIndex        =   43
         Top             =   4140
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4875
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   915
      Width           =   9300
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1485
         Index           =   0
         Left            =   90
         TabIndex        =   35
         Top             =   105
         Width           =   8865
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2760
            Picture         =   "OrcamentoVendaCro.ctx":0004
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   225
            Width           =   300
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   5355
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1035
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
            Left            =   3750
            TabIndex        =   6
            Top             =   1065
            Width           =   495
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2850
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   615
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1800
            TabIndex        =   3
            Top             =   615
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
            Left            =   1800
            TabIndex        =   0
            Top             =   210
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
            Left            =   5355
            TabIndex        =   2
            Top             =   195
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1800
            TabIndex        =   5
            Top             =   1065
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   3630
            TabIndex        =   139
            Top             =   645
            Width           =   1650
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
            Left            =   5355
            TabIndex        =   138
            Top             =   585
            Width           =   525
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
            Left            =   1065
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   255
            Width           =   720
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
            Left            =   990
            TabIndex        =   39
            Top             =   660
            Width           =   765
         End
         Begin VB.Label NumeroBaseLabel 
            AutoSize        =   -1  'True
            Caption         =   "Número Base:"
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
            Left            =   4035
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   240
            Width           =   1200
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
            Left            =   4755
            TabIndex        =   37
            Top             =   1095
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
            Left            =   1065
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   36
            Top             =   1110
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Preços"
         Height          =   690
         Index           =   2
         Left            =   90
         TabIndex        =   31
         Top             =   2295
         Width           =   8865
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   1305
            TabIndex        =   10
            Top             =   240
            Width           =   1875
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   4530
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
         Begin MSMask.MaskEdBox PercAcrescFin 
            Height          =   315
            Left            =   7995
            TabIndex        =   12
            Top             =   240
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
            TabIndex        =   34
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Acrésc Financ:"
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
            Left            =   6480
            TabIndex        =   33
            Top             =   300
            Width           =   1485
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
            Height          =   195
            Left            =   3390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   615
         Index           =   6
         Left            =   90
         TabIndex        =   28
         Top             =   1650
         Width           =   8865
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5340
            TabIndex        =   9
            Top             =   195
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1785
            TabIndex        =   8
            Top             =   195
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
            Left            =   1125
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
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
            Left            =   4740
            TabIndex        =   29
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Outros"
         Height          =   1575
         Left            =   105
         TabIndex        =   20
         Top             =   3000
         Width           =   8865
         Begin VB.ComboBox StatusComercial 
            Height          =   315
            ItemData        =   "OrcamentoVendaCro.ctx":00EE
            Left            =   7320
            List            =   "OrcamentoVendaCro.ctx":00FE
            TabIndex        =   17
            Top             =   660
            Width           =   1485
         End
         Begin VB.ComboBox Status 
            Height          =   315
            Left            =   4065
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   660
            Width           =   1440
         End
         Begin VB.ComboBox MotivoPerda 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1125
            Width           =   7485
         End
         Begin MSMask.MaskEdBox PrazoValidade 
            Height          =   300
            Left            =   5055
            TabIndex        =   14
            Top             =   255
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
            Left            =   1320
            TabIndex        =   13
            Top             =   255
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Vendedor2 
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Top             =   660
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
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
            Left            =   5670
            TabIndex        =   111
            Top             =   705
            Width           =   1650
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
            Left            =   3390
            TabIndex        =   27
            Top             =   300
            Width           =   1620
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
            Left            =   435
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   300
            Width           =   885
         End
         Begin VB.Label UsuarioLabel 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
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
            Left            =   6045
            TabIndex        =   25
            Top             =   315
            Width           =   720
         End
         Begin VB.Label Usuario 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6825
            TabIndex        =   24
            Top             =   255
            Width           =   1965
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
            Left            =   45
            TabIndex        =   23
            Top             =   1170
            Width           =   1200
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
            Left            =   3450
            TabIndex        =   22
            Top             =   720
            Width           =   615
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
            Left            =   90
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            Top             =   705
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4875
      Index           =   5
      Left            =   30
      TabIndex        =   112
      Top             =   915
      Visible         =   0   'False
      Width           =   9315
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
         TabIndex        =   137
         Top             =   4485
         Width           =   2085
      End
      Begin VB.Frame Frame6 
         Caption         =   "Preços Calculados"
         Height          =   2010
         Left            =   105
         TabIndex        =   126
         Top             =   -15
         Width           =   9165
         Begin VB.ComboBox PCUnidMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OrcamentoVendaCro.ctx":0122
            Left            =   3315
            List            =   "OrcamentoVendaCro.ctx":0124
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   360
            Width           =   720
         End
         Begin VB.TextBox PCDescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1635
            MaxLength       =   50
            TabIndex        =   130
            Top             =   435
            Width           =   1485
         End
         Begin VB.ComboBox PCSituacao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoVendaCro.ctx":0126
            Left            =   6345
            List            =   "OrcamentoVendaCro.ctx":0136
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   1125
            Width           =   1155
         End
         Begin VB.OptionButton PCSelecionado 
            Caption         =   "Option1"
            Height          =   225
            Left            =   105
            TabIndex        =   127
            Top             =   255
            Width           =   495
         End
         Begin MSMask.MaskEdBox PCPrecoUnitCalc 
            Height          =   225
            Left            =   5550
            TabIndex        =   128
            Top             =   390
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
         Begin MSMask.MaskEdBox PCProduto 
            Height          =   225
            Left            =   750
            TabIndex        =   131
            Top             =   1035
            Width           =   1215
            _ExtentX        =   2143
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
            TabIndex        =   132
            Top             =   810
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
         Begin MSMask.MaskEdBox PCPrecoUnit 
            Height          =   225
            Left            =   4860
            TabIndex        =   133
            Top             =   780
            Width           =   1170
            _ExtentX        =   2064
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
            Left            =   4275
            TabIndex        =   134
            Top             =   405
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
         Begin MSFlexGridLib.MSFlexGrid GridPrecosCalculados 
            Height          =   1650
            Left            =   105
            TabIndex        =   136
            Top             =   195
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   2910
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Formação de preço do item selecionado acima"
         Height          =   2325
         Left            =   105
         TabIndex        =   115
         Top             =   2085
         Width           =   9165
         Begin VB.ComboBox FPSituacao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoVendaCro.ctx":0157
            Left            =   5880
            List            =   "OrcamentoVendaCro.ctx":0167
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   1650
            Width           =   1155
         End
         Begin VB.TextBox FPDescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1815
            MaxLength       =   250
            TabIndex        =   122
            Top             =   600
            Width           =   1935
         End
         Begin VB.ComboBox FPUnidMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OrcamentoVendaCro.ctx":0188
            Left            =   645
            List            =   "OrcamentoVendaCro.ctx":018A
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   1110
            Width           =   720
         End
         Begin MSMask.MaskEdBox FPQtde 
            Height          =   225
            Left            =   1500
            TabIndex        =   116
            Top             =   1155
            Width           =   825
            _ExtentX        =   1455
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
            TabIndex        =   117
            Top             =   1710
            Width           =   975
            _ExtentX        =   1720
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
            TabIndex        =   118
            Top             =   1695
            Width           =   870
            _ExtentX        =   1535
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
            TabIndex        =   119
            Top             =   1695
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FPProduto 
            Height          =   225
            Left            =   630
            TabIndex        =   121
            Top             =   585
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FPPrecoTotal 
            Height          =   225
            Left            =   4590
            TabIndex        =   124
            Top             =   1710
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridFormacaoPreco 
            Height          =   1935
            Left            =   105
            TabIndex        =   125
            Top             =   210
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   3413
            _Version        =   393216
         End
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
         TabIndex        =   114
         Top             =   4485
         Width           =   2145
      End
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
         Left            =   8055
         TabIndex        =   113
         Top             =   4470
         Width           =   1200
      End
   End
   Begin VB.CheckBox ImprimeOrcamentoGravacao 
      Caption         =   "Imprimir o orçamento ao gravar"
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
      Left            =   240
      TabIndex        =   106
      Top             =   210
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6210
      ScaleHeight     =   450
      ScaleWidth      =   3150
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   60
      Width           =   3210
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   579
         Picture         =   "OrcamentoVendaCro.ctx":018C
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   1098
         Picture         =   "OrcamentoVendaCro.ctx":028E
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1617
         Picture         =   "OrcamentoVendaCro.ctx":03E8
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   2136
         Picture         =   "OrcamentoVendaCro.ctx":0572
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2655
         Picture         =   "OrcamentoVendaCro.ctx":0AA4
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoEmail 
         Height          =   345
         Left            =   60
         Picture         =   "OrcamentoVendaCro.ctx":0C22
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Enviar email"
         Top             =   60
         Width           =   420
      End
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
      Left            =   3510
      TabIndex        =   98
      Top             =   210
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4620
      Index           =   3
      Left            =   90
      TabIndex        =   76
      Top             =   1065
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   3855
         Left            =   150
         TabIndex        =   78
         Top             =   435
         Width           =   8970
         Begin VB.ComboBox TipoDesconto3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3105
            TabIndex        =   83
            Top             =   1845
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            TabIndex        =   82
            Top             =   1530
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   81
            Top             =   1215
            Width           =   1965
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3960
            Picture         =   "OrcamentoVendaCro.ctx":15C4
            Style           =   1  'Graphical
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3960
            Picture         =   "OrcamentoVendaCro.ctx":161E
            Style           =   1  'Graphical
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7470
            TabIndex        =   84
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
            TabIndex        =   85
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
            TabIndex        =   86
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
            TabIndex        =   87
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
            TabIndex        =   88
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
            TabIndex        =   89
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
            TabIndex        =   90
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
            TabIndex        =   91
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
            TabIndex        =   92
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
            TabIndex        =   93
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
            TabIndex        =   94
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
            Left            =   2850
            TabIndex        =   95
            Top             =   240
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
            Height          =   2745
            Left            =   180
            TabIndex        =   96
            Top             =   675
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   4842
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
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
            Left            =   1020
            TabIndex        =   97
            Top             =   285
            Width           =   1740
         End
      End
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
         TabIndex        =   77
         Top             =   150
         Value           =   1  'Checked
         Width           =   3360
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5265
      Left            =   15
      TabIndex        =   107
      Top             =   570
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9287
      MultiRow        =   -1  'True
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
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "OrcamentoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTOrcamentoVenda
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTOrcamentoVenda
    Set objCT.objUserControl = Me
    
    'Cromaton
    Set objCT.gobjInfoUsu = New CTOrcVendaVGCro
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTOrcVendaCro

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

Private Sub BotaoKitVenda_Click()
    Call objCT.BotaoKitVenda_Click
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

Private Sub Escolhido_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Escolhido_GotFocus(objCT)
End Sub

Private Sub Escolhido_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Escolhido_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Escolhido_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Escolhido_Validate(objCT, Cancel)
End Sub

Private Sub Concorrente_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Concorrente_Change(objCT)
End Sub

Private Sub Concorrente_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Concorrente_GotFocus(objCT)
End Sub

Private Sub Concorrente_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Concorrente_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Concorrente_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Concorrente_Validate(objCT, Cancel)
End Sub

Private Sub CondPagtoItem_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondPagtoItem_Change(objCT)
End Sub

Private Sub CondPagtoItem_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondPagtoItem_GotFocus(objCT)
End Sub

Private Sub CondPagtoItem_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondPagtoItem_KeyPress(objCT, KeyAscii)
End Sub

Private Sub CondPagtoItem_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondPagtoItem_Validate(objCT, Cancel)
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
    Call objCT.Codigo_Validate(Cancel)
End Sub


