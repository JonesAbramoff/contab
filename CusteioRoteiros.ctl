VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl CusteioRoteiros 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1020
      Width           =   9225
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Calcular Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   105
         ToolTipText     =   "Abre Browse dos Roteiros de Produção cadastrados"
         Top             =   3975
         Width           =   1545
      End
      Begin VB.Frame FrameDatas 
         Caption         =   "Datas"
         Height          =   1170
         Left            =   6615
         TabIndex        =   96
         Top             =   0
         Width           =   2580
         Begin MSComCtl2.UpDown UpDownDataCusteio 
            Height          =   300
            Left            =   2130
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCusteio 
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   315
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataValidade 
            Height          =   300
            Left            =   2130
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   735
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   315
            Left            =   960
            TabIndex        =   5
            Top             =   735
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataValidade 
            AutoSize        =   -1  'True
            Caption         =   "Validade:"
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
            TabIndex        =   98
            Top             =   795
            Width           =   810
         End
         Begin VB.Label LabelData 
            AutoSize        =   -1  'True
            Caption         =   "Custeio:"
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
            Left            =   195
            TabIndex        =   97
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.TextBox Observacao 
         Height          =   1095
         Left            =   1545
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   2715
         Width           =   4110
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2385
         Picture         =   "CusteioRoteiros.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Numeração Automática"
         Top             =   30
         Width           =   300
      End
      Begin VB.Frame FrameCustoDetalhado 
         Caption         =   "Custo"
         Height          =   2115
         Left            =   5895
         TabIndex        =   82
         Top             =   1725
         Width           =   3300
         Begin VB.Label Label1 
            Caption         =   "Total:"
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
            Left            =   1200
            TabIndex        =   100
            Top             =   1605
            Width           =   555
         End
         Begin VB.Label LabelDescCustoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1755
            TabIndex        =   99
            Top             =   1605
            Width           =   1350
         End
         Begin VB.Label LabelCustoMO 
            Caption         =   "Mão de Obra:"
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
            Left            =   540
            TabIndex        =   88
            Top             =   1170
            Width           =   1200
         End
         Begin VB.Label LabelDescCustoMO 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1755
            TabIndex        =   87
            Top             =   1155
            Width           =   1350
         End
         Begin VB.Label LabelCustoInsMaq 
            Caption         =   "Insumos Máquina:"
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
            Left            =   165
            TabIndex        =   86
            Top             =   720
            Width           =   1560
         End
         Begin VB.Label LabelDescInsMaq 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1755
            TabIndex        =   85
            Top             =   705
            Width           =   1350
         End
         Begin VB.Label LabelCustoMP 
            Caption         =   "Insumos Kit:"
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
            Left            =   645
            TabIndex        =   84
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label LabelDescCustoMP 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1755
            TabIndex        =   83
            Top             =   270
            Width           =   1350
         End
      End
      Begin VB.CommandButton BotaoRoteiro 
         Caption         =   "Roteiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7635
         TabIndex        =   18
         ToolTipText     =   "Abre Browse dos Roteiros de Produção cadastrados"
         Top             =   3960
         Width           =   1545
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   885
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   450
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoRaiz 
         Height          =   315
         Left            =   1530
         TabIndex        =   6
         Top             =   1320
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Versao 
         Height          =   315
         Left            =   1530
         TabIndex        =   7
         Top             =   1770
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label LabelDescUM 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4050
         TabIndex        =   102
         Top             =   2235
         Width           =   990
      End
      Begin VB.Label LabelQuantidade 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   450
         TabIndex        =   104
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label LabelUM 
         Caption         =   "Un. Medida:"
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
         Left            =   2940
         TabIndex        =   103
         Top             =   2295
         Width           =   1260
      End
      Begin VB.Label LabelDescQuantidade 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1545
         TabIndex        =   101
         Top             =   2235
         Width           =   990
      End
      Begin VB.Label LabelObservacao 
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
         Height          =   315
         Left            =   375
         TabIndex        =   95
         Top             =   2715
         Width           =   1125
      End
      Begin VB.Label LabelVersao 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   810
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   94
         Top             =   1815
         Width           =   660
      End
      Begin VB.Label DescricaoProd 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3075
         TabIndex        =   93
         Top             =   1320
         Width           =   6105
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   735
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   92
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label LabelNomeReduzido 
         Caption         =   "Nome Reduzido:"
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
         Left            =   75
         TabIndex        =   91
         Top             =   495
         Width           =   1410
      End
      Begin VB.Label LabelCodigo 
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
         Height          =   315
         Left            =   840
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   90
         Top             =   45
         Width           =   690
      End
      Begin VB.Label LabelDescricao 
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
         Height          =   315
         Left            =   570
         TabIndex        =   89
         Top             =   915
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   2
      Left            =   90
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame2 
         Caption         =   "Insumos Usados no Kit"
         Height          =   3720
         Index           =   1
         Left            =   210
         TabIndex        =   24
         Top             =   690
         Width           =   8865
         Begin VB.TextBox ObservacaoKit 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6255
            MaxLength       =   255
            TabIndex        =   57
            Top             =   1185
            Width           =   3540
         End
         Begin MSMask.MaskEdBox CustoTotalKit 
            Height          =   315
            Left            =   5295
            TabIndex        =   56
            Top             =   1185
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox VariacaoKit 
            Height          =   315
            Left            =   6630
            TabIndex        =   55
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoUnitKit 
            Height          =   315
            Left            =   5670
            TabIndex        =   54
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
         Begin VB.ComboBox UMProdKit 
            Height          =   315
            Left            =   3135
            TabIndex        =   50
            Top             =   1620
            Width           =   600
         End
         Begin VB.TextBox DescricaoProdKit 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1365
            TabIndex        =   49
            Top             =   1605
            Width           =   1770
         End
         Begin MSMask.MaskEdBox CustoUnitCalculadoKit 
            Height          =   315
            Left            =   4710
            TabIndex        =   51
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoKit 
            Height          =   315
            Left            =   630
            TabIndex        =   52
            Top             =   1605
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProdKit 
            Height          =   315
            Left            =   3975
            TabIndex        =   53
            Top             =   1605
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridKit 
            Height          =   2820
            Left            =   165
            TabIndex        =   14
            Top             =   315
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4974
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label CustoTotalInsumosKit 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7215
            TabIndex        =   26
            Top             =   3360
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Left            =   6105
            TabIndex        =   25
            Top             =   3390
            Width           =   1050
         End
      End
      Begin VB.Frame SSFrame1 
         Height          =   555
         Index           =   1
         Left            =   210
         TabIndex        =   34
         Top             =   90
         Width           =   8865
         Begin VB.Label LabelVersaoInsumosKit 
            Height          =   210
            Left            =   6795
            TabIndex        =   77
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label LabelProdutoInsumosKit 
            Height          =   210
            Left            =   975
            TabIndex        =   76
            Top             =   210
            Width           =   4935
         End
         Begin VB.Label VersaoLabel 
            Height          =   210
            Index           =   1
            Left            =   6735
            TabIndex        =   38
            Top             =   195
            Width           =   1665
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   0
            Left            =   6060
            TabIndex        =   37
            Top             =   195
            Width           =   810
         End
         Begin VB.Label ProdutoLabel 
            Height          =   210
            Index           =   1
            Left            =   1170
            TabIndex        =   36
            Top             =   180
            Width           =   4665
         End
         Begin VB.Label Label30 
            Caption         =   "Produto:"
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
            Index           =   1
            Left            =   195
            TabIndex        =   35
            Top             =   210
            Width           =   810
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4605
      Index           =   4
      Left            =   90
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame SSFrame1 
         Height          =   555
         Index           =   2
         Left            =   210
         TabIndex        =   44
         Top             =   90
         Width           =   8865
         Begin VB.Label LabelVersaoMO 
            Height          =   210
            Left            =   6795
            TabIndex        =   81
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label LabelProdutoMO 
            Height          =   210
            Left            =   975
            TabIndex        =   80
            Top             =   210
            Width           =   4935
         End
         Begin VB.Label Label30 
            Caption         =   "Produto:"
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
            Index           =   5
            Left            =   195
            TabIndex        =   48
            Top             =   210
            Width           =   810
         End
         Begin VB.Label ProdutoLabel 
            Height          =   210
            Index           =   3
            Left            =   1155
            TabIndex        =   47
            Top             =   180
            Width           =   4665
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   4
            Left            =   6060
            TabIndex        =   46
            Top             =   195
            Width           =   810
         End
         Begin VB.Label VersaoLabel 
            Height          =   210
            Index           =   3
            Left            =   6735
            TabIndex        =   45
            Top             =   210
            Width           =   1665
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mão de Obra Usadas"
         Height          =   3720
         Index           =   5
         Left            =   210
         TabIndex        =   31
         Top             =   690
         Width           =   8865
         Begin VB.TextBox ObservacaoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6330
            MaxLength       =   255
            TabIndex        =   69
            Top             =   1020
            Width           =   3540
         End
         Begin VB.ComboBox UMMO 
            Height          =   315
            Left            =   3210
            TabIndex        =   68
            Top             =   1455
            Width           =   600
         End
         Begin VB.TextBox DescricaoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1440
            TabIndex        =   67
            Top             =   1440
            Width           =   1770
         End
         Begin MSMask.MaskEdBox CustoTotalMO 
            Height          =   315
            Left            =   5370
            TabIndex        =   70
            Top             =   1020
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox VariacaoMO 
            Height          =   315
            Left            =   6705
            TabIndex        =   71
            Top             =   1440
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoUnitMO 
            Height          =   315
            Left            =   5745
            TabIndex        =   72
            Top             =   1440
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox CustoUnitCalculadoMO 
            Height          =   315
            Left            =   4785
            TabIndex        =   73
            Top             =   1440
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox TipoMO 
            Height          =   315
            Left            =   705
            TabIndex        =   74
            Top             =   1440
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   3
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeMO 
            Height          =   315
            Left            =   4050
            TabIndex        =   75
            Top             =   1440
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaoDeObra 
            Height          =   2820
            Left            =   165
            TabIndex        =   20
            Top             =   315
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4974
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Index           =   2
            Left            =   6105
            TabIndex        =   33
            Top             =   3390
            Width           =   1050
         End
         Begin VB.Label CustoTotalMaoDeObra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7215
            TabIndex        =   32
            Top             =   3360
            Width           =   1500
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   3
      Left            =   90
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame2 
         Caption         =   "Insumos Usados na Máquina"
         Height          =   3720
         Index           =   4
         Left            =   210
         TabIndex        =   28
         Top             =   690
         Width           =   8865
         Begin VB.TextBox DescricaoProdMaq 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1395
            TabIndex        =   63
            Top             =   1650
            Width           =   1770
         End
         Begin VB.ComboBox UMProdMaq 
            Height          =   315
            Left            =   3165
            TabIndex        =   62
            Top             =   1665
            Width           =   600
         End
         Begin VB.TextBox ObservacaoMaq 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6285
            MaxLength       =   255
            TabIndex        =   58
            Top             =   1230
            Width           =   3540
         End
         Begin MSMask.MaskEdBox CustoTotalMaq 
            Height          =   315
            Left            =   5325
            TabIndex        =   59
            Top             =   1230
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox VariacaoMaq 
            Height          =   315
            Left            =   6660
            TabIndex        =   60
            Top             =   1650
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoUnitMaq 
            Height          =   315
            Left            =   5700
            TabIndex        =   61
            Top             =   1650
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox CustoUnitCalculadoMaq 
            Height          =   315
            Left            =   4740
            TabIndex        =   64
            Top             =   1650
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox ProdutoMaq 
            Height          =   315
            Left            =   660
            TabIndex        =   65
            Top             =   1650
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProdMaq 
            Height          =   315
            Left            =   4005
            TabIndex        =   66
            Top             =   1650
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaquinas 
            Height          =   2820
            Left            =   165
            TabIndex        =   15
            Top             =   316
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4974
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Left            =   6105
            TabIndex        =   30
            Top             =   3390
            Width           =   1050
         End
         Begin VB.Label CustoTotalInsumosMaquina 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7215
            TabIndex        =   29
            Top             =   3360
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame1 
         Height          =   555
         Index           =   0
         Left            =   210
         TabIndex        =   39
         Top             =   90
         Width           =   8865
         Begin VB.Label LabelVersaoInsumosMaq 
            Height          =   210
            Left            =   6795
            TabIndex        =   79
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label LabelProdutoInsumosMaq 
            Height          =   210
            Left            =   975
            TabIndex        =   78
            Top             =   210
            Width           =   4935
         End
         Begin VB.Label Label30 
            Caption         =   "Produto:"
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
            Index           =   3
            Left            =   195
            TabIndex        =   43
            Top             =   210
            Width           =   810
         End
         Begin VB.Label ProdutoLabel 
            Height          =   210
            Index           =   2
            Left            =   1170
            TabIndex        =   42
            Top             =   180
            Width           =   4665
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   2
            Left            =   6060
            TabIndex        =   41
            Top             =   195
            Width           =   810
         End
         Begin VB.Label VersaoLabel 
            Height          =   210
            Index           =   2
            Left            =   6720
            TabIndex        =   40
            Top             =   195
            Width           =   1665
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7275
      ScaleHeight     =   450
      ScaleWidth      =   2055
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   90
      Width           =   2115
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1590
         Picture         =   "CusteioRoteiros.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1080
         Picture         =   "CusteioRoteiros.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   555
         Picture         =   "CusteioRoteiros.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   45
         Picture         =   "CusteioRoteiros.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5070
      Left            =   60
      TabIndex        =   12
      Top             =   570
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   8943
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insumos Kit"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insumos Máquina"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mão de Obra"
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
Attribute VB_Name = "CusteioRoteiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iGridKitAlterado As Integer
Dim iGridMaqAlterado As Integer
Dim iGridMOAlterado As Integer

'Grid de Kits
Dim objGridKit As AdmGrid
Dim iGrid_ProdutoKit_Col As Integer
Dim iGrid_DescricaoProdKit_Col As Integer
Dim iGrid_UMProdKit_Col As Integer
Dim iGrid_QuantidadeProdKit_Col As Integer
Dim iGrid_CustoUnitCalculadoKit_Col As Integer
Dim iGrid_CustoUnitKit_Col As Integer
Dim iGrid_VariacaoKit_Col As Integer
Dim iGrid_CustoTotalKit_Col As Integer
Dim iGrid_ObservacaoKit_Col As Integer

'Grid de Maquinas
Dim objGridMaquinas As AdmGrid
Dim iGrid_ProdutoMaq_Col As Integer
Dim iGrid_DescricaoProdMaq_Col As Integer
Dim iGrid_UMProdMaq_Col As Integer
Dim iGrid_QuantidadeProdMaq_Col As Integer
Dim iGrid_CustoUnitCalculadoMaq_Col As Integer
Dim iGrid_CustoUnitMaq_Col As Integer
Dim iGrid_VariacaoMaq_Col As Integer
Dim iGrid_CustoTotalMaq_Col As Integer
Dim iGrid_ObservacaoMaq_Col As Integer

'Grid de MaoDeObra
Dim objGridMaoDeObra As AdmGrid
Dim iGrid_TipoMO_Col As Integer
Dim iGrid_DescricaoMo_Col As Integer
Dim iGrid_UMMO_Col As Integer
Dim iGrid_QuantidadeMO_Col As Integer
Dim iGrid_CustoUnitCalculadoMO_Col As Integer
Dim iGrid_CustoUnitMO_Col As Integer
Dim iGrid_VariacaoMO_Col As Integer
Dim iGrid_CustoTotalMO_Col As Integer
Dim iGrid_ObservacaoMO_Col As Integer

Private WithEvents objEventoCusteioRoteiro As AdmEvento
Attribute objEventoCusteioRoteiro.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1
Private WithEvents objEventoRoteiro As AdmEvento
Attribute objEventoRoteiro.VB_VarHelpID = -1

Dim gobjCusteioRoteiro As ClassCusteioRoteiro

Private Const TAB_Inicial = 1
Private Const TAB_InsumosKit = 2
Private Const TAB_InsumosMaq = 3
Private Const TAB_MaoDeObra = 4

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Custeio de Roteiros de Fabricação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CusteioRoteiros"

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

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCusteioRoteiro As New ClassCusteioRoteiro
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Critica dados da tela
    If Len(Trim(Codigo.Text)) = 0 Then gError 139283

    Set objCusteioRoteiro = New ClassCusteioRoteiro
    
    objCusteioRoteiro.lCodigo = StrParaLong(Trim(Codigo.Text))

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CUSTEIOROTEIRO", Trim(Codigo.Text))

    If vbMsgRes = vbYes Then

        'Exclui o CusteioRoteiro
        lErro = CF("CusteioRoteiro_Exclui", objCusteioRoteiro)
        If lErro <> SUCESSO Then gError 139284
    
        'Limpa Tela
        Call Limpa_Tela_CusteioRoteiro

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 139283
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CUSTEIOROTEIRO_NAO_PREENCHIDO", gErr)

        Case 139284

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158488)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158489)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Grava o CusteioRot
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 139285

    'Limpa Tela
    Call Limpa_Tela_CusteioRoteiro
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 139285

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158490)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 139286
    
    Call Limpa_Tela_CusteioRoteiro

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 139286

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158491)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para um Custeio de Roteiro
    lErro = CF("CusteioRoteiro_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 139287
    
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 139287
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158492)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoRoteiro_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelProdutoRaiz_Click

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 139288

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(Versao.Text)) <> 0 Then
        objRoteirosDeFabricacao.sVersao = Versao.Text
    End If

    Call Chama_Tela("RoteirosDeFabricacaoLista", colSelecao, objRoteirosDeFabricacao, objEventoRoteiro)

    Exit Sub

Erro_LabelProdutoRaiz_Click:

    Select Case gErr

        Case 139288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158493)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRoteiroFabricacao As ClassRoteirosDeFabricacao
Dim dCustoTotal As Double

On Error GoTo Erro_BotaoTrazer_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'se o produto não está preenchido... erro
    If Len(Trim(ProdutoRaiz.Text)) = 0 Then gError 139289
    
    'se a versão não está preenchida... erro
    If Len(Trim(Versao.Text)) = 0 Then gError 139290
    
    'formata o codigo do produto que esta na tela
    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 139291
    
    'inicializa o objeto Roteiro
    Set objRoteiroFabricacao = New ClassRoteirosDeFabricacao
    
    objRoteiroFabricacao.sProdutoRaiz = sProdutoFormatado
    objRoteiroFabricacao.sVersao = Trim(Versao.Text)
    
    'Le o Roteiro
    lErro = CF("RoteirosDeFabricacao_Le", objRoteiroFabricacao)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 139292
    
    'se não encontrou... erro
    If lErro <> SUCESSO Then gError 139293
    
    'Se mudou o produto ou a versão... refaz
    If gobjCusteioRoteiro.sProduto <> sProdutoFormatado Or _
        gobjCusteioRoteiro.sVersao <> Trim(Versao.Text) Then
        
        Call Limpa_DadosCusteio
        
        Set gobjCusteioRoteiro = New ClassCusteioRoteiro
    
        gobjCusteioRoteiro.sProduto = sProdutoFormatado
        gobjCusteioRoteiro.sVersao = Trim(Versao.Text)
        gobjCusteioRoteiro.dQuantidade = objRoteiroFabricacao.dQuantidade
        gobjCusteioRoteiro.sUMedida = objRoteiroFabricacao.sUM
            
        'Preenche Dados do Custeio
        LabelDescQuantidade.Caption = Formata_Estoque(gobjCusteioRoteiro.dQuantidade)
        LabelDescUM.Caption = gobjCusteioRoteiro.sUMedida
        
        'Calcula as necessidades de Produção do Item
        lErro = CalculaNecessidadesProducao(objRoteiroFabricacao)
        If lErro <> SUCESSO Then gError 139294
        
        LabelDescCustoMP.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosKit, "Standard")
        LabelDescInsMaq.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosMaq, "Standard")
        LabelDescCustoMO.Caption = Format(gobjCusteioRoteiro.dCustoTotalMaoDeObra, "Standard")
        
        dCustoTotal = gobjCusteioRoteiro.dCustoTotalInsumosKit + gobjCusteioRoteiro.dCustoTotalInsumosMaq + gobjCusteioRoteiro.dCustoTotalMaoDeObra
        LabelDescCustoTotal.Caption = Format(dCustoTotal, "Standard")
                        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoTrazer_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 139289
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTORAIZ_ROTFABRICACAO_NAO_PREENCHIDO", gErr)
        
        Case 139290
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSDEFABRICACAO_NAO_PREENCHIDO", gErr)
        
        Case 139291, 139292, 139294
            'erros tratados nas rotinas chamadas
        
        Case 139293
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSDEFABRICACAO_NAO_CADASTRADO", gErr, objRoteiroFabricacao.sProdutoRaiz, objRoteiroFabricacao.sVersao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158494)
        
    End Select

    Exit Sub
    
End Sub

Private Sub DataCusteio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCusteio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCusteio, iAlterado)
    
End Sub

Private Sub DataCusteio_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_DataCusteio_Validate

    'Verifica se DataCusteio está preenchida
    If Len(Trim(DataCusteio.Text)) <> 0 Then

        lErro = Data_Critica(DataCusteio.Text)
        If lErro <> SUCESSO Then gError 139295

    End If

    Exit Sub

Erro_DataCusteio_Validate:

    Cancel = True

    Select Case gErr

        Case 139295

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158495)

    End Select

    Exit Sub

End Sub

Private Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataValidade, iAlterado)
    
End Sub


Private Sub DataValidade_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_DataValidade_Validate

    'Verifica se DataValidade está preenchida
    If Len(Trim(DataValidade.Text)) <> 0 Then

        lErro = Data_Critica(DataValidade.Text)
        If lErro <> SUCESSO Then gError 139296

    End If

    Exit Sub

Erro_DataValidade_Validate:

    Cancel = True

    Select Case gErr

        Case 139296

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158496)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objCusteioRoteiro As New ClassCusteioRoteiro
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then
        
        objCusteioRoteiro.lCodigo = StrParaLong(Trim(Codigo.Text))
        
    End If

    Call Chama_Tela("CusteioRoteirosLista", colSelecao, objCusteioRoteiro, objEventoCusteioRoteiro)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158497)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 139297

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 139297

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158498)

    End Select

    Exit Sub

End Sub

Private Sub LabelVersao_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelVersao_Click

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 139298

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objKit.sProdutoRaiz = sProdutoFormatado
        
        If Len(Trim(Versao.ClipText)) > 0 Then objKit.sVersao = Versao.Text
            
        colSelecao.Add sProdutoFormatado
        
        Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
    Else
         gError 139299
         
    End If

    Exit Sub

Erro_LabelVersao_Click:

    Select Case gErr

        Case 139298
        
        Case 139299
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158499)

    End Select

    Exit Sub

End Sub


Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 139300
    
    ProdutoRaiz.Text = sProdutoMascarado
    
    Call ProdutoRaiz_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
    
        Case 139300

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158500)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCusteioRoteiro_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCusteioRoteiro As New ClassCusteioRoteiro

On Error GoTo Erro_objEventoCusteioRoteiro_evSelecao

    Set objCusteioRoteiro = obj1
    
    Codigo.Text = objCusteioRoteiro.lCodigo
    
    lErro = Traz_CusteioRoteiro_Tela(objCusteioRoteiro)
    If lErro <> SUCESSO Then gError 138560
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub

Erro_objEventoCusteioRoteiro_evSelecao:

    Select Case gErr
    
        Case 138560

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158501)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRoteiro_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoRoteiro_evSelecao

    Set objRoteirosDeFabricacao = obj1
    
    lErro = Mascara_RetornaProdutoTela(objRoteirosDeFabricacao.sProdutoRaiz, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 139301
    
    ProdutoRaiz.Text = sProdutoMascarado

    ProdutoRaiz_Validate (bSGECancelDummy)

    Versao.Text = objRoteirosDeFabricacao.sVersao

    Call BotaoTrazer_Click

    Me.Show

    Exit Sub

Erro_objEventoRoteiro_evSelecao:

    Select Case gErr

        Case 139301

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158502)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1
    
    Versao.Text = objKit.sVersao
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158503)

    End Select

    Exit Sub
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchido
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'Critica a Codigo
        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 139302

    End If
            
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 139302
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158504)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoRaiz_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoRaiz_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoRaiz, iAlterado)
    
End Sub

Private Sub ProdutoRaiz_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProduto As String

On Error GoTo Erro_Produto_Validate
   
    DescricaoProd.Caption = ""
    
    If Len(Trim(ProdutoRaiz.ClipText)) <> 0 Then
    
        sProduto = ProdutoRaiz.Text
    
        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 139304
    
        'se o produto não estiver cadastrado ==> erro
        If lErro <> SUCESSO Then gError 139305
    
        'se o produto for gerencial, não pode fazer parte de um kit
        If objProduto.iGerencial = GERENCIAL Then gError 139306
        
        DescricaoProd.Caption = objProduto.sDescricao
            
    End If
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 139304
            'erro tratado na rotina chamada
            
        Case 139305
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 139306
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158505)

    End Select

    Exit Sub

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
        
        'Se Frame selecionado foi diferente de Inicial
        If TabStrip1.SelectedItem.Index <> TAB_Inicial Then
                        
            Call Trata_Custeio(TabStrip1.SelectedItem.Index)
                
        End If
                
    End If

End Sub

Private Sub UpDownDataCusteio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCusteio_DownClick

    DataCusteio.SetFocus

    If Len(DataCusteio.ClipText) > 0 Then

        sData = DataCusteio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 139307

        DataCusteio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCusteio_DownClick:

    Select Case gErr

        Case 139307

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158506)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCusteio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCusteio_UpClick

    DataCusteio.SetFocus

    If Len(Trim(DataCusteio.ClipText)) > 0 Then

        sData = DataCusteio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 139308

        DataCusteio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCusteio_UpClick:

    Select Case gErr

        Case 139308

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158507)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidade_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataValidade_DownClick

    DataValidade.SetFocus

    If Len(DataValidade.ClipText) > 0 Then

        sData = DataValidade.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 139309

        DataValidade.Text = sData

    End If

    Exit Sub

Erro_UpDownDataValidade_DownClick:

    Select Case gErr

        Case 139309

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158508)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidade_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataValidade_UpClick

    DataValidade.SetFocus

    If Len(Trim(DataValidade.ClipText)) > 0 Then

        sData = DataValidade.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 139310

        DataValidade.Text = sData

    End If

    Exit Sub

Erro_UpDownDataValidade_UpClick:

    Select Case gErr

        Case 139310

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158509)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
        
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

On Error GoTo Erro_Form_Unload
    
    Call ComandoSeta_Liberar(Me.Name)

    Set objEventoCusteioRoteiro = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVersao = Nothing
    Set objEventoRoteiro = Nothing

    Set gobjCusteioRoteiro = Nothing
    
    Set objGridKit = Nothing
    Set objGridMaquinas = Nothing
    Set objGridMaoDeObra = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158510)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
   
    iFrameAtual = 1
    
    Set objEventoCusteioRoteiro = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Set objEventoRoteiro = New AdmEvento

    Set gobjCusteioRoteiro = New ClassCusteioRoteiro
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoRaiz)
    If lErro <> SUCESSO Then gError 139311
    
    DataCusteio.PromptInclude = False
    DataCusteio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCusteio.PromptInclude = True
    
    'Grid Kit
    Set objGridKit = New AdmGrid
    
    'tela em questão
    Set objGridKit.objForm = Me
    
    lErro = Inicializa_GridKit(objGridKit)
    If lErro <> SUCESSO Then gError 139312
    
    'Grid Maquinas
    Set objGridMaquinas = New AdmGrid
    
    'tela em questão
    Set objGridMaquinas.objForm = Me
    
    lErro = Inicializa_GridMaquinas(objGridMaquinas)
    If lErro <> SUCESSO Then gError 139313
    
    'Grid MaoDeObra
    Set objGridMaoDeObra = New AdmGrid
    
    'tela em questão
    Set objGridMaoDeObra.objForm = Me
    
    lErro = Inicializa_GridMaoDeObra(objGridMaoDeObra)
    If lErro <> SUCESSO Then gError 139314
    
    iAlterado = 0
    iGridKitAlterado = 0
    iGridMaqAlterado = 0
    iGridMOAlterado = 0
            
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 139311 To 139314
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158511)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCusteioRoteiro As ClassCusteioRoteiro) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCusteioRoteiro Is Nothing) Then
    
        lErro = Traz_CusteioRoteiro_Tela(objCusteioRoteiro)
        If lErro <> SUCESSO And lErro <> 139382 Then gError 139315
        
        If lErro <> SUCESSO Then
        
            ProdutoRaiz.Text = objCusteioRoteiro.sProduto
            Versao.Text = objCusteioRoteiro.sVersao
        
        End If

    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 139315

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158512)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridKit
        If objGridInt.objGrid.Name = GridKit.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CustoUnitKit_Col

                    lErro = Saida_Celula_CustoUnitKit(objGridInt)
                    If lErro <> SUCESSO Then gError 139316

                Case iGrid_VariacaoKit_Col

                    lErro = Saida_Celula_VariacaoKit(objGridInt)
                    If lErro <> SUCESSO Then gError 139317
                
                Case iGrid_ObservacaoKit_Col

                    lErro = Saida_Celula_ObservacaoKit(objGridInt)
                    If lErro <> SUCESSO Then gError 139318
                    
            End Select
        
        'GridMaquinas
        ElseIf objGridInt.objGrid.Name = GridMaquinas.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CustoUnitMaq_Col

                    lErro = Saida_Celula_CustoUnitMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 139319

                Case iGrid_VariacaoMaq_Col

                    lErro = Saida_Celula_VariacaoMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 139320
                
                Case iGrid_ObservacaoMaq_Col

                    lErro = Saida_Celula_ObservacaoMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 139321
                    
            End Select
                    
        'GridMaoDeObra
        ElseIf objGridInt.objGrid.Name = GridMaoDeObra.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CustoUnitMO_Col

                    lErro = Saida_Celula_CustoUnitMO(objGridInt)
                    If lErro <> SUCESSO Then gError 139322

                Case iGrid_VariacaoMO_Col

                    lErro = Saida_Celula_VariacaoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 139323
                
                Case iGrid_ObservacaoMO_Col

                    lErro = Saida_Celula_ObservacaoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 139324
                    
            End Select
                
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 139325

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 139316 To 139324
            'erros tratatos nas rotinas chamadas
        
        Case 139325
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158513)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iProdutoPreenchidoKit As Integer
Dim iProdutoPreenchidoMaq As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Guardo o valor do Codigo do Produto de InsumosKit
    sProduto = GridKit.TextMatrix(GridKit.Row, iGrid_ProdutoKit_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchidoKit)
    If lErro <> SUCESSO Then gError 139326
    
    'Guardo o valor do Codigo do Produto de InsumosMaquina
    sProduto = GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_ProdutoMaq_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchidoMaq)
    If lErro <> SUCESSO Then gError 139327
    
    'Grid Kit
    If objControl.Name = "ProdutoKit" Or _
            objControl.Name = "DescricaoProdKit" Or _
            objControl.Name = "UMProdKit" Or _
            objControl.Name = "QuantidadeProdKit" Or _
            objControl.Name = "CustoUnitCalculadoKit" Or _
            objControl.Name = "CustoTotalKit" Then

        objControl.Enabled = False
    
    ElseIf objControl.Name = "CustoUnitKit" Or _
            objControl.Name = "VariacaoKit" Or _
            objControl.Name = "ObservacaoKit" Then
        
        If iProdutoPreenchidoKit = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If

    'Grid Maquina
    ElseIf objControl.Name = "ProdutoMaq" Or _
            objControl.Name = "DescricaoProdMaq" Or _
            objControl.Name = "UMProdMaq" Or _
            objControl.Name = "QuantidadeProdMaq" Or _
            objControl.Name = "CustoUnitCalculadoMaq" Or _
            objControl.Name = "CustoTotalMaq" Then

        objControl.Enabled = False
    
    ElseIf objControl.Name = "CustoUnitMaq" Or _
            objControl.Name = "VariacaoMaq" Or _
            objControl.Name = "ObservacaoMaq" Then
        
        If iProdutoPreenchidoMaq = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If

    'Grid MaoDeObra
    ElseIf objControl.Name = "TipoMO" Or _
            objControl.Name = "DescricaoMO" Or _
            objControl.Name = "UMMO" Or _
            objControl.Name = "QuantidadeMO" Or _
            objControl.Name = "CustoUnitCalculadoMO" Or _
            objControl.Name = "CustoTotalMO" Then

        objControl.Enabled = False
    
    ElseIf objControl.Name = "CustoUnitMO" Or _
            objControl.Name = "VariacaoMO" Or _
            objControl.Name = "ObservacaoMO" Then
        
        If Len(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_TipoMO_Col)) > 0 Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If

    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 139326, 139327

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158514)

    End Select

    Exit Sub

End Sub

Private Sub GridKit_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridKit, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridKit, iAlterado)
        End If

End Sub

Private Sub GridKit_GotFocus()
    
    Call Grid_Recebe_Foco(objGridKit)

End Sub

Private Sub GridKit_EnterCell()

    Call Grid_Entrada_Celula(objGridKit, iAlterado)

End Sub

Private Sub GridKit_LeaveCell()
    
    Call Saida_Celula(objGridKit)

End Sub

Private Sub GridKit_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridKit, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridKit, iAlterado)
    End If

End Sub

Private Sub GridKit_RowColChange()

    Call Grid_RowColChange(objGridKit)

End Sub

Private Sub GridKit_Scroll()

    Call Grid_Scroll(objGridKit)

End Sub

Private Function Inicializa_GridKit(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Unit.Calc.")
    objGrid.colColuna.Add ("Unitário")
    objGrid.colColuna.Add ("Variação")
    objGrid.colColuna.Add ("Custo Total")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoKit.Name)
    objGrid.colCampo.Add (DescricaoProdKit.Name)
    objGrid.colCampo.Add (UMProdKit.Name)
    objGrid.colCampo.Add (QuantidadeProdKit.Name)
    objGrid.colCampo.Add (CustoUnitCalculadoKit.Name)
    objGrid.colCampo.Add (CustoUnitKit.Name)
    objGrid.colCampo.Add (VariacaoKit.Name)
    objGrid.colCampo.Add (CustoTotalKit.Name)
    objGrid.colCampo.Add (ObservacaoKit.Name)

    'Colunas do Grid
    iGrid_ProdutoKit_Col = 1
    iGrid_DescricaoProdKit_Col = 2
    iGrid_UMProdKit_Col = 3
    iGrid_QuantidadeProdKit_Col = 4
    iGrid_CustoUnitCalculadoKit_Col = 5
    iGrid_CustoUnitKit_Col = 6
    iGrid_VariacaoKit_Col = 7
    iGrid_CustoTotalKit_Col = 8
    iGrid_ObservacaoKit_Col = 9

    objGrid.objGrid = GridKit

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 21

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridKit.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridKit = SUCESSO

End Function

Private Sub ProdutoKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub ProdutoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub ProdutoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = ProdutoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProdKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProdKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub DescricaoProdKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub DescricaoProdKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = DescricaoProdKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProdKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProdKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub UMProdKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub UMProdKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = UMProdKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProdKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProdKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub QuantidadeProdKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub QuantidadeProdKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = QuantidadeProdKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitCalculadoKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitCalculadoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub CustoUnitCalculadoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub CustoUnitCalculadoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = CustoUnitCalculadoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitKit_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridKitAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub CustoUnitKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub CustoUnitKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = CustoUnitKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VariacaoKit_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridKitAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VariacaoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub VariacaoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub VariacaoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = VariacaoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoTotalKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoTotalKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub CustoTotalKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub CustoTotalKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = CustoTotalKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoKit_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridKitAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub ObservacaoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub ObservacaoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = ObservacaoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CustoUnitKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CustoUnitKit do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dVariacao As Double
Dim dCustoTotalKit As Double
Dim dAjusteCustoKit As Double

On Error GoTo Erro_Saida_Celula_CustoUnitKit

    Set objGridInt.objControle = CustoUnitKit

    'Se o campo foi preenchido
    If Len(Trim(CustoUnitKit.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(CustoUnitKit.Text)
        If lErro <> SUCESSO Then gError 139328
    
    End If
    
    If iGridKitAlterado = REGISTRO_ALTERADO Then
    
        'Calcula a Variação
        lErro = Calcula_Variacao(StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitCalculadoKit_Col)), StrParaDbl(CustoUnitKit.Text), dVariacao, iGrid_VariacaoKit_Col)
        If lErro <> SUCESSO Then gError 139329
        
        'Altera a Variação no grid
        If dVariacao <> 0 Then
           GridKit.TextMatrix(GridKit.Row, iGrid_VariacaoKit_Col) = Format(dVariacao, "Percent")
        Else
           GridKit.TextMatrix(GridKit.Row, iGrid_VariacaoKit_Col) = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalKit = StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_QuantidadeProdKit_Col)) * StrParaDbl(CustoUnitKit.Text)
        
        'Calcula a diferença a ajustar
        dAjusteCustoKit = dCustoTotalKit - StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col))
        
        'Altera o Custo Total do Item no grid
        GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col) = Format(dCustoTotalKit, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosKit(CustoUnitKit.Text, iGrid_CustoUnitKit_Col, dAjusteCustoKit)
        If lErro <> SUCESSO Then gError 139330
        
        iGridKitAlterado = 0
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139331

    Saida_Celula_CustoUnitKit = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitKit:

    Saida_Celula_CustoUnitKit = gErr

    Select Case gErr
        
        Case 139328 To 139331
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158515)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VariacaoKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula VariacaoKit do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dCustoInformado As Double
Dim dVariacao As Double
Dim dCustoTotalKit As Double
Dim dAjusteCustoKit As Double

On Error GoTo Erro_Saida_Celula_VariacaoKit

    Set objGridInt.objControle = VariacaoKit
    
    If iGridKitAlterado = REGISTRO_ALTERADO Then

        'Ajusta e Converte o valor da tela
        If InStr(VariacaoKit.Text, "%") > 0 Then
        
            dVariacao = StrParaDbl(Left(VariacaoKit.Text, Len(VariacaoKit.Text) - 1))
        
        Else
        
            dVariacao = StrParaDbl(VariacaoKit.Text)
        
        End If
        
        'Calcula Variacao retornando o novo Custo Unitario
        lErro = Calcula_Variacao(StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitCalculadoKit_Col)), dCustoInformado, dVariacao, iGrid_CustoUnitKit_Col)
        If lErro <> SUCESSO Then gError 139332
        
        'Altera o Custo Unitario no grid
        If dCustoInformado <> 0 Then
           GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitKit_Col) = Format(dCustoInformado, gobjFAT.sFormatoPrecoUnitario)
        Else
           GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitKit_Col) = ""
        End If
        
        'se não tem Variacao ... Limpa no grid
        If dVariacao = 0 Then
           VariacaoKit.Text = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalKit = StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_QuantidadeProdKit_Col)) * dCustoInformado
        
        'Calcula a diferença a ajustar
        dAjusteCustoKit = dCustoTotalKit - StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col))
        
        'Altera o Custo Total no grid
        GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col) = Format(dCustoTotalKit, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosKit(CStr(dCustoInformado), iGrid_CustoUnitKit_Col, dAjusteCustoKit)
        If lErro <> SUCESSO Then gError 139333
        
        iGridKitAlterado = 0
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139334

    Saida_Celula_VariacaoKit = SUCESSO

    Exit Function

Erro_Saida_Celula_VariacaoKit:

    Saida_Celula_VariacaoKit = gErr

    Select Case gErr
        
        Case 139332 To 139334
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158516)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ObservacaoKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ObservacaoKit do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ObservacaoKit

    Set objGridInt.objControle = ObservacaoKit

    If iGridKitAlterado = REGISTRO_ALTERADO Then
    
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosKit(ObservacaoKit.Text, iGrid_ObservacaoKit_Col)
        If lErro <> SUCESSO Then gError 139335
            
        iGridKitAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139336

    Saida_Celula_ObservacaoKit = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoKit:

    Saida_Celula_ObservacaoKit = gErr

    Select Case gErr
        
        Case 139335, 139336
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158517)

    End Select

    Exit Function

End Function

Private Sub GridMaquinas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaquinas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
    End If

End Sub

Private Sub GridMaquinas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub GridMaquinas_EnterCell()

    Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)

End Sub

Private Sub GridMaquinas_LeaveCell()
    
    Call Saida_Celula(objGridMaquinas)

End Sub

Private Sub GridMaquinas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridMaquinas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
    End If

End Sub

Private Sub GridMaquinas_RowColChange()

    Call Grid_RowColChange(objGridMaquinas)

End Sub

Private Sub GridMaquinas_Scroll()

    Call Grid_Scroll(objGridMaquinas)

End Sub

Private Function Inicializa_GridMaquinas(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Unit.Calc.")
    objGrid.colColuna.Add ("Unitário")
    objGrid.colColuna.Add ("Variação")
    objGrid.colColuna.Add ("Custo Total")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoMaq.Name)
    objGrid.colCampo.Add (DescricaoProdMaq.Name)
    objGrid.colCampo.Add (UMProdMaq.Name)
    objGrid.colCampo.Add (QuantidadeProdMaq.Name)
    objGrid.colCampo.Add (CustoUnitCalculadoMaq.Name)
    objGrid.colCampo.Add (CustoUnitMaq.Name)
    objGrid.colCampo.Add (VariacaoMaq.Name)
    objGrid.colCampo.Add (CustoTotalMaq.Name)
    objGrid.colCampo.Add (ObservacaoMaq.Name)

    'Colunas do Grid
    iGrid_ProdutoMaq_Col = 1
    iGrid_DescricaoProdMaq_Col = 2
    iGrid_UMProdMaq_Col = 3
    iGrid_QuantidadeProdMaq_Col = 4
    iGrid_CustoUnitCalculadoMaq_Col = 5
    iGrid_CustoUnitMaq_Col = 6
    iGrid_VariacaoMaq_Col = 7
    iGrid_CustoTotalMaq_Col = 8
    iGrid_ObservacaoMaq_Col = 9

    objGrid.objGrid = GridMaquinas

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 21

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridMaquinas.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaquinas = SUCESSO

End Function

Private Sub ProdutoMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub ProdutoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub ProdutoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = ProdutoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProdMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProdMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub DescricaoProdMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub DescricaoProdMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = DescricaoProdMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProdMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProdMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub UMProdMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub UMProdMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = UMProdMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProdMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProdMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub QuantidadeProdMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub QuantidadeProdMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = QuantidadeProdMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitCalculadoMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitCalculadoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub CustoUnitCalculadoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub CustoUnitCalculadoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = CustoUnitCalculadoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitMaq_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMaqAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub CustoUnitMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub CustoUnitMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = CustoUnitMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VariacaoMaq_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMaqAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VariacaoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub VariacaoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub VariacaoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = VariacaoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoTotalMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoTotalMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub CustoTotalMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub CustoTotalMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = CustoTotalMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoMaq_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMaqAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub ObservacaoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub ObservacaoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = ObservacaoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CustoUnitMaq(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CustoUnitMaq do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dVariacao As Double
Dim dCustoTotalMaq As Double
Dim dAjusteCustoMaq As Double

On Error GoTo Erro_Saida_Celula_CustoUnitMaq

    Set objGridInt.objControle = CustoUnitMaq

    'Se o campo foi preenchido
    If Len(Trim(CustoUnitMaq.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(CustoUnitMaq.Text)
        If lErro <> SUCESSO Then gError 139337

    End If
    
    If iGridMaqAlterado = REGISTRO_ALTERADO Then
        
        'Calcula a Variação
        lErro = Calcula_Variacao(StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitCalculadoMaq_Col)), StrParaDbl(CustoUnitMaq.Text), dVariacao, iGrid_VariacaoMaq_Col)
        If lErro <> SUCESSO Then gError 139338
        
        'Altera a Variação no grid
        If dVariacao <> 0 Then
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_VariacaoMaq_Col) = Format(dVariacao, "Percent")
        Else
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_VariacaoMaq_Col) = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMaq = StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_QuantidadeProdMaq_Col)) * StrParaDbl(CustoUnitMaq.Text)
        
        'Calcula a diferença a ajustar
        dAjusteCustoMaq = dCustoTotalMaq - StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col))
        
        'Altera o Custo Total do Item no grid
        GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col) = Format(dCustoTotalMaq, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosMaquina(CustoUnitMaq.Text, iGrid_CustoUnitMaq_Col, dAjusteCustoMaq)
        If lErro <> SUCESSO Then gError 139339
        
        iGridMaqAlterado = 0
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139340

    Saida_Celula_CustoUnitMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitMaq:

    Saida_Celula_CustoUnitMaq = gErr

    Select Case gErr
        
        Case 139337 To 139340
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158518)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VariacaoMaq(objGridInt As AdmGrid) As Long
'Faz a crítica da célula VariacaoMaq do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dCustoInformado As Double
Dim dVariacao As Double
Dim dCustoTotalMaq As Double
Dim dAjusteCustoMaq As Double

On Error GoTo Erro_Saida_Celula_VariacaoMaq

    Set objGridInt.objControle = VariacaoMaq

    If iGridMaqAlterado = REGISTRO_ALTERADO Then
    
        'Ajusta e Converte o valor da tela
        If InStr(VariacaoMaq.Text, "%") > 0 Then
        
            dVariacao = StrParaDbl(Left(VariacaoMaq.Text, Len(VariacaoMaq.Text) - 1))
        
        Else
        
            dVariacao = StrParaDbl(VariacaoMaq.Text)
        
        End If
        
        'Calcula Variacao retornando o novo Custo Unitario
        lErro = Calcula_Variacao(StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitCalculadoMaq_Col)), dCustoInformado, dVariacao, iGrid_CustoUnitMaq_Col)
        If lErro <> SUCESSO Then gError 139341
        
        'Altera o Custo Unitario no grid
        If dCustoInformado <> 0 Then
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitMaq_Col) = Format(dCustoInformado, gobjFAT.sFormatoPrecoUnitario)
        Else
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitMaq_Col) = ""
        End If
        
        'se não tem Variacao ... Limpa no grid
        If dVariacao = 0 Then
           VariacaoMaq.Text = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMaq = StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_QuantidadeProdMaq_Col)) * dCustoInformado
        
        'Calcula a diferença a ajustar
        dAjusteCustoMaq = dCustoTotalMaq - StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col))
        
        'Altera o Custo Total no grid
        GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col) = Format(dCustoTotalMaq, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosMaquina(CStr(dCustoInformado), iGrid_CustoUnitMaq_Col, dAjusteCustoMaq)
        If lErro <> SUCESSO Then gError 139342
            
        iGridMaqAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139343

    Saida_Celula_VariacaoMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_VariacaoMaq:

    Saida_Celula_VariacaoMaq = gErr

    Select Case gErr
        
        Case 139341 To 139343
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158519)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ObservacaoMaq(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ObservacaoMaq do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ObservacaoMaq

    Set objGridInt.objControle = ObservacaoMaq

    If iGridMaqAlterado = REGISTRO_ALTERADO Then
    
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosMaquina(ObservacaoMaq.Text, iGrid_ObservacaoMaq_Col)
        If lErro <> SUCESSO Then gError 139344
            
        iGridMaqAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139345

    Saida_Celula_ObservacaoMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoMaq:

    Saida_Celula_ObservacaoMaq = gErr

    Select Case gErr
        
        Case 139344, 139345
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158520)

    End Select

    Exit Function

End Function

Private Sub GridMaoDeObra_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridMaoDeObra, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridMaoDeObra, iAlterado)
        End If

End Sub

Private Sub GridMaoDeObra_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub GridMaoDeObra_EnterCell()

    Call Grid_Entrada_Celula(objGridMaoDeObra, iAlterado)

End Sub

Private Sub GridMaoDeObra_LeaveCell()
    
    Call Saida_Celula(objGridMaoDeObra)

End Sub

Private Sub GridMaoDeObra_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaoDeObra, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaoDeObra, iAlterado)
    End If

End Sub

Private Sub GridMaoDeObra_RowColChange()

    Call Grid_RowColChange(objGridMaoDeObra)

End Sub

Private Sub GridMaoDeObra_Scroll()

    Call Grid_Scroll(objGridMaoDeObra)

End Sub

Private Function Inicializa_GridMaoDeObra(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Tipo")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Unit.Calc.")
    objGrid.colColuna.Add ("Unitário")
    objGrid.colColuna.Add ("Variação")
    objGrid.colColuna.Add ("Custo Total")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (TipoMO.Name)
    objGrid.colCampo.Add (DescricaoMO.Name)
    objGrid.colCampo.Add (UMMO.Name)
    objGrid.colCampo.Add (QuantidadeMO.Name)
    objGrid.colCampo.Add (CustoUnitCalculadoMO.Name)
    objGrid.colCampo.Add (CustoUnitMO.Name)
    objGrid.colCampo.Add (VariacaoMO.Name)
    objGrid.colCampo.Add (CustoTotalMO.Name)
    objGrid.colCampo.Add (ObservacaoMO.Name)

    'Colunas do Grid
    iGrid_TipoMO_Col = 1
    iGrid_DescricaoMo_Col = 2
    iGrid_UMMO_Col = 3
    iGrid_QuantidadeMO_Col = 4
    iGrid_CustoUnitCalculadoMO_Col = 5
    iGrid_CustoUnitMO_Col = 6
    iGrid_VariacaoMO_Col = 7
    iGrid_CustoTotalMO_Col = 8
    iGrid_ObservacaoMO_Col = 9

    objGrid.objGrid = GridMaoDeObra

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 21

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridMaoDeObra.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaoDeObra = SUCESSO

End Function

Private Sub TipoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub TipoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub TipoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = TipoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub DescricaoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub DescricaoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = DescricaoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub UMMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub UMMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = UMMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub QuantidadeMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub QuantidadeMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = QuantidadeMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitCalculadoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitCalculadoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub CustoUnitCalculadoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub CustoUnitCalculadoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = CustoUnitCalculadoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitMO_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMOAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub CustoUnitMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub CustoUnitMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = CustoUnitMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VariacaoMO_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMOAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VariacaoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub VariacaoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub VariacaoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = VariacaoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoTotalMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoTotalMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub CustoTotalMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub CustoTotalMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = CustoTotalMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoMO_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMOAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub ObservacaoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub ObservacaoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = ObservacaoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CustoUnitMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CustoUnitMO do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dVariacao As Double
Dim dCustoTotalMO As Double
Dim dAjusteCustoMO As Double

On Error GoTo Erro_Saida_Celula_CustoUnitMO

    Set objGridInt.objControle = CustoUnitMO

    'Se o campo foi preenchido
    If Len(Trim(CustoUnitMO.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(CustoUnitMO.Text)
        If lErro <> SUCESSO Then gError 139346

    End If
    
    If iGridMOAlterado = REGISTRO_ALTERADO Then
        
        'Calcula a Variação
        lErro = Calcula_Variacao(StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitCalculadoMO_Col)), StrParaDbl(CustoUnitMO.Text), dVariacao, iGrid_VariacaoMO_Col)
        If lErro <> SUCESSO Then gError 139347
        
        'Altera a Variação no grid
        If dVariacao <> 0 Then
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_VariacaoMO_Col) = Format(dVariacao, "Percent")
        Else
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_VariacaoMO_Col) = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMO = StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_QuantidadeMO_Col)) * StrParaDbl(CustoUnitMO.Text)
        
        'Calcula a diferença a ajustar
        dAjusteCustoMO = dCustoTotalMO - StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col))
        
        'Altera o Custo Total do Item no grid
        GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col) = Format(dCustoTotalMO, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_MaoDeObra(CustoUnitMO.Text, iGrid_CustoUnitMO_Col, dAjusteCustoMO)
        If lErro <> SUCESSO Then gError 139348
        
        iGridMOAlterado = 0
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139349

    Saida_Celula_CustoUnitMO = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitMO:

    Saida_Celula_CustoUnitMO = gErr

    Select Case gErr
        
        Case 139346 To 139349
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158521)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VariacaoMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula VariacaoMO do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dCustoInformado As Double
Dim dVariacao As Double
Dim dCustoTotalMO As Double
Dim dAjusteCustoMO As Double

On Error GoTo Erro_Saida_Celula_VariacaoMO

    Set objGridInt.objControle = VariacaoMO

    If iGridMOAlterado = REGISTRO_ALTERADO Then
        
        'Ajusta e Converte o valor da tela
        If InStr(VariacaoMO.Text, "%") > 0 Then
        
            dVariacao = StrParaDbl(Left(VariacaoMO.Text, Len(VariacaoMO.Text) - 1))
        
        Else
        
            dVariacao = StrParaDbl(VariacaoMO.Text)
        
        End If
        
        'Calcula Variacao retornando o novo Custo Unitario
        lErro = Calcula_Variacao(StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitCalculadoMO_Col)), dCustoInformado, dVariacao, iGrid_CustoUnitMO_Col)
        If lErro <> SUCESSO Then gError 139350
        
        'Altera o Custo Unitario no grid
        If dCustoInformado <> 0 Then
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitMO_Col) = Format(dCustoInformado, gobjFAT.sFormatoPrecoUnitario)
        Else
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitMO_Col) = ""
        End If
        
        'se não tem Variacao ... Limpa no grid
        If dVariacao = 0 Then
           VariacaoMO.Text = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMO = StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_QuantidadeMO_Col)) * dCustoInformado
        
        'Calcula a diferença a ajustar
        dAjusteCustoMO = dCustoTotalMO - StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col))
        
        'Altera o Custo Total no grid
        GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col) = Format(dCustoTotalMO, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_MaoDeObra(CStr(dCustoInformado), iGrid_CustoUnitMO_Col, dAjusteCustoMO)
        If lErro <> SUCESSO Then gError 139351
        
        iGridMOAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139352

    Saida_Celula_VariacaoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_VariacaoMO:

    Saida_Celula_VariacaoMO = gErr

    Select Case gErr
        
        Case 139350 To 139352
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158522)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ObservacaoMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ObservacaoMO do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ObservacaoMO

    Set objGridInt.objControle = ObservacaoMO

    If iGridMOAlterado = REGISTRO_ALTERADO Then
    
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_MaoDeObra(ObservacaoMO.Text, iGrid_ObservacaoMO_Col)
        If lErro <> SUCESSO Then gError 139353
            
        iGridMOAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139354

    Saida_Celula_ObservacaoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoMO:

    Saida_Celula_ObservacaoMO = gErr

    Select Case gErr
        
        Case 139353, 139354
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158523)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCusteioRoteiro As New ClassCusteioRoteiro

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CusteioRoteiroFabricacao"

    'Lê os dados da Tela de CusteioRoteiro
    lErro = Move_Tela_Memoria()
    If lErro <> SUCESSO Then gError 139355

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", gobjCusteioRoteiro.lCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 139355

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158524)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCusteioRoteiro As New ClassCusteioRoteiro

On Error GoTo Erro_Tela_Preenche

    objCusteioRoteiro.lCodigo = colCampoValor.Item("Codigo").vValor

    If objCusteioRoteiro.lCodigo <> 0 Then
        lErro = Traz_CusteioRoteiro_Tela(objCusteioRoteiro)
        If lErro <> SUCESSO And lErro <> 139382 Then gError 139356
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 139356

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158525)

    End Select

    Exit Function

End Function

Function Limpa_Tela_CusteioRoteiro() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CusteioRoteiro
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    'Limpa o objeto global e suas coleções
    Set gobjCusteioRoteiro = New ClassCusteioRoteiro
    
    'Limpa os demais Grids
    Call Limpa_Todos_Grids
    
    DescricaoProd.Caption = ""
    
    'Coloca a DataAtual como Data do Novo Custeio
    DataCusteio.PromptInclude = False
    DataCusteio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCusteio.PromptInclude = True
    
    Call Limpa_DadosCusteio
        
    iAlterado = 0
    iGridKitAlterado = 0
    iGridMaqAlterado = 0
    iGridMOAlterado = 0
        
    Limpa_Tela_CusteioRoteiro = SUCESSO

    Exit Function

Erro_Limpa_Tela_CusteioRoteiro:

    Limpa_Tela_CusteioRoteiro = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158526)

    End Select

    Exit Function

End Function
Function Limpa_Todos_Grids() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Todos_Grids
    
    'Limpa Tab Insumos do Kit
    LabelProdutoInsumosKit.Caption = ""
    LabelVersaoInsumosKit.Caption = ""
    
    Call Grid_Limpa(objGridKit)
    
    CustoTotalInsumosKit.Caption = ""

    'Limpa Tab Insumos da Maquina
    LabelProdutoInsumosMaq.Caption = ""
    LabelVersaoInsumosMaq.Caption = ""
    
    Call Grid_Limpa(objGridMaquinas)
    
    CustoTotalInsumosMaquina.Caption = ""

    'Limpa Tab Mão-de-Obra
    LabelProdutoMO.Caption = ""
    LabelVersaoMO.Caption = ""
    
    Call Grid_Limpa(objGridMaoDeObra)
    
    CustoTotalMaoDeObra.Caption = ""

    Limpa_Todos_Grids = SUCESSO
    
    Exit Function

Erro_Limpa_Todos_Grids:

    Limpa_Todos_Grids = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158527)
    
    End Select
    
    Exit Function

End Function

Function Trata_GridInsumosKit() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_GridInsumosKit
    
    'preenche labels de produto e versão
    LabelProdutoInsumosKit.Caption = Trim(ProdutoRaiz.Text) & SEPARADOR & DescricaoProd.Caption
    LabelVersaoInsumosKit.Caption = Trim(Versao.Text)
    
    'Limpa o Grid de Kits
    Call Grid_Limpa(objGridKit)
    
    'preenche GridInsumosKit
    lErro = Preenche_GridInsumosKit()
    If lErro <> SUCESSO Then gError 139357
    
    'Calcula o CustoTotal
    Call Calcula_CustoTotalKit
    
    Trata_GridInsumosKit = SUCESSO
    
    Exit Function

Erro_Trata_GridInsumosKit:

    Trata_GridInsumosKit = gErr
    
    Select Case gErr
    
        Case 139357
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158528)
    
    End Select
    
    Exit Function

End Function

Function Trata_GridMaoDeObra() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_GridMaoDeObra
    
    'preenche labels de produto e versão
    LabelProdutoMO.Caption = Trim(ProdutoRaiz.Text) & SEPARADOR & DescricaoProd.Caption
    LabelVersaoMO.Caption = Trim(Versao.Text)
    
    'Limpa o Grid de MaoDeObra
    Call Grid_Limpa(objGridMaoDeObra)
    
    'preenche GridMaoDeObra
    lErro = Preenche_GridMaoDeObra()
    If lErro <> SUCESSO Then gError 139358
    
    'Exibe o CustoTotal
    Call Calcula_CustoTotalMO
    
    'preenche GridMaoDeObra
        
    Trata_GridMaoDeObra = SUCESSO
    
    Exit Function

Erro_Trata_GridMaoDeObra:

    Trata_GridMaoDeObra = gErr
    
    Select Case gErr
        
        Case 139358
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158529)
    
    End Select
    
    Exit Function

End Function

Function Trata_GridInsumosMaquina() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_GridInsumosMaquina
    
    'preenche labels de produto e versão
    LabelProdutoInsumosMaq.Caption = Trim(ProdutoRaiz.Text) & SEPARADOR & DescricaoProd.Caption
    LabelVersaoInsumosMaq.Caption = Trim(Versao.Text)
    
    'Limpa o Grid de Maquinas
    Call Grid_Limpa(objGridMaquinas)
    
    'preenche GridInsumosMaquinas
    lErro = Preenche_GridInsumosMaq()
    If lErro <> SUCESSO Then gError 139359
    
    'Exibe o CustoTotal
    Call Calcula_CustoTotalMaq
        
    Trata_GridInsumosMaquina = SUCESSO
    
    Exit Function

Erro_Trata_GridInsumosMaquina:

    Trata_GridInsumosMaquina = gErr
    
    Select Case gErr
    
        Case 139359
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158530)
    
    End Select
    
    Exit Function

End Function

Function CalculaNecessidade_InsumosKit(ByVal objRoteiroFabricacao As ClassRoteirosDeFabricacao, iSeq As Integer, dCustoTotalInsumosKit As Double) As Long

Dim lErro As Long
Dim objKit As ClassKit
Dim objProdutoKit As New ClassProdutoKit
Dim objKitIntermediario As ClassKit
Dim sVersaoKit As String
Dim bProdutoNovo As Boolean
Dim objProdutoRaiz As New ClassProduto
Dim objProduto As New ClassProduto
Dim dFatorUMProduto As Double
Dim dFatorUMCusteioRot As Double
Dim dQuantidade As Double
Dim objRoteiroFabricacaoIntermediario As ClassRoteirosDeFabricacao
Dim dCustoProduto As Double
Dim objCusteioRotInsumosKit As ClassCusteioRotInsumosKit
Dim objAuxCRInsumosKit As New ClassCusteioRotInsumosKit

On Error GoTo Erro_CalculaNecessidade_InsumosKit
    
    Set objProdutoRaiz = New ClassProduto
   
    objProdutoRaiz.sCodigo = objRoteiroFabricacao.sProdutoRaiz
    
    'Lê o produto para descobrir as unidades de medidas associadas
    lErro = CF("Produto_Le", objProdutoRaiz)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 139360
   
    'Descobre o fator de conversao da UM de CusteioRotItens p/UM de estoque do produto
    lErro = CF("UM_Conversao_Trans", objProdutoRaiz.iClasseUM, objRoteiroFabricacao.sUM, objProdutoRaiz.sSiglaUMEstoque, dFatorUMCusteioRot)
    If lErro <> SUCESSO Then gError 139361
   
    Set objKit = New ClassKit

    objKit.sProdutoRaiz = objRoteiroFabricacao.sProdutoRaiz
    objKit.sVersao = objRoteiroFabricacao.sVersao
    
    'Leio os ProdutosKits que compõem este Kit
    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then gError 139362
            
    'para cada produto componente do kit ...
    For Each objProdutoKit In objKit.colComponentes

        'se não é o produto raiz ...
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
        
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objProdutoKit.sProduto
        
            'se tem composição variável ...
            If objProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
            
                'Lê o produto para descobrir as unidades de medidas associadas
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 139363
            
                'Descobre o fator de conversao da UM do Kit p/UM de estoque do produto
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorUMProduto)
                If lErro <> SUCESSO Then gError 139364
                            
                dQuantidade = (objProdutoKit.dQuantidade * dFatorUMProduto) * (objRoteiroFabricacao.dQuantidade * dFatorUMCusteioRot) / dFatorUMProduto
                                
            Else
                
                'senão... usa a quantidade fixa
                dQuantidade = objProdutoKit.dQuantidade
            
            End If
            
            Set objKitIntermediario = New ClassKit
            
            objKitIntermediario.sProdutoRaiz = objProdutoKit.sProduto
        
            'Le as Versoes Ativas e a Padrao
            lErro = CF("Kit_Le_Padrao", objKitIntermediario)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 139365

            'se encontrou - é outro Kit (Produto Intermediário)
            If lErro = SUCESSO Then
            
                'se a versão está preenchida -> usa ela
                If Len(objProdutoKit.sVersaoKitComp) <> 0 Then
                
                    sVersaoKit = objProdutoKit.sVersaoKitComp
                    
                Else
                    
                    'senão usa a Padrão
                    sVersaoKit = objKitIntermediario.sVersao
                    
                End If
                
                'Cria um novo objRoteiroFabricacao
                Set objRoteiroFabricacaoIntermediario = New ClassRoteirosDeFabricacao
                
                objRoteiroFabricacaoIntermediario.sProdutoRaiz = objProdutoKit.sProduto
                objRoteiroFabricacaoIntermediario.sVersao = sVersaoKit
                objRoteiroFabricacaoIntermediario.sUM = objProdutoKit.sUnidadeMed
                objRoteiroFabricacaoIntermediario.dQuantidade = dQuantidade
                
                'e chama esta função recursivamente ...
                lErro = CalculaNecessidade_InsumosKit(objRoteiroFabricacaoIntermediario, iSeq, dCustoTotalInsumosKit)
                If lErro <> SUCESSO Then gError 139366
                
            Else
                
                'senão... vamos incluir na coleção...
                bProdutoNovo = True
                
                'Verifica se já há algum produto do kit na coleção
                For Each objAuxCRInsumosKit In gobjCusteioRoteiro.colCusteioRotInsumosKit
            
                    'se encontrou ...
                    If objAuxCRInsumosKit.sProduto = objProdutoKit.sProduto Then
                        
                        'subtrai o valor anterior do Total do Custo no acumulador
                        dCustoTotalInsumosKit = dCustoTotalInsumosKit - (objAuxCRInsumosKit.dQuantidade * objAuxCRInsumosKit.dCustoUnitarioInformado)
                        
                        'acumula a quantidade
                        objAuxCRInsumosKit.dQuantidade = objAuxCRInsumosKit.dQuantidade + dQuantidade
                        
                        'lança o novo valor Total do Custo no acumulador
                        dCustoTotalInsumosKit = dCustoTotalInsumosKit + (objAuxCRInsumosKit.dQuantidade * objAuxCRInsumosKit.dCustoUnitarioInformado)
                        
                        'avisa que acumulou ...
                        bProdutoNovo = False
                        Exit For
                        
                    End If
                    
                Next
                
                'se não tem o produto ...
                If bProdutoNovo Then
                    
                    'Descobre qual seu custo
                    lErro = CF("Produto_Le_CustoProduto", objProduto, dCustoProduto)
                    If lErro <> SUCESSO Then gError 139367
                                                
                    'incrementa o sequencial
                    iSeq = iSeq + 1
                    
                    'reCria o obj
                    Set objCusteioRotInsumosKit = New ClassCusteioRotInsumosKit
                    
                    objCusteioRotInsumosKit.iSeq = iSeq
                    objCusteioRotInsumosKit.sProduto = objProdutoKit.sProduto
                    objCusteioRotInsumosKit.sUMedida = objProdutoKit.sUnidadeMed
                    objCusteioRotInsumosKit.dQuantidade = dQuantidade
                    objCusteioRotInsumosKit.dCustoUnitarioCalculado = dCustoProduto
                    objCusteioRotInsumosKit.dCustoUnitarioInformado = dCustoProduto
                    
                    'lança o valor Total do Custo no acumulador
                    dCustoTotalInsumosKit = dCustoTotalInsumosKit + (dQuantidade * dCustoProduto)
                    
                    'e inclui na coleção.
                    gobjCusteioRoteiro.colCusteioRotInsumosKit.Add objCusteioRotInsumosKit
                    
                End If
            
            End If
            
        End If
    
    Next
    
    CalculaNecessidade_InsumosKit = SUCESSO
    
    Exit Function

Erro_CalculaNecessidade_InsumosKit:

    CalculaNecessidade_InsumosKit = gErr
    
    Select Case gErr
    
        Case 139360 To 139367
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158531)
    
    End Select
    
    Exit Function

End Function

Function CalculaNecessidade_InsumosMaquina(ByVal objRoteiroOriginal As ClassRoteirosDeFabricacao, iSeq As Integer, dCustoTotalInsumosMaq As Double) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim dFator As Double
Dim dFatorQuantidade As Double
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim objOperacoes As New ClassOperacoes
Dim objPO As ClassPlanoOperacional
Dim objOPOperacoes As ClassOrdemProducaoOperacoes
Dim objPOMaquinas As New ClassPOMaquinas
Dim objMaquinas As ClassMaquinas
Dim objMaquinasInsumos As New ClassMaquinasInsumos
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim objRoteiroFabricacaoIntermediario As ClassRoteirosDeFabricacao
Dim objPMPItens As ClassPMPItens

On Error GoTo Erro_CalculaNecessidade_InsumosMaquina

    Set objRoteirosDeFabricacao = New ClassRoteirosDeFabricacao
    
    objRoteirosDeFabricacao.sProdutoRaiz = objRoteiroOriginal.sProdutoRaiz
    objRoteirosDeFabricacao.sVersao = objRoteiroOriginal.sVersao
    
    lErro = CF("RoteirosDeFabricacao_Le", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 139368
    
    If lErro = SUCESSO Then
    
        Set objProduto = New ClassProduto
       
        objProduto.sCodigo = objRoteiroOriginal.sProdutoRaiz
    
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 139369
       
        'Descobre o fator de conversao da UM de objRoteiroOriginal p/UM de RoteirosDeFabricacao
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objRoteiroOriginal.sUM, objRoteirosDeFabricacao.sUM, dFator)
        If lErro <> SUCESSO Then gError 139370

        dFatorQuantidade = (objRoteiroOriginal.dQuantidade * dFator) / objRoteirosDeFabricacao.dQuantidade
        
        Set objPO = New ClassPlanoOperacional
        
        objPO.sProduto = objRoteiroOriginal.sProdutoRaiz
        objPO.dQuantidade = objRoteiroOriginal.dQuantidade
        objPO.sUM = objRoteiroOriginal.sUM
        
        For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes
        
            Set objOPOperacoes = New ClassOrdemProducaoOperacoes
            
            objOPOperacoes.lNumIntDocCT = objOperacoes.lNumIntDocCT
            objOPOperacoes.lNumIntDocCompet = objOperacoes.lNumIntDocCompet
            objOPOperacoes.iConsideraCarga = MARCADO
            
            Set objOPOperacoes.objOperacoesTempo = objOperacoes.objOperacoesTempo
            
            Set objPMPItens = New ClassPMPItens
                        
            lErro = CF("PlanoOperacional_Calcula_Tempos", objPMPItens, objPO, objOPOperacoes, MRP_ACERTA_POR_DATA_FIM)
            If lErro <> SUCESSO Then gError 139371
            
            For Each objPOMaquinas In objPO.colAlocacaoMaquinas
            
                Set objMaquinas = New ClassMaquinas
            
                objMaquinas.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
                 
                lErro = CF("Maquinas_Le_Itens", objMaquinas)
                If lErro <> SUCESSO Then gError 139372
                
                For Each objMaquinasInsumos In objMaquinas.colProdutos
                
                    lErro = IncluiNecessidade_InsumosMaquina(objMaquinasInsumos, objPOMaquinas, iSeq, dCustoTotalInsumosMaq)
                    If lErro <> SUCESSO Then gError 139373
                
                Next
            
            Next
            
            For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos
            
                Set objRoteiroFabricacaoIntermediario = New ClassRoteirosDeFabricacao
                
                objRoteiroFabricacaoIntermediario.sProdutoRaiz = objOperacaoInsumos.sProduto
                objRoteiroFabricacaoIntermediario.sVersao = objOperacaoInsumos.sVersaoKitComp
                objRoteiroFabricacaoIntermediario.sUM = objOperacaoInsumos.sUMProduto
                objRoteiroFabricacaoIntermediario.dQuantidade = objOperacaoInsumos.dQuantidade * dFatorQuantidade
                
                lErro = CalculaNecessidade_InsumosMaquina(objRoteiroFabricacaoIntermediario, iSeq, dCustoTotalInsumosMaq)
                If lErro <> SUCESSO Then gError 139374
                
            Next
        
        Next
    
    End If
    
    CalculaNecessidade_InsumosMaquina = SUCESSO
    
    Exit Function

Erro_CalculaNecessidade_InsumosMaquina:

    CalculaNecessidade_InsumosMaquina = gErr
    
    Select Case gErr
    
        Case 139368 To 139374
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158532)
    
    End Select
    
    Exit Function

End Function

Function CalculaNecessidade_MaoDeObra(ByVal objRoteiroOriginal As ClassRoteirosDeFabricacao, iSeq As Integer, dCustoTotalMaoDeObra As Double) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim dFator As Double
Dim dFatorQuantidade As Double
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim objOperacoes As New ClassOperacoes
Dim objPO As ClassPlanoOperacional
Dim objOPOperacoes As ClassOrdemProducaoOperacoes
Dim objPOMaquinas As New ClassPOMaquinas
Dim objMaquinas As ClassMaquinas
Dim objMaquinaOperadores As New ClassMaquinaOperadores
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim objRoteiroFabricacaoIntermediario As ClassRoteirosDeFabricacao
Dim objPMPItens As ClassPMPItens

On Error GoTo Erro_CalculaNecessidade_MaoDeObra

    Set objRoteirosDeFabricacao = New ClassRoteirosDeFabricacao

    objRoteirosDeFabricacao.sProdutoRaiz = objRoteiroOriginal.sProdutoRaiz
    objRoteirosDeFabricacao.sVersao = objRoteiroOriginal.sVersao

    lErro = CF("RoteirosDeFabricacao_Le", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 139375

    If lErro = SUCESSO Then
    
        Set objProduto = New ClassProduto
       
        objProduto.sCodigo = objRoteiroOriginal.sProdutoRaiz
    
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 139376
       
        'Descobre o fator de conversao da UM de objRoteiroOriginal p/UM de RoteiroDeFabricacao
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objRoteiroOriginal.sUM, objRoteirosDeFabricacao.sUM, dFator)
        If lErro <> SUCESSO Then gError 134976
    
        dFatorQuantidade = (objRoteiroOriginal.dQuantidade * dFator) / objRoteirosDeFabricacao.dQuantidade
        
        Set objPO = New ClassPlanoOperacional

        objPO.sProduto = objRoteiroOriginal.sProdutoRaiz
        objPO.dQuantidade = objRoteiroOriginal.dQuantidade
        objPO.sUM = objRoteiroOriginal.sUM

        For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes

            Set objOPOperacoes = New ClassOrdemProducaoOperacoes

            objOPOperacoes.lNumIntDocCT = objOperacoes.lNumIntDocCT
            objOPOperacoes.lNumIntDocCompet = objOperacoes.lNumIntDocCompet
            objOPOperacoes.iConsideraCarga = MARCADO

            Set objOPOperacoes.objOperacoesTempo = objOperacoes.objOperacoesTempo

            Set objPMPItens = New ClassPMPItens

            lErro = CF("PlanoOperacional_Calcula_Tempos", objPMPItens, objPO, objOPOperacoes, MRP_ACERTA_POR_DATA_FIM)
            If lErro <> SUCESSO Then gError 139377

            For Each objPOMaquinas In objPO.colAlocacaoMaquinas

                Set objMaquinas = New ClassMaquinas

                objMaquinas.lNumIntDoc = objPOMaquinas.lNumIntDocMaq

                lErro = CF("Maquinas_Le_Itens", objMaquinas)
                If lErro <> SUCESSO Then gError 139378

                For Each objMaquinaOperadores In objMaquinas.colTipoOperadores

                    lErro = IncluiNecessidade_MaoDeObra(objMaquinaOperadores, objPOMaquinas, iSeq, dCustoTotalMaoDeObra)
                    If lErro <> SUCESSO Then gError 139379

                Next

            Next

            For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos

                Set objRoteiroFabricacaoIntermediario = New ClassRoteirosDeFabricacao
                
                objRoteiroFabricacaoIntermediario.sProdutoRaiz = objOperacaoInsumos.sProduto
                objRoteiroFabricacaoIntermediario.sVersao = objOperacaoInsumos.sVersaoKitComp
                objRoteiroFabricacaoIntermediario.sUM = objOperacaoInsumos.sUMProduto
                objRoteiroFabricacaoIntermediario.dQuantidade = objOperacaoInsumos.dQuantidade * dFatorQuantidade

                lErro = CalculaNecessidade_MaoDeObra(objRoteiroFabricacaoIntermediario, iSeq, dCustoTotalMaoDeObra)
                If lErro <> SUCESSO Then gError 139380

            Next

        Next

    End If

    CalculaNecessidade_MaoDeObra = SUCESSO

    Exit Function

Erro_CalculaNecessidade_MaoDeObra:

    CalculaNecessidade_MaoDeObra = gErr

    Select Case gErr

        Case 139375 To 139380
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158533)

    End Select

    Exit Function

End Function

Function Traz_CusteioRoteiro_Tela(objCusteioRoteiro As ClassCusteioRoteiro) As Long

Dim lErro As Long
Dim sProdutoMascarado As String
Dim dCustoTotal As Double

On Error GoTo Erro_Traz_CusteioRoteiro_Tela

    'Verifica se o CusteioRoteiro existe, lendo no BD
    lErro = CF("CusteioRoteiro_Le", objCusteioRoteiro)
    If lErro <> SUCESSO And lErro <> 137942 Then gError 139381
    
    'se não existe ...
    If lErro <> SUCESSO Then gError 139382
    
    Call Limpa_Tela_CusteioRoteiro

    Codigo.Text = objCusteioRoteiro.lCodigo
    NomeReduzido.Text = objCusteioRoteiro.sNomeReduzido
    Descricao.Text = objCusteioRoteiro.sDescricao

    'Exibe Data do Custeio na Tela
    If objCusteioRoteiro.dtDataCusteio <> DATA_NULA Then
        DataCusteio.PromptInclude = False
        DataCusteio.Text = Format(objCusteioRoteiro.dtDataCusteio, "dd/mm/yy")
        DataCusteio.PromptInclude = True
    End If

    'Exibe Data de Validade na Tela
    If objCusteioRoteiro.dtDataValidade <> DATA_NULA Then
        DataValidade.PromptInclude = False
        DataValidade.Text = Format(objCusteioRoteiro.dtDataCusteio, "dd/mm/yy")
        DataValidade.PromptInclude = True
    End If
    
    lErro = Mascara_RetornaProdutoTela(objCusteioRoteiro.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 139383
    
    ProdutoRaiz.Text = sProdutoMascarado
    
    Call ProdutoRaiz_Validate(bSGECancelDummy)
    
    Versao.Text = objCusteioRoteiro.sVersao
    
    Observacao.Text = objCusteioRoteiro.sObservacao

    Set gobjCusteioRoteiro = objCusteioRoteiro
    
    LabelDescQuantidade.Caption = Formata_Estoque(gobjCusteioRoteiro.dQuantidade)
    LabelDescUM.Caption = gobjCusteioRoteiro.sUMedida
    
    LabelDescCustoMP.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosKit, "Standard")
    LabelDescInsMaq.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosMaq, "Standard")
    LabelDescCustoMO.Caption = Format(gobjCusteioRoteiro.dCustoTotalMaoDeObra, "Standard")
    
    dCustoTotal = gobjCusteioRoteiro.dCustoTotalInsumosKit + gobjCusteioRoteiro.dCustoTotalInsumosMaq + gobjCusteioRoteiro.dCustoTotalMaoDeObra
    LabelDescCustoTotal.Caption = Format(dCustoTotal, "Standard")
        
    iAlterado = 0
    iGridKitAlterado = 0
    iGridMaqAlterado = 0
    iGridMOAlterado = 0
            
    Traz_CusteioRoteiro_Tela = SUCESSO

    Exit Function

Erro_Traz_CusteioRoteiro_Tela:

    Traz_CusteioRoteiro_Tela = gErr

    Select Case gErr

        Case 139381, 139383
            'erros tratados nas rotinas chamadas
            
        Case 139382
            'erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158534)

    End Select

    Exit Function

End Function

Function Trata_Custeio(ByVal iTabSelecionada As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Custeio

    'se está preenchido o produto e versão, seleciona a Tab
    If Len(gobjCusteioRoteiro.sProduto) <> 0 And Len(gobjCusteioRoteiro.sVersao) <> 0 Then
                
        'Trata o Grid da Tab Selecionada
        Select Case iTabSelecionada
        
            Case Is = TAB_InsumosKit
                
                Call Trata_GridInsumosKit
            
            Case Is = TAB_InsumosMaq
                
                Call Trata_GridInsumosMaquina
            
            Case Is = TAB_MaoDeObra
            
                Call Trata_GridMaoDeObra
                
        End Select
    
    Else
            
        'senão -> limpa tudo
        Call Limpa_Todos_Grids
    
    End If
                
    Trata_Custeio = SUCESSO
    
    Exit Function
            
Erro_Trata_Custeio:

    Trata_Custeio = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158535)

    End Select

    Exit Function

End Function

Function Preenche_GridInsumosKit() As Long

Dim lErro As Long
Dim objCusteioRotInsumosKit As New ClassCusteioRotInsumosKit
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim dVariacao As Double

On Error GoTo Erro_Preenche_GridInsumosKit

    For Each objCusteioRotInsumosKit In gobjCusteioRoteiro.colCusteioRotInsumosKit
    
        objProduto.sCodigo = objCusteioRotInsumosKit.sProduto
        
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 139384
        
        sProdutoFormatado = objProduto.sCodigo
        
        'Mascara o produto para exibição no Grid
        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 139385
       
        'inicializa linha do grid
        iLinha = objCusteioRotInsumosKit.iSeq
        
        'inclui linha no grid
        GridKit.TextMatrix(iLinha, iGrid_ProdutoKit_Col) = sProdutoMascarado
        GridKit.TextMatrix(iLinha, iGrid_DescricaoProdKit_Col) = objProduto.sDescricao
        GridKit.TextMatrix(iLinha, iGrid_UMProdKit_Col) = objCusteioRotInsumosKit.sUMedida
        GridKit.TextMatrix(iLinha, iGrid_QuantidadeProdKit_Col) = Formata_Estoque(objCusteioRotInsumosKit.dQuantidade)
        
        'se tem custo unitário calculado ...
        If objCusteioRotInsumosKit.dCustoUnitarioCalculado > 0 Then
            
            GridKit.TextMatrix(iLinha, iGrid_CustoUnitCalculadoKit_Col) = Format(objCusteioRotInsumosKit.dCustoUnitarioCalculado, gobjFAT.sFormatoPrecoUnitario)
        
        End If
        
        'se tem custo unitario informado
        If objCusteioRotInsumosKit.dCustoUnitarioInformado > 0 Then
            
            GridKit.TextMatrix(iLinha, iGrid_CustoUnitKit_Col) = Format(objCusteioRotInsumosKit.dCustoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
            
            'inicializa variável
            dVariacao = 0
                    
            'calcula a variação e põe no grid
            Call Calcula_Variacao(objCusteioRotInsumosKit.dCustoUnitarioCalculado, objCusteioRotInsumosKit.dCustoUnitarioInformado, dVariacao, iGrid_VariacaoKit_Col)
            If dVariacao <> 0 Then
                GridKit.TextMatrix(iLinha, iGrid_VariacaoKit_Col) = Format(dVariacao, "Percent")
            End If
            
            'exibe o total do custo em função da quantidade
            GridKit.TextMatrix(iLinha, iGrid_CustoTotalKit_Col) = Format(objCusteioRotInsumosKit.dQuantidade * objCusteioRotInsumosKit.dCustoUnitarioInformado, "Standard")
        
        End If
        
        'exibe a observacao
        GridKit.TextMatrix(iLinha, iGrid_ObservacaoKit_Col) = objCusteioRotInsumosKit.sObservacao
        
    Next
    
    'fixa as linhas do grid
    objGridKit.iLinhasExistentes = gobjCusteioRoteiro.colCusteioRotInsumosKit.Count
    
    Preenche_GridInsumosKit = SUCESSO
    
    Exit Function
            
Erro_Preenche_GridInsumosKit:

    Preenche_GridInsumosKit = gErr
    
    Select Case gErr
    
        Case 139384, 139385
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158536)

    End Select

    Exit Function

End Function

Function Preenche_GridInsumosMaq() As Long

Dim lErro As Long
Dim objCusteioRotInsumosMaq As New ClassCusteioRotInsumosMaq
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim dVariacao As Double

On Error GoTo Erro_Preenche_GridInsumosMaq

    For Each objCusteioRotInsumosMaq In gobjCusteioRoteiro.colCusteioRotInsumosMaq
    
        objProduto.sCodigo = objCusteioRotInsumosMaq.sProduto
        
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 139386
        
        sProdutoFormatado = objProduto.sCodigo
        
        'Mascara o produto para exibição no Grid
        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 139387
       
        'inicializa linha do grid
        iLinha = objCusteioRotInsumosMaq.iSeq
        
        'inclui linha no grid
        GridMaquinas.TextMatrix(iLinha, iGrid_ProdutoMaq_Col) = sProdutoMascarado
        GridMaquinas.TextMatrix(iLinha, iGrid_DescricaoProdMaq_Col) = objProduto.sDescricao
        GridMaquinas.TextMatrix(iLinha, iGrid_UMProdMaq_Col) = objCusteioRotInsumosMaq.sUMedida
        GridMaquinas.TextMatrix(iLinha, iGrid_QuantidadeProdMaq_Col) = Formata_Estoque(objCusteioRotInsumosMaq.dQuantidade)
        
        'se tem custo unitário calculado ...
        If objCusteioRotInsumosMaq.dCustoUnitarioCalculado > 0 Then
            
            GridMaquinas.TextMatrix(iLinha, iGrid_CustoUnitCalculadoMaq_Col) = Format(objCusteioRotInsumosMaq.dCustoUnitarioCalculado, gobjFAT.sFormatoPrecoUnitario)
        
        End If
        
        'se tem custo unitario informado
        If objCusteioRotInsumosMaq.dCustoUnitarioInformado > 0 Then
            
            GridMaquinas.TextMatrix(iLinha, iGrid_CustoUnitMaq_Col) = Format(objCusteioRotInsumosMaq.dCustoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
            
            'inicializa variável
            dVariacao = 0
                    
            'calcula a variação e põe no grid
            Call Calcula_Variacao(objCusteioRotInsumosMaq.dCustoUnitarioCalculado, objCusteioRotInsumosMaq.dCustoUnitarioInformado, dVariacao, iGrid_VariacaoMaq_Col)
            If dVariacao <> 0 Then
                GridMaquinas.TextMatrix(iLinha, iGrid_VariacaoMaq_Col) = Format(dVariacao, "Percent")
            End If
            
            'exibe o total do custo em função da quantidade
            GridMaquinas.TextMatrix(iLinha, iGrid_CustoTotalMaq_Col) = Format(objCusteioRotInsumosMaq.dQuantidade * objCusteioRotInsumosMaq.dCustoUnitarioInformado, "Standard")
        
        End If
        
        'exibe a observacao
        GridMaquinas.TextMatrix(iLinha, iGrid_ObservacaoMaq_Col) = objCusteioRotInsumosMaq.sObservacao
                
    Next
    
    'fixa as linhas do grid
    objGridMaquinas.iLinhasExistentes = gobjCusteioRoteiro.colCusteioRotInsumosMaq.Count
    
    Preenche_GridInsumosMaq = SUCESSO
    
    Exit Function
            
Erro_Preenche_GridInsumosMaq:

    Preenche_GridInsumosMaq = gErr
    
    Select Case gErr
    
        Case 139386, 139387
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158537)

    End Select

    Exit Function

End Function

Function Preenche_GridMaoDeObra() As Long

Dim lErro As Long
Dim objCusteioRotMaoDeObra As New ClassCusteioRotMaoDeObra
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer
Dim dVariacao As Double

On Error GoTo Erro_Preenche_GridMaoDeObra

    For Each objCusteioRotMaoDeObra In gobjCusteioRoteiro.colCusteioRotMaoDeObra
    
        Set objTiposDeMaodeObra = New ClassTiposDeMaodeObra
        
        objTiposDeMaodeObra.iCodigo = objCusteioRotMaoDeObra.iCodMO
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 135004 Then gError 139388
       
        'inicializa linha do grid
        iLinha = objCusteioRotMaoDeObra.iSeq
        
        'inclui linha no grid
        GridMaoDeObra.TextMatrix(iLinha, iGrid_TipoMO_Col) = objCusteioRotMaoDeObra.iCodMO
        GridMaoDeObra.TextMatrix(iLinha, iGrid_DescricaoMo_Col) = objTiposDeMaodeObra.sDescricao
        GridMaoDeObra.TextMatrix(iLinha, iGrid_UMMO_Col) = objCusteioRotMaoDeObra.sUMedida
        GridMaoDeObra.TextMatrix(iLinha, iGrid_QuantidadeMO_Col) = Formata_Estoque(objCusteioRotMaoDeObra.dQuantidade)
        
        'se tem custo unitário calculado ...
        If objCusteioRotMaoDeObra.dCustoUnitarioCalculado > 0 Then
            
            GridMaoDeObra.TextMatrix(iLinha, iGrid_CustoUnitCalculadoMO_Col) = Format(objCusteioRotMaoDeObra.dCustoUnitarioCalculado, gobjFAT.sFormatoPrecoUnitario)
        
        End If
        
        'se tem custo unitario informado
        If objCusteioRotMaoDeObra.dCustoUnitarioInformado > 0 Then
            
            GridMaoDeObra.TextMatrix(iLinha, iGrid_CustoUnitMO_Col) = Format(objCusteioRotMaoDeObra.dCustoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
            
            'inicializa variável
            dVariacao = 0
                    
            'calcula a variação e põe no grid
            Call Calcula_Variacao(objCusteioRotMaoDeObra.dCustoUnitarioCalculado, objCusteioRotMaoDeObra.dCustoUnitarioInformado, dVariacao, iGrid_VariacaoMO_Col)
            If dVariacao <> 0 Then
                GridMaoDeObra.TextMatrix(iLinha, iGrid_VariacaoMO_Col) = Format(dVariacao, "Percent")
            End If
            
            'exibe o total do custo em função da quantidade
            GridMaoDeObra.TextMatrix(iLinha, iGrid_CustoTotalMO_Col) = Format(objCusteioRotMaoDeObra.dQuantidade * objCusteioRotMaoDeObra.dCustoUnitarioInformado, "Standard")
        
        End If
        
        'exibe a observacao
        GridMaoDeObra.TextMatrix(iLinha, iGrid_ObservacaoMO_Col) = objCusteioRotMaoDeObra.sObservacao
                
    Next
    
    'fixa as linhas do grid
    objGridMaoDeObra.iLinhasExistentes = gobjCusteioRoteiro.colCusteioRotMaoDeObra.Count
    
    Preenche_GridMaoDeObra = SUCESSO
    
    Exit Function
            
Erro_Preenche_GridMaoDeObra:

    Preenche_GridMaoDeObra = gErr
    
    Select Case gErr
    
        Case 139388
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158538)

    End Select

    Exit Function

End Function

Function CalculaNecessidadesProducao(ByVal objRoteiroFabricacao As ClassRoteirosDeFabricacao) As Long
    
Dim lErro As Long
Dim dCustoTotalInsumosKit As Double
Dim dCustoTotalInsumosMaq As Double
Dim dCustoTotalMaoDeObra As Double
Dim objProduto As ClassProduto
Dim dCustoTotalItem As Double
Dim dPrecoTotalItem As Double

On Error GoTo Erro_CalculaNecessidadesProducao

    'calcula as Necessidade de Produção de Insumos do Kit
    lErro = CalculaNecessidade_InsumosKit(objRoteiroFabricacao, 0, dCustoTotalInsumosKit)
    If lErro <> SUCESSO Then gError 139389
    
    gobjCusteioRoteiro.dCustoTotalInsumosKit = dCustoTotalInsumosKit
    
    'calcula as Necessidade de Produção de Insumos da Maquina
    lErro = CalculaNecessidade_InsumosMaquina(objRoteiroFabricacao, 0, dCustoTotalInsumosMaq)
    If lErro <> SUCESSO Then gError 139390
    
    gobjCusteioRoteiro.dCustoTotalInsumosMaq = dCustoTotalInsumosMaq
    
    'calcula as Necessidade de Produção de Mão-de-Obra da Maquina
    lErro = CalculaNecessidade_MaoDeObra(objRoteiroFabricacao, 0, dCustoTotalMaoDeObra)
    If lErro <> SUCESSO Then gError 139391
    
    gobjCusteioRoteiro.dCustoTotalMaoDeObra = dCustoTotalMaoDeObra
    
    'Calcula o Custo Total do Item e o Preço Total do Item
    dCustoTotalItem = dCustoTotalInsumosKit + dCustoTotalInsumosMaq + dCustoTotalMaoDeObra
    
    CalculaNecessidadesProducao = SUCESSO
    
    Exit Function
    
Erro_CalculaNecessidadesProducao:

    CalculaNecessidadesProducao = gErr

    Select Case gErr
    
        Case 139389 To 139392
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158539)

    End Select

    Exit Function

End Function

Function Calcula_Variacao(ByVal dCustoUnitarioCalculado As Double, dCustoUnitarioInformado As Double, dVariacao As Double, iColunaGrid As Integer) As Long

Dim lErro As Long
Dim dDiferenca As Double

On Error GoTo Erro_Calcula_Variacao

    Select Case iColunaGrid
    
        'se quero calcular o percentual
        Case Is = iGrid_VariacaoKit_Col, _
                  iGrid_VariacaoMaq_Col, _
                  iGrid_VariacaoMO_Col
        
            'se o calculado não for zero
            If dCustoUnitarioCalculado <> 0 Then
                
                'se o informado não for zero
                If dCustoUnitarioInformado <> 0 Then
                
                    'acho a diferença
                    dDiferenca = dCustoUnitarioInformado - dCustoUnitarioCalculado
                    'e calculo a variacao
                    dVariacao = (dDiferenca / dCustoUnitarioCalculado)
                
                Else
                
                    'a variacao é negativa
                    dVariacao = -1

                End If
                
            Else
                
                'se o informado não for zero
                If dCustoUnitarioInformado <> 0 Then
                
                    'a variacao é total
                    dVariacao = 1
                    
                Else
                
                    'nao tem variacao
                    dVariacao = 0

                End If
                
            End If
        
        'se quero calcular o custo unitario
        Case Is = iGrid_CustoUnitKit_Col, _
                  iGrid_CustoUnitMaq_Col, _
                  iGrid_CustoUnitMO_Col
            
            'independentemente de qual membro da equação for zerado ou negativo ...
            'calcula o informado = (calculado + (calculado * variacao))
            dCustoUnitarioInformado = dCustoUnitarioCalculado + ((dCustoUnitarioCalculado * dVariacao) / 100)
            
    End Select
    
    Calcula_Variacao = SUCESSO
    
    Exit Function
    
Erro_Calcula_Variacao:

    Calcula_Variacao = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158540)
        
    End Select
    
    Exit Function

End Function

Function AjustaCustoNecessidade_InsumosKit(ByVal sConteudoAAjustar As String, ByVal iColunaAAjustar As Integer, Optional ByVal dAjusteCustoKit As Double) As Long

Dim lErro As Long
Dim objCusteioRotInsumosKit As New ClassCusteioRotInsumosKit

On Error GoTo Erro_AjustaCustoNecessidade_InsumosKit

    'Percorre a coleção no gobjCusteioRoteiro localizado
    For Each objCusteioRotInsumosKit In gobjCusteioRoteiro.colCusteioRotInsumosKit
    
        'Verifica se é o item que está alterando ...
        If objCusteioRotInsumosKit.iSeq = GridKit.Row Then
            
            'encontrou... altera o valor da coluna passada
            Select Case iColunaAAjustar
            
                Case Is = iGrid_CustoUnitKit_Col
                
                    objCusteioRotInsumosKit.dCustoUnitarioInformado = StrParaDbl(sConteudoAAjustar)
                
                Case Is = iGrid_ObservacaoKit_Col
                    
                    objCusteioRotInsumosKit.sObservacao = sConteudoAAjustar
                
            End Select
            
            'e encerra a busca
            Exit For
            
        End If
    
    Next
    
    'se a coluna a ajustar é a do Custo
    If iColunaAAjustar = iGrid_CustoUnitKit_Col Then
    
        'Ajusta o valor no obj
        gobjCusteioRoteiro.dCustoTotalInsumosKit = gobjCusteioRoteiro.dCustoTotalInsumosKit + dAjusteCustoKit
        
        'Ajusta o valor no Total do Tab
        CustoTotalInsumosKit.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosKit, "Standard")
        
        'Ajusta o valor no Grid de Itens do CusteioRot
        If gobjCusteioRoteiro.dCustoTotalInsumosKit > 0 Then
            LabelDescCustoMP.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosKit, "Standard")
        Else
            LabelDescCustoMP.Caption = ""
        End If
        
        'Recalcula o Custo Total do Item e o Preço Total do Item
        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 139393

    End If

    AjustaCustoNecessidade_InsumosKit = SUCESSO
    
    Exit Function
    
Erro_AjustaCustoNecessidade_InsumosKit:

    AjustaCustoNecessidade_InsumosKit = gErr
    
    Select Case gErr
    
        Case 139393
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158541)
            
    End Select
    
    Exit Function

End Function

Function AjustaCustoNecessidade_InsumosMaquina(ByVal sConteudoAAjustar As String, ByVal iColunaAAjustar As Integer, Optional ByVal dAjusteCustoMaq As Double) As Long

Dim lErro As Long
Dim objCusteioRotInsumosMaq As New ClassCusteioRotInsumosMaq

On Error GoTo Erro_AjustaCustoNecessidade_InsumosMaquina

    'Percorre a coleção no gobjCusteioRoteiro localizado
    For Each objCusteioRotInsumosMaq In gobjCusteioRoteiro.colCusteioRotInsumosMaq
    
        'Verifica se é o item que está alterando ...
        If objCusteioRotInsumosMaq.iSeq = GridMaquinas.Row Then
            
            'encontrou... altera o valor da coluna passada
            Select Case iColunaAAjustar
            
                Case Is = iGrid_CustoUnitMaq_Col
                
                    objCusteioRotInsumosMaq.dCustoUnitarioInformado = StrParaDbl(sConteudoAAjustar)
                
                Case Is = iGrid_ObservacaoMaq_Col
                    
                    objCusteioRotInsumosMaq.sObservacao = sConteudoAAjustar
                
            End Select
            
            'e encerra a busca
            Exit For
            
        End If
    
    Next
    
    'se a coluna a ajustar é a do Custo
    If iColunaAAjustar = iGrid_CustoUnitMaq_Col Then
    
        'Ajusta o valor no obj
        gobjCusteioRoteiro.dCustoTotalInsumosMaq = gobjCusteioRoteiro.dCustoTotalInsumosMaq + dAjusteCustoMaq
        
        'Ajusta o valor no Total do Tab
        CustoTotalInsumosMaquina.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosMaq, "Standard")
        
        'Ajusta o valor no Grid de Itens do CusteioRot
        If gobjCusteioRoteiro.dCustoTotalInsumosMaq > 0 Then
            LabelDescInsMaq.Caption = Format(gobjCusteioRoteiro.dCustoTotalInsumosMaq, "Standard")
        Else
            LabelDescInsMaq.Caption = ""
        End If
        
        'Recalcula o Custo Total do Item e o Preço Total do Item
        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 139394

    End If

    AjustaCustoNecessidade_InsumosMaquina = SUCESSO
    
    Exit Function
    
Erro_AjustaCustoNecessidade_InsumosMaquina:

    AjustaCustoNecessidade_InsumosMaquina = gErr
    
    Select Case gErr
        
        Case 139394
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158542)
            
    End Select
    
    Exit Function

End Function

Function AjustaCustoNecessidade_MaoDeObra(ByVal sConteudoAAjustar As String, ByVal iColunaAAjustar As Integer, Optional ByVal dAjusteCustoMO As Double) As Long

Dim lErro As Long
Dim objCusteioRotMaoDeObra As New ClassCusteioRotMaoDeObra

On Error GoTo Erro_AjustaCustoNecessidade_MaoDeObra

    'Percorre a coleção no gobjCusteioRoteiro localizado
    For Each objCusteioRotMaoDeObra In gobjCusteioRoteiro.colCusteioRotMaoDeObra
    
        'Verifica se é o item que está alterando ...
        If objCusteioRotMaoDeObra.iSeq = GridMaoDeObra.Row Then
            
            'encontrou... altera o valor da coluna passada
            Select Case iColunaAAjustar
            
                Case Is = iGrid_CustoUnitMO_Col
                
                    objCusteioRotMaoDeObra.dCustoUnitarioInformado = StrParaDbl(sConteudoAAjustar)
                
                Case Is = iGrid_ObservacaoMO_Col
                    
                    objCusteioRotMaoDeObra.sObservacao = sConteudoAAjustar
                
            End Select
            
            'e encerra a busca
            Exit For
            
        End If
    
    Next
    
    'se a coluna a ajustar é a do Custo
    If iColunaAAjustar = iGrid_CustoUnitMO_Col Then
    
        'Ajusta o valor no obj
        gobjCusteioRoteiro.dCustoTotalMaoDeObra = gobjCusteioRoteiro.dCustoTotalMaoDeObra + dAjusteCustoMO
        
        'Ajusta o valor no Total do Tab
        CustoTotalMaoDeObra.Caption = Format(gobjCusteioRoteiro.dCustoTotalMaoDeObra, "Standard")
        
        'Ajusta o valor no Grid de Itens do CusteioRot
        If gobjCusteioRoteiro.dCustoTotalMaoDeObra > 0 Then
            LabelDescCustoMO.Caption = Format(gobjCusteioRoteiro.dCustoTotalMaoDeObra, "Standard")
        Else
            LabelDescCustoMO.Caption = ""
        End If
        
        'Recalcula o Custo Total do Item e o Preço Total do Item
        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 139395

    End If

    AjustaCustoNecessidade_MaoDeObra = SUCESSO
    
    Exit Function
    
Erro_AjustaCustoNecessidade_MaoDeObra:

    AjustaCustoNecessidade_MaoDeObra = gErr
    
    Select Case gErr
        
        Case 139395
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158543)
            
    End Select
    
    Exit Function

End Function

Function Recalcula_Totais() As Long

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As ClassProduto
Dim dCustoTotalItem As Double

On Error GoTo Erro_Recalcula_Totais

    'Recalcula o Custo Total do Item e o Preço Total do Item
    dCustoTotalItem = gobjCusteioRoteiro.dCustoTotalInsumosKit + gobjCusteioRoteiro.dCustoTotalInsumosMaq + gobjCusteioRoteiro.dCustoTotalMaoDeObra
    
    sProduto = ProdutoRaiz.Text
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 139396
                
    LabelDescCustoTotal.Caption = Format(dCustoTotalItem, "Standard")
        
    Recalcula_Totais = SUCESSO
    
    Exit Function
        
Erro_Recalcula_Totais:

    Recalcula_Totais = gErr
    
    Select Case gErr

        Case 139396, 139397
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158544)

    End Select
    
    Exit Function
        
End Function

Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Codigo está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 139398
    
    'Verifica se o NomeReduzido está preenchido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 139399
    
    'Verifica se a Data do Custeio está preenchida
    If Len(Trim(DataCusteio.ClipText)) = 0 Then gError 139400
    
    'se o produto não está preenchido... erro
    If Len(Trim(ProdutoRaiz.Text)) = 0 Then gError 139401
    
    'se a versão não está preenchida... erro
    If Len(Trim(Versao.Text)) = 0 Then gError 139402
    
    'Preenche o gobjCusteioRotCusteio
    lErro = Move_Tela_Memoria()
    If lErro <> SUCESSO Then gError 139403

    lErro = Trata_Alteracao(gobjCusteioRoteiro, gobjCusteioRoteiro.lCodigo)
    If lErro <> SUCESSO Then gError 139404

    'Grava o CusteioRoteiro no Banco de Dados
    lErro = CF("CusteioRoteiro_Grava", gobjCusteioRoteiro)
    If lErro <> SUCESSO Then gError 139405
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 139398
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CUSTEIOROTEIRO_NAO_PREENCHIDO", gErr)

        Case 139399
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMERED_CUSTEIOROTEIRO_NAO_PREENCHIDO", gErr)

        Case 139400
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_CUSTEIOROTEIRO_NAO_PREENCHIDA", gErr)
                
        Case 139401
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_CUSTEIOROTEIRO_NAO_PREENCHIDO", gErr)
        
        Case 139402
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_CUSTEIOROTEIRO_NAO_PREENCHIDO", gErr)
                
        Case 139403 To 139405
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158545)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria() As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    If Len(Trim(Codigo.Text)) > 0 Then
        gobjCusteioRoteiro.lCodigo = StrParaLong(Trim(Codigo.Text))
    End If
    
    If Len(Trim(NomeReduzido.Text)) > 0 Then
        gobjCusteioRoteiro.sNomeReduzido = Trim(NomeReduzido.Text)
    End If

    If Len(Trim(Descricao.Text)) > 0 Then
        gobjCusteioRoteiro.sDescricao = Trim(Descricao.Text)
    End If

    If Len(Trim(DataCusteio.ClipText)) > 0 Then
        gobjCusteioRoteiro.dtDataCusteio = CDate(DataCusteio.Text)
    Else
        gobjCusteioRoteiro.dtDataCusteio = DATA_NULA
    End If

    If Len(Trim(DataValidade.ClipText)) > 0 Then
        gobjCusteioRoteiro.dtDataValidade = CDate(DataValidade.Text)
    Else
        gobjCusteioRoteiro.dtDataValidade = DATA_NULA
    End If
    
    If Len(Trim(Observacao.Text)) > 0 Then
        gobjCusteioRoteiro.sObservacao = Trim(Observacao.Text)
    End If
        
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158546)

    End Select

    Exit Function

End Function

Function IncluiNecessidade_InsumosMaquina(ByVal objMaquinasInsumos As ClassMaquinasInsumos, ByVal objPOMaquinas As ClassPOMaquinas, iSeq As Integer, dCustoTotalInsumosMaq As Double) As Long

Dim lErro As Long
Dim bProdutoNovo As Boolean
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim dCustoProduto As Double
Dim objCusteioRotInsumosMaq As ClassCusteioRotInsumosMaq
Dim objAuxCRInsumosMaquina As New ClassCusteioRotInsumosMaq

On Error GoTo Erro_IncluiNecessidade_InsumosMaquina

    'Descobre o fator de conversao da UM de Tempo utilizada p/UM de tempo padrão
    lErro = CF("UM_Conversao_Trans", gobjEST.iClasseUMTempo, objMaquinasInsumos.sUMTempo, TAXA_CONSUMO_TEMPO_PADRAO, dFator)
    If lErro <> SUCESSO Then gError 139406
    
    'Converte a quantidade
    dQuantidade = (objMaquinasInsumos.dQuantidade * dFator) * objPOMaquinas.dHorasMaquina
    
    'Vamos incluir na coleção...
    bProdutoNovo = True
    
    'Verifica se já há algum produto insumos de Maquinas na coleção
    For Each objAuxCRInsumosMaquina In gobjCusteioRoteiro.colCusteioRotInsumosMaq

        'se encontrou ...
        If objAuxCRInsumosMaquina.sProduto = objMaquinasInsumos.sProduto Then
                        
            'subtrai o valor anterior do Total do Custo no acumulador
            dCustoTotalInsumosMaq = dCustoTotalInsumosMaq - (objAuxCRInsumosMaquina.dQuantidade * objAuxCRInsumosMaquina.dCustoUnitarioInformado)
            
            'acumula a quantidade
            objAuxCRInsumosMaquina.dQuantidade = objAuxCRInsumosMaquina.dQuantidade + dQuantidade
            
            'lança o novo valor Total do Custo no acumulador
            dCustoTotalInsumosMaq = dCustoTotalInsumosMaq + (objAuxCRInsumosMaquina.dQuantidade * objAuxCRInsumosMaquina.dCustoUnitarioInformado)
            
            'avisa que acumulou ...
            bProdutoNovo = False
            Exit For
            
        End If
        
    Next
    
    'se não tem o produto ...
    If bProdutoNovo Then
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objMaquinasInsumos.sProduto

        'Descobre qual seu custo
        lErro = CF("Produto_Le_CustoProduto", objProduto, dCustoProduto)
        If lErro <> SUCESSO Then gError 139407
                                    
        'incrementa o sequencial
        iSeq = iSeq + 1
        
        'reCria o obj
        Set objCusteioRotInsumosMaq = New ClassCusteioRotInsumosMaq
        
        objCusteioRotInsumosMaq.iSeq = iSeq
        objCusteioRotInsumosMaq.sProduto = objMaquinasInsumos.sProduto
        objCusteioRotInsumosMaq.sUMedida = objMaquinasInsumos.sUMProduto
        objCusteioRotInsumosMaq.dQuantidade = dQuantidade
        objCusteioRotInsumosMaq.dCustoUnitarioCalculado = dCustoProduto
        objCusteioRotInsumosMaq.dCustoUnitarioInformado = dCustoProduto
        
        'lança o valor Total do Custo no acumulador
        dCustoTotalInsumosMaq = dCustoTotalInsumosMaq + (dQuantidade * dCustoProduto)
        
        'e inclui na coleção.
        gobjCusteioRoteiro.colCusteioRotInsumosMaq.Add objCusteioRotInsumosMaq
        
    End If
        
    IncluiNecessidade_InsumosMaquina = SUCESSO
    
    Exit Function
    
Erro_IncluiNecessidade_InsumosMaquina:

    IncluiNecessidade_InsumosMaquina = gErr
    
    Select Case gErr
    
        Case 139406, 139407
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158547)
    
    End Select
    
    Exit Function

End Function

Function IncluiNecessidade_MaoDeObra(ByVal objMaquinaOperadores As ClassMaquinaOperadores, ByVal objPOMaquinas As ClassPOMaquinas, iSeq As Integer, dCustoTotalMaoDeObra As Double) As Long

Dim lErro As Long
Dim bTipoNovo As Boolean
Dim objTipoMO As New ClassTiposDeMaodeObra
Dim dQuantidade As Double
Dim dCustoMO As Double
Dim objCusteioRotMaoDeObra As ClassCusteioRotMaoDeObra
Dim objAuxCRMaoDeObra As New ClassCusteioRotMaoDeObra

On Error GoTo Erro_IncluiNecessidade_MaoDeObra

    'Converte a quantidade
    dQuantidade = (objMaquinaOperadores.iQuantidade * objMaquinaOperadores.dPercentualUso) * objPOMaquinas.dHorasMaquina
    
    'Vamos incluir na coleção...
    bTipoNovo = True
    
    'Verifica se já há algum Tipo de MaoDeObra na coleção
    For Each objAuxCRMaoDeObra In gobjCusteioRoteiro.colCusteioRotMaoDeObra

        'se encontrou ...
        If objAuxCRMaoDeObra.iCodMO = objMaquinaOperadores.iTipoMaoDeObra Then
                        
            'subtrai o valor anterior do Total do Custo no acumulador
            dCustoTotalMaoDeObra = dCustoTotalMaoDeObra - (objAuxCRMaoDeObra.dQuantidade * objAuxCRMaoDeObra.dCustoUnitarioInformado)
            
            'acumula a quantidade
            objAuxCRMaoDeObra.dQuantidade = objAuxCRMaoDeObra.dQuantidade + dQuantidade
            
            'lança o novo valor Total do Custo no acumulador
            dCustoTotalMaoDeObra = dCustoTotalMaoDeObra + (objAuxCRMaoDeObra.dQuantidade * objAuxCRMaoDeObra.dCustoUnitarioInformado)
            
            'avisa que acumulou ...
            bTipoNovo = False
            Exit For
            
        End If
        
    Next
    
    'se não tem o tipo ...
    If bTipoNovo Then
    
        Set objTipoMO = New ClassTiposDeMaodeObra
        
        objTipoMO.iCodigo = objMaquinaOperadores.iTipoMaoDeObra

        'Descobre qual seu custo
        lErro = CF("TiposDeMaodeObra_Le", objTipoMO)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 139408
                                    
        'incrementa o sequencial
        iSeq = iSeq + 1
        
        'reCria o obj
        Set objCusteioRotMaoDeObra = New ClassCusteioRotMaoDeObra
        
        objCusteioRotMaoDeObra.iSeq = iSeq
        objCusteioRotMaoDeObra.iCodMO = objMaquinaOperadores.iTipoMaoDeObra
        objCusteioRotMaoDeObra.sUMedida = TAXA_CONSUMO_TEMPO_PADRAO
        objCusteioRotMaoDeObra.dQuantidade = dQuantidade
        objCusteioRotMaoDeObra.dCustoUnitarioCalculado = objTipoMO.dCustoHora
        objCusteioRotMaoDeObra.dCustoUnitarioInformado = objTipoMO.dCustoHora
        
        'lança o valor Total do Custo no acumulador
        dCustoTotalMaoDeObra = dCustoTotalMaoDeObra + (dQuantidade * objTipoMO.dCustoHora)
        
        'e inclui na coleção.
        gobjCusteioRoteiro.colCusteioRotMaoDeObra.Add objCusteioRotMaoDeObra
        
    End If
        
    IncluiNecessidade_MaoDeObra = SUCESSO
    
    Exit Function
    
Erro_IncluiNecessidade_MaoDeObra:

    IncluiNecessidade_MaoDeObra = gErr
    
    Select Case gErr
    
        Case 139408
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158548)
    
    End Select
    
    Exit Function

End Function

Private Sub Versao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Versao_GotFocus()

    Call MaskEdBox_TrataGotFocus(Versao, iAlterado)

End Sub

Private Sub Calcula_CustoTotalKit()

Dim dTotal As Double
Dim iLinha As Integer

    dTotal = 0
    For iLinha = 1 To objGridKit.iLinhasExistentes
    
        dTotal = dTotal + StrParaDbl(GridKit.TextMatrix(iLinha, iGrid_CustoTotalKit_Col))
    
    Next
    
    CustoTotalInsumosKit.Caption = Format(dTotal, "Standard")
    
End Sub

Private Sub Calcula_CustoTotalMaq()

Dim dTotal As Double
Dim iLinha As Integer

    dTotal = 0
    For iLinha = 1 To objGridMaquinas.iLinhasExistentes
    
        dTotal = dTotal + StrParaDbl(GridMaquinas.TextMatrix(iLinha, iGrid_CustoTotalMaq_Col))
    
    Next
    
    CustoTotalInsumosMaquina.Caption = Format(dTotal, "Standard")
    
End Sub

Private Sub Calcula_CustoTotalMO()

Dim dTotal As Double
Dim iLinha As Integer

    dTotal = 0
    For iLinha = 1 To objGridMaoDeObra.iLinhasExistentes
    
        dTotal = dTotal + StrParaDbl(GridMaoDeObra.TextMatrix(iLinha, iGrid_CustoTotalMO_Col))
    
    Next
    
    CustoTotalMaoDeObra.Caption = Format(dTotal, "Standard")
    
End Sub

Private Sub Limpa_DadosCusteio()

    LabelDescQuantidade.Caption = ""
    LabelDescUM.Caption = ""
    LabelDescCustoMP.Caption = ""
    LabelDescInsMaq.Caption = ""
    LabelDescCustoMO.Caption = ""
    LabelDescCustoTotal.Caption = ""

End Sub

