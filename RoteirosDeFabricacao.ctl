VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RoteirosDeFabricacao 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   1
      Left            =   195
      TabIndex        =   48
      Top             =   750
      Width           =   9105
      Begin VB.CommandButton BotaoOnde 
         Caption         =   "Onde é Usado ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2730
         Picture         =   "RoteirosDeFabricacao.ctx":0000
         TabIndex        =   14
         ToolTipText     =   "Lista dos Roteiros de Fabricação onde este roteiro é utilizado"
         Top             =   4365
         Width           =   1200
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relatório"
         Height          =   870
         Left            =   5355
         TabIndex        =   104
         Top             =   4170
         Width           =   3675
         Begin VB.CommandButton BotaoRelRoteiro 
            Caption         =   "Roteiro Completo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   2295
            TabIndex        =   16
            ToolTipText     =   "Abre o Relatório de Roteiro de Fabricação"
            Top             =   210
            Width           =   1200
         End
         Begin VB.CheckBox DetalharInsumos 
            Caption         =   "Detalhar Insumos"
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
            Left            =   195
            TabIndex        =   15
            Top             =   345
            Width           =   1905
         End
      End
      Begin VB.CommandButton BotaoKits 
         Caption         =   "&Kits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1425
         Picture         =   "RoteirosDeFabricacao.ctx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Visualiza os Kits cadastrados"
         Top             =   4365
         Width           =   1200
      End
      Begin VB.CommandButton BotaoVerRoteiros 
         Caption         =   "&Roteiros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   120
         Picture         =   "RoteirosDeFabricacao.ctx":0614
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Roteiros de Fabricação cadastrados"
         Top             =   4365
         Width           =   1200
      End
      Begin VB.Frame Frame8 
         Caption         =   "Dados do Kit"
         Height          =   720
         Left            =   120
         TabIndex        =   95
         Top             =   3450
         Width           =   8910
         Begin VB.ComboBox Composicao 
            Height          =   315
            ItemData        =   "RoteirosDeFabricacao.ctx":091E
            Left            =   4290
            List            =   "RoteirosDeFabricacao.ctx":0928
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   270
            Width           =   1350
         End
         Begin MSMask.MaskEdBox PercentualPerda 
            Height          =   315
            Left            =   1215
            TabIndex        =   9
            Top             =   270
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoStandard 
            Height          =   315
            Left            =   7215
            TabIndex        =   11
            Top             =   255
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Custo Standard:"
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
            Left            =   5805
            TabIndex        =   98
            Top             =   300
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Composição:"
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
            Left            =   3120
            TabIndex        =   97
            Top             =   300
            Width           =   1095
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   555
            TabIndex        =   96
            Top             =   300
            Width           =   570
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Datas"
         Height          =   1185
         Index           =   1
         Left            =   2505
         TabIndex        =   66
         Top             =   2235
         Width           =   6510
         Begin MSComCtl2.UpDown UpDownDataCriacao 
            Height          =   300
            Left            =   3060
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCriacao 
            Height          =   315
            Left            =   1890
            TabIndex        =   7
            Top             =   285
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataUltModificacao 
            Caption         =   "Última modificação:"
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
            Left            =   150
            TabIndex        =   57
            Top             =   780
            Width           =   1755
         End
         Begin VB.Label LabelAutor 
            Alignment       =   1  'Right Justify
            Caption         =   "Autor:"
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
            Left            =   4110
            TabIndex        =   70
            Top             =   765
            Width           =   660
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data de criação:"
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
            Left            =   375
            TabIndex        =   69
            Top             =   330
            Width           =   1440
         End
         Begin VB.Label DataUltModificacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1890
            TabIndex        =   68
            Top             =   735
            Width           =   1410
         End
         Begin VB.Label Autor 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4830
            TabIndex        =   67
            Top             =   735
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Quantidade"
         Height          =   1185
         Left            =   105
         TabIndex        =   63
         Top             =   2220
         Width           =   2325
         Begin VB.ComboBox UM 
            Height          =   315
            Left            =   1200
            TabIndex        =   6
            Top             =   750
            Width           =   975
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   1230
            TabIndex        =   5
            Top             =   300
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   105
            TabIndex        =   65
            Top             =   330
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   75
            TabIndex        =   64
            Top             =   780
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificação"
         Height          =   2175
         Left            =   105
         TabIndex        =   58
         Top             =   15
         Width           =   8895
         Begin VB.CheckBox VersaoFormPreco 
            Caption         =   "Usar para Formação de Preços"
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
            Left            =   240
            TabIndex        =   4
            Top             =   1815
            Width           =   3045
         End
         Begin MSMask.MaskEdBox ProdutoRaiz 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   270
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   315
            Left            =   1260
            TabIndex        =   2
            Top             =   1050
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Versao 
            Height          =   315
            Left            =   1260
            TabIndex        =   1
            Top             =   690
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumMaxMaqPorOper 
            Height          =   315
            Left            =   5805
            TabIndex        =   3
            Top             =   1485
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Número Máximo Padrão de Máquinas usadas em cada Operação:"
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
            TabIndex        =   105
            Top             =   1530
            Width           =   5535
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
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   62
            Top             =   330
            Width           =   735
         End
         Begin VB.Label DescricaoProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2865
            TabIndex        =   61
            Top             =   285
            Width           =   5880
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
            Left            =   525
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   60
            Top             =   735
            Width           =   660
         End
         Begin VB.Label LabelDescricao 
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
            Left            =   240
            TabIndex        =   59
            Top             =   1125
            Width           =   930
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5085
      Index           =   2
      Left            =   225
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   9060
      Begin VB.Frame Frame2 
         Caption         =   "Operação"
         Height          =   2625
         Index           =   1
         Left            =   90
         TabIndex        =   56
         Top             =   2325
         Width           =   8775
         Begin VB.CommandButton BotaoMaquinasAptas 
            Caption         =   "Máquinas Aptas"
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
            Left            =   6690
            Picture         =   "RoteirosDeFabricacao.ctx":093C
            TabIndex        =   107
            ToolTipText     =   "Lista as máquinas aptas a fazer a compeência especificada."
            Top             =   2085
            Width           =   1935
         End
         Begin VB.TextBox Observacao 
            Height          =   315
            Left            =   1830
            MaxLength       =   255
            TabIndex        =   27
            Top             =   1245
            Width           =   6795
         End
         Begin MSMask.MaskEdBox CodigoCTPadrao 
            Height          =   315
            Left            =   1830
            TabIndex        =   26
            Top             =   780
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoCompetencia 
            Height          =   315
            Left            =   1845
            TabIndex        =   25
            Top             =   330
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumMaxMaqPorOperItem 
            Height          =   315
            Left            =   5370
            TabIndex        =   28
            Top             =   1680
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Repeticao 
            Height          =   315
            Left            =   8175
            TabIndex        =   29
            Top             =   1680
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label LabelRepeticao 
            AutoSize        =   -1  'True
            Caption         =   "Número de Repetições:"
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
            Left            =   6150
            TabIndex        =   111
            Top             =   1740
            Width           =   2010
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Número Máximo de Máquinas usadas em cada Operação:"
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
            Left            =   480
            TabIndex        =   106
            Top             =   1740
            Width           =   4875
         End
         Begin VB.Label DescricaoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4320
            TabIndex        =   94
            Top             =   780
            Width           =   4305
         End
         Begin VB.Label DescricaoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4320
            TabIndex        =   93
            Top             =   330
            Width           =   4305
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nível:"
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
            Left            =   1110
            TabIndex        =   77
            Top             =   2175
            Width           =   540
         End
         Begin VB.Label Sequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3615
            TabIndex        =   76
            Top             =   2130
            Width           =   420
         End
         Begin VB.Label Nivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1830
            TabIndex        =   75
            Top             =   2130
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sequencial:"
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
            Left            =   2460
            TabIndex        =   74
            Top             =   2175
            Width           =   1020
         End
         Begin VB.Label Label3 
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
            Left            =   510
            TabIndex        =   71
            Top             =   1275
            Width           =   1155
         End
         Begin VB.Label CTLabel 
            Caption         =   "CT Padrão:"
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
            Left            =   645
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   73
            Top             =   810
            Width           =   990
         End
         Begin VB.Label CompetenciaLabel 
            Caption         =   "Competência:"
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
            Height          =   330
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   72
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Insumos por Operações"
         Height          =   2655
         Index           =   2
         Left            =   90
         TabIndex        =   51
         Top             =   2295
         Visible         =   0   'False
         Width           =   8775
         Begin VB.CommandButton BotaoAbrirRoteiro 
            Caption         =   "&Abrir Roteiro do Insumo"
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
            Left            =   4295
            TabIndex        =   33
            ToolTipText     =   "Abre a tela de Roteiro de Fabricação para o Insumo"
            Top             =   2145
            Width           =   2400
         End
         Begin VB.TextBox OrigemProduto 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3045
            TabIndex        =   103
            Top             =   1050
            Width           =   375
         End
         Begin VB.ComboBox VersaoKitComp 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "RoteirosDeFabricacao.ctx":0C46
            Left            =   4590
            List            =   "RoteirosDeFabricacao.ctx":0C48
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   1050
            Width           =   855
         End
         Begin VB.ComboBox ComposicaoInsumos 
            Height          =   315
            ItemData        =   "RoteirosDeFabricacao.ctx":0C4A
            Left            =   6255
            List            =   "RoteirosDeFabricacao.ctx":0C54
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   1050
            Width           =   990
         End
         Begin MSMask.MaskEdBox PerdaInsumos 
            Height          =   315
            Left            =   5520
            TabIndex        =   100
            Top             =   1050
            Width           =   735
            _ExtentX        =   1296
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
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoStandardInsumos 
            Height          =   315
            Left            =   7410
            TabIndex        =   101
            Top             =   1050
            Width           =   765
            _ExtentX        =   1349
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
         Begin VB.CommandButton BotaoProdutos 
            Caption         =   "&Produtos"
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
            Left            =   105
            TabIndex        =   31
            ToolTipText     =   "Abre o Browse de Produtos"
            Top             =   2145
            Width           =   1485
         End
         Begin VB.CommandButton BotaoLimparGrid 
            Caption         =   "&Limpar Grid"
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
            Left            =   7005
            TabIndex        =   34
            ToolTipText     =   "Limpa os dados do Grid"
            Top             =   2130
            Width           =   1620
         End
         Begin VB.ComboBox UMProduto 
            Height          =   315
            Left            =   3705
            TabIndex        =   53
            Top             =   1035
            Width           =   720
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1200
            TabIndex        =   52
            Top             =   1035
            Width           =   1770
         End
         Begin VB.CommandButton BotaoTrazerInsumosKit 
            Caption         =   "Trazer &Insumos Kit"
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
            Left            =   1900
            TabIndex        =   32
            ToolTipText     =   "Traz os Insumos vinculados ao Kit"
            Top             =   2145
            Width           =   2085
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   465
            TabIndex        =   54
            Top             =   600
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProduto 
            Height          =   315
            Left            =   2970
            TabIndex        =   55
            Top             =   450
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   15
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridOperacaoInsumos 
            Height          =   1500
            Left            =   60
            TabIndex        =   30
            Top             =   225
            Width           =   8610
            _ExtentX        =   15187
            _ExtentY        =   2646
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Produção"
         Height          =   2670
         Index           =   3
         Left            =   90
         TabIndex        =   78
         Top             =   2280
         Visible         =   0   'False
         Width           =   8775
         Begin VB.Frame FrameTaxas 
            Caption         =   "Taxa"
            Height          =   1545
            Left            =   180
            TabIndex        =   87
            Top             =   1020
            Width           =   3600
            Begin VB.ComboBox Tipo 
               Height          =   315
               ItemData        =   "RoteirosDeFabricacao.ctx":0C68
               Left            =   915
               List            =   "RoteirosDeFabricacao.ctx":0C6A
               TabIndex        =   36
               Top             =   240
               Width           =   2520
            End
            Begin VB.ComboBox UMTempo 
               Height          =   315
               Left            =   2415
               TabIndex        =   38
               Top             =   675
               Width           =   1050
            End
            Begin MSMask.MaskEdBox TempoOperacao 
               Height          =   315
               Left            =   915
               TabIndex        =   37
               Top             =   675
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin VB.Label LabelUMTempo 
               Alignment       =   1  'Right Justify
               Caption         =   "UM:"
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
               Height          =   300
               Left            =   1965
               TabIndex        =   92
               Top             =   720
               Width           =   405
            End
            Begin VB.Label LabelTaxa 
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
               Height          =   315
               Left            =   315
               TabIndex        =   91
               Top             =   1140
               Width           =   585
            End
            Begin VB.Label LabelTaxaDeProducao 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   915
               TabIndex        =   90
               Top             =   1110
               Width           =   2520
            End
            Begin VB.Label LabelTipo 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo:"
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
               Height          =   300
               Left            =   120
               TabIndex        =   89
               Top             =   270
               Width           =   660
            End
            Begin VB.Label LabelTempoOperacao 
               Alignment       =   1  'Right Justify
               Caption         =   "Tempo:"
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
               Left            =   90
               TabIndex        =   88
               Top             =   705
               Width           =   705
            End
         End
         Begin VB.Frame FrameLotes 
            Caption         =   "Lotes"
            Height          =   1545
            Left            =   6540
            TabIndex        =   83
            Top             =   1020
            Width           =   2040
            Begin MSMask.MaskEdBox LoteMinimo 
               Height          =   315
               Left            =   960
               TabIndex        =   42
               Top             =   240
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox LoteMaximo 
               Height          =   315
               Left            =   960
               TabIndex        =   43
               Top             =   675
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox LotePadrao 
               Height          =   315
               Left            =   960
               TabIndex        =   44
               Top             =   1110
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin VB.Label LabelPadrao 
               Caption         =   "Padrão:"
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
               Left            =   225
               TabIndex        =   86
               Top             =   1140
               Width           =   735
            End
            Begin VB.Label LabelLoteMin 
               Caption         =   "Mínimo:"
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
               Left            =   195
               TabIndex        =   85
               Top             =   270
               Width           =   750
            End
            Begin VB.Label LabelLoteMax 
               Caption         =   "Máximo:"
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
               Left            =   195
               TabIndex        =   84
               Top             =   705
               Width           =   735
            End
         End
         Begin VB.Frame FrameTempos 
            Caption         =   "Tempos (Horas)"
            Height          =   1545
            Left            =   3885
            TabIndex        =   79
            Top             =   1020
            Width           =   2550
            Begin MSMask.MaskEdBox TempoPreparacao 
               Height          =   315
               Left            =   1620
               TabIndex        =   39
               Top             =   240
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TempoDescarga 
               Height          =   315
               Left            =   1620
               TabIndex        =   41
               Top             =   1110
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TempoMovimentacao 
               Height          =   315
               Left            =   1620
               TabIndex        =   40
               Top             =   675
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin VB.Label LabelTempoMovimentacao 
               Caption         =   "Movimentação:"
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
               Left            =   150
               TabIndex        =   82
               Top             =   705
               Width           =   1335
            End
            Begin VB.Label LabelTempoPreparacao 
               Caption         =   "Preparação:"
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
               Left            =   390
               TabIndex        =   81
               Top             =   270
               Width           =   1065
            End
            Begin VB.Label LabelTempoDescarga 
               Caption         =   "Descarga:"
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
               Left            =   585
               TabIndex        =   80
               Top             =   1140
               Width           =   870
            End
         End
         Begin VB.CheckBox IgnoraTaxa 
            Caption         =   "Ignora o cadastro de Taxa de Produção"
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
            Left            =   195
            TabIndex        =   47
            Top             =   255
            Width           =   3870
         End
         Begin MSMask.MaskEdBox Maquina 
            Height          =   315
            Left            =   1095
            TabIndex        =   108
            Top             =   615
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelMaquina 
            Alignment       =   1  'Right Justify
            Caption         =   "Máquina:"
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
            Left            =   60
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   110
            Top             =   645
            Width           =   900
         End
         Begin VB.Label DescMaquina 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3120
            TabIndex        =   109
            Top             =   615
            Width           =   5445
         End
      End
      Begin VB.CommandButton BotaoIncluir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7605
         Picture         =   "RoteirosDeFabricacao.ctx":0C6C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Inclui a Operação na Árvore do Roteiro"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton BotaoRemover 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7605
         Picture         =   "RoteirosDeFabricacao.ctx":24BA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exclui a Operação da Árvore do Roteiro"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton BotaoAlterar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7605
         Picture         =   "RoteirosDeFabricacao.ctx":3DE0
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Altera a Operação da Árvore do Roteiro"
         Top             =   1305
         Width           =   1335
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   3150
         Left            =   15
         TabIndex        =   46
         Top             =   1905
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   5556
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhe"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Insumos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tempo de Produção"
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
      Begin VB.Frame Frame4 
         Caption         =   "Roteiro de Fabricação:"
         Height          =   1785
         Index           =   1
         Left            =   0
         TabIndex        =   50
         Top             =   30
         Width           =   7485
         Begin MSComctlLib.TreeView Roteiro 
            Height          =   1380
            Left            =   150
            TabIndex        =   21
            Top             =   270
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2434
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   354
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   1
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7290
      ScaleHeight     =   480
      ScaleWidth      =   2025
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   105
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RoteirosDeFabricacao.ctx":5706
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "RoteirosDeFabricacao.ctx":5860
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "RoteirosDeFabricacao.ctx":59EA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "RoteirosDeFabricacao.ctx":5F1C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5520
      Left            =   120
      TabIndex        =   45
      Top             =   405
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   9737
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Roteiro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Operações"
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
Attribute VB_Name = "RoteirosDeFabricacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iCompetenciaAlterada As Integer
Dim iFrameAtual As Integer
Dim iFrameAtualOper As Integer

Dim colComponentes As New Collection
Dim iProxChave As Integer
Dim bOperacaoNova As Boolean

Dim glNumIntCompetAnt As Long
Dim glNumIntCTAnt As Long
Dim gsProdutoAnt As String

'variaveis auxiliares para recalculo de nivel e sequencial
Dim aNivelSequencial(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o maior sequencial
Dim aSeqPai(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o SeqPai

Dim iUltimoNivel As Integer

'Formato para quantidades de Produtos
Const FORMATO_ESTOQUE_KIT = "#,##0.0####"

'Grid de OperacaoInsumos
Dim objGridOperacaoInsumos As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_OrigemProduto_Col As Integer
Dim iGrid_QuantidadeProduto_Col As Integer
Dim iGrid_UMProduto_Col As Integer
Dim iGrid_VersaoKitComp_Col As Integer
Dim iGrid_PerdaInsumos_Col As Integer
Dim iGrid_ComposicaoInsumos_Col As Integer
Dim iGrid_CustoStandardInsumos_Col As Integer

Private WithEvents objEventoRoteiroDeFabricacao As AdmEvento
Attribute objEventoRoteiroDeFabricacao.VB_VarHelpID = -1
Private WithEvents objEventoCompetencias As AdmEvento
Attribute objEventoCompetencias.VB_VarHelpID = -1
Private WithEvents objEventoCentroDeTrabalho As AdmEvento
Attribute objEventoCentroDeTrabalho.VB_VarHelpID = -1
Private WithEvents objEventoInsumos As AdmEvento
Attribute objEventoInsumos.VB_VarHelpID = -1
Private WithEvents objEventoKit As AdmEvento
Attribute objEventoKit.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1
Private WithEvents objEventoProdutoRaiz As AdmEvento
Attribute objEventoProdutoRaiz.VB_VarHelpID = -1
Private WithEvents objEventoCusteio As AdmEvento
Attribute objEventoCusteio.VB_VarHelpID = -1
Private WithEvents objEventoMaquina As AdmEvento
Attribute objEventoMaquina.VB_VarHelpID = -1


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Roteiros de Fabricação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RoteirosDeFabricacao"

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

Private Sub BotaoAlterar_Click()

Dim lErro As Long
Dim objNode As Node
Dim sChave As String
Dim objOperacoes As ClassOperacoes
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objProduto As ClassProduto
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim sTexto As String

On Error GoTo Erro_BotaoAlterar_Click

    Set objNode = Roteiro.SelectedItem

    If objNode Is Nothing Then gError 134680
    If objNode.Selected = False Then gError 134681
    
    If Len(Trim(CodigoCompetencia.ClipText)) = 0 Then gError 134973
    
    Set objOperacoes = New ClassOperacoes
    Set objCompetencias = New ClassCompetencias
    Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
    'preenche objOperacoes à partir dos dados da tela
    lErro = Move_Operacoes_Memoria(objOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134682

    sChave = objNode.Tag
        
    'prepara texto que identificará a nova Operação que está sendo incluida
    
    sTexto = objCompetencias.sNomeReduzido
            
    Set objProduto = New ClassProduto
    
    sCodProduto = ProdutoRaiz.Text
        
    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sCodProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 137695
            
    sTexto = sTexto & " (" & objProduto.sNomeReduzido

    If Len(Trim(CodigoCTPadrao.ClipText)) <> 0 Then
       sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
    End If
        
    sTexto = sTexto & ")"

    objNode.Text = sTexto

    colComponentes.Remove (sChave)
    colComponentes.Add objOperacoes, sChave

    Call Recalcula_Nivel_Sequencial

    'Limpa a tab de operacoes
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 134683

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoAlterar_Click:

    Select Case gErr

        Case 134680, 134681
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 134682, 134683, 137694, 137695
            'erro tratado na rotina chamada
        
        Case 134973
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174221)

    End Select

    Exit Sub

End Sub

'Private Sub BotaoCusteio_Click()
'
'Dim lErro As Long
'Dim objCusteioRoteiro As New ClassCusteioRoteiro
'Dim sProduto As String
'Dim sVersao As String
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim sFiltro As String
'Dim colSelecao As New Collection
'
'On Error GoTo Erro_BotaoCusteio_Click
'
'    sProduto = ProdutoRaiz.Text
'    sVersao = Versao.Text
'
'    'formata o código do produto que está no grid
'    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
'    If lErro <> SUCESSO Then gError 139109
'
'    sFiltro = "Produto = ? And Versao = ? "
'
'    colSelecao.Add sProdutoFormatado
'    colSelecao.Add sVersao
'
'    Call Chama_Tela("CusteioRoteirosLista", colSelecao, Nothing, objEventoCusteio, sFiltro)
'
'    Exit Sub
'
'Erro_BotaoCusteio_Click:
'
'    Select Case gErr
'
'        Case 139109
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174222)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub BotaoIncluir_Click()

Dim lErro As Long, iIndice As Integer
Dim sChave As String, sTexto As String
Dim objNode As Node, objNodePai As Node, sChaveTvw As String
Dim iNivel As Integer
Dim objOperacoes As New ClassOperacoes
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objProduto As ClassProduto
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoIncluir_Click

    If Len(Trim(CodigoCompetencia.ClipText)) = 0 Then gError 134684
    
    Set objOperacoes = New ClassOperacoes
    Set objCompetencias = New ClassCompetencias
    Set objCentrodeTrabalho = New ClassCentrodeTrabalho

    lErro = Move_Operacoes_Memoria(objOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134685

    Set objNodePai = Roteiro.SelectedItem

    If objNodePai Is Nothing Then
    
        iNivel = 0
        
    Else
    
        If objNodePai.Selected = False Then gError 134686
        
        iNivel = objNodePai.Index + 1
    
    End If

    'prepara texto que identificará a nova Operação que está sendo incluida
    sTexto = objCompetencias.sNomeReduzido
            
    Set objProduto = New ClassProduto
    
    sCodProduto = ProdutoRaiz.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sCodProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 137697
            
    sTexto = sTexto & " (" & objProduto.sNomeReduzido

    If Len(Trim(CodigoCTPadrao.ClipText)) <> 0 Then
       sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
    End If
        
    sTexto = sTexto & ")"

    'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
    Call Calcula_Proxima_Chave(sChaveTvw)
    
    sChave = sChaveTvw
    sChaveTvw = sChaveTvw & CodigoCompetencia.ClipText
    
    'inclui o node na treeview
    If iNivel = 0 Then
    
        Set objNode = Roteiro.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)
    
    Else
        
        Set objNode = Roteiro.Nodes.Add(objNodePai.Index, tvwChild, sChaveTvw, sTexto)
        
        If Not Roteiro.Nodes.Item(objNodePai.Index).Expanded Then
    
            Roteiro.Nodes.Item(objNodePai.Index).Expanded = True
                
        End If
    
    End If
    
    colComponentes.Add objOperacoes, sChave
    objNode.Tag = sChave

    Call Recalcula_Nivel_Sequencial

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 134687

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 134686
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 134685, 134687
            'erro tratado na rotina chamada

        Case 134684
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174223)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRemover_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objNode As Node
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoRemover_Click

    Set objNode = Roteiro.SelectedItem

    If objNode Is Nothing Then gError 134688
    If objNode.Selected = False Then gError 134689

    If objNode.Children > 0 Then

        'Envia aviso perguntando se realmente deseja excluir
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_TEM_FILHOS")

        If vbMsgRes = vbNo Then Exit Sub

        'chama rotina que exclui filhos
        Call Remove_Filhos(objNode.Child)
    
    End If

    colComponentes.Remove (objNode.Tag)
    Roteiro.Nodes.Remove (objNode.Index)
    
    Call Recalcula_Nivel_Sequencial

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 134690

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoRemover_Click:

    Select Case gErr

        Case 134688, 134689
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 134690

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174224)

    End Select

    Exit Sub

End Sub

Private Sub CodigoCompetencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoCompetencia, iAlterado)
    
End Sub

Private Sub CodigoCompetencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho

On Error GoTo Erro_CodigoCompetencia_Validate

    DescricaoCompetencia.Caption = ""

    'Verifica se CodigoCompetencia não está preenchida
    If Len(Trim(CodigoCompetencia.Text)) <> 0 Then
    
        Set objCompetencias = New ClassCompetencias
        
        'Verifica sua existencia
        lErro = CF("TP_Competencia_Le", CodigoCompetencia, objCompetencias)
        If lErro <> SUCESSO Then gError 134693
        
        If glNumIntCompetAnt <> objCompetencias.lNumIntDoc Then
        
            Observacao.Text = ""
            
            Call Grid_Limpa(objGridOperacaoInsumos)
            
            IgnoraTaxa.Value = vbUnchecked
            
            Call Habilita_TaxaDeProducao
                
        End If
        
        DescricaoCompetencia.Caption = objCompetencias.sDescricao
        
        'Verifica se CodigoCompetencia foi alterado
        If iCompetenciaAlterada = REGISTRO_ALTERADO Then
                
            CodigoCTPadrao.Text = ""
            DescricaoCTPadrao.Caption = ""
            
            'Verifica se existe CTPadrao cadastrado na Competencia e traz seus dados
            lErro = CF("Competencias_Le_CTPadrao", objCompetencias, objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134909 Then gError 134910
            
            If lErro = SUCESSO Then
            
               CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
               
               Call CodigoCTPadrao_Validate(bSGECancelDummy)
            
            End If
            
            iCompetenciaAlterada = 0
        
        End If
       
    End If

    Exit Sub

Erro_CodigoCompetencia_Validate:

    Cancel = True

    Select Case gErr

        Case 134693, 134910
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174225)

    End Select

    Exit Sub

End Sub

Private Sub CodigoCompetencia_Change()

    iAlterado = REGISTRO_ALTERADO
    iCompetenciaAlterada = REGISTRO_ALTERADO

End Sub

Private Sub CodigoCTPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoCTPadrao, iAlterado)
    
End Sub

Private Sub CodigoCTPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objCTCompetencias As New ClassCTCompetencias
Dim objCompetencias As ClassCompetencias
Dim bCompetenciaCadastrada As Boolean

On Error GoTo Erro_CodigoCTPadrao_Validate

    DescricaoCTPadrao.Caption = ""

    'Verifica se CodigoCTPadrao não está preenchido
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then

        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CodigoCTPadrao, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 134918
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = CodigoCompetencia.Text
        
        'Lê a Competencia pelo NomeReduzido para verificar seu NumIntDoc
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 134920
    
        If lErro <> SUCESSO Then gError 134921
        
        lErro = CF("CentrodeTrabalho_Le_CTCompetencias", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134453 Then gError 134919
    
        bCompetenciaCadastrada = False
        
        For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
        
            If objCTCompetencias.lNumIntDocCompet = objCompetencias.lNumIntDoc Then
            
                bCompetenciaCadastrada = True
                Exit For
                
            End If
        
        Next
            
        If bCompetenciaCadastrada = False Then gError 134922
            
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
       
    End If
    
    lErro = Trata_TaxaDeProducao
    If lErro <> SUCESSO Then gError 180356
    
    Exit Sub

Erro_CodigoCTPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 134918, 134919, 134920, 180356
            'erro tratado na rotina chamada
        
        Case 134921, 134922
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_NAO_CADASTRADA_CT", gErr, objCentrodeTrabalho.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174226)

    End Select

    Exit Sub

End Sub

Private Sub CodigoCTPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CTLabel_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_CTLabel

    'Verifica se o CodigoCTPadrao foi preenchido
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Verifica o CodigoCTPadrao, lendo no BD a partir do NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 139065
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCentroDeTrabalho)

    Exit Sub

Erro_CTLabel:

    Select Case gErr
    
        Case 139065
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174227)

    End Select

    Exit Sub

End Sub

Private Sub CompetenciaLabel_Click()

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
Dim colSelecao As New Collection

On Error GoTo Erro_CompetenciaLabel_Click

    'Verifica se a Competencia foi preenchida
    If Len(Trim(CodigoCompetencia.Text)) <> 0 Then
            
        objCompetencias.sNomeReduzido = CodigoCompetencia.Text

        'Verifica a Competencia no BD a partir do NomeReduzido
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 139064

    End If

    Call Chama_Tela("CompetenciasLista", colSelecao, objCompetencias, objEventoCompetencias)

    Exit Sub

Erro_CompetenciaLabel_Click:

    Select Case gErr
    
        Case 139064
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174228)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridOperacaoInsumos.Row = 0 Then gError 134698
        
        sProduto = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_Produto_Col)
        
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134699
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutosKitLista", colSelecao, objProduto, objEventoInsumos)
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 134699
        
        Case 134698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174229)

    End Select

    Exit Sub

End Sub

Private Sub BotaoKits_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoKits_Click

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 139266
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objKit.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(Versao.Text)) <> 0 Then
    
        objKit.sVersao = Versao.Text
    
    Else
    
        'Le as Versoes Ativas e a Padrao
        lErro = CF("Kit_Le_Padrao", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then gError 139267
    
    End If

    'Chama tela
    Call Chama_Tela("KitLista", colSelecao, objKit, objEventoKit)

    Exit Sub

Erro_BotaoKits_Click:

    Select Case gErr
    
        Case 139266, 139267
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174230)

    End Select

    Exit Sub

End Sub


Private Sub BotaoLimparGrid_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimparGrid_Click

    Call Grid_Limpa(objGridOperacaoInsumos)

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoLimparGrid_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174231)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazerInsumosKit_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objKit As ClassKit
Dim objKitAux As New ClassKit
Dim iLinha As Integer
Dim objProdutoKit As New ClassProdutoKit
Dim objProdutoKitAux As New ClassProdutoKit
Dim dFatorUMKit As Double
Dim objProduto As New ClassProduto
Dim dFatorQuantidade As Double
Dim dQuantidade As Double
Dim sProdutoMascarado As String
Dim objOperacoes As New ClassOperacoes
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim dFatorUMCol As Double
Dim bExisteProduto As Boolean
Dim dFator As Double

On Error GoTo Erro_BotaoTrazerInsumosKit_Click
    
    'Formata o Produto e verifica ...
    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134701

    'se o produto raiz existe cadastrado
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 137163
    
    'Verifica se a Versão está preenchida
    If Len(Trim(Versao.Text)) = 0 Then gError 137679
    
    'Verifica se a Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then gError 134702
    
    'Verifica se a U.M. está preenchida
    If Len(Trim(UM.Text)) = 0 Then gError 137680

    Set objKit = New ClassKit
    
    objKit.sProdutoRaiz = sProdutoFormatado
    objKit.sVersao = Versao.Text
    
    'Leio os ProdutosKits que compõem este Kit
    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then gError 134703
    
    'Se o Kit não existir... erro
    If lErro <> SUCESSO Then gError 134704

   
    For Each objProdutoKit In objKit.colComponentes
        
        bExisteProduto = False
        For Each objProdutoKitAux In objKitAux.colComponentes
    
            If objProdutoKitAux.sProduto = objProdutoKit.sProduto And objProdutoKitAux.iNivel <> KIT_NIVEL_RAIZ Then
                If objProdutoKitAux.iComposicao <> objProdutoKit.iComposicao Or _
                    objProdutoKitAux.dCustoStandard <> objProdutoKit.dCustoStandard Or _
                    objProdutoKitAux.dPercentualPerda <> objProdutoKit.dPercentualPerda Or _
                    objProdutoKitAux.sVersaoKitComp <> objProdutoKit.sVersaoKitComp Then gError 181410
                bExisteProduto = True
                Exit For
            End If
    
        Next
        
        If Not bExisteProduto Then
        
            Set objProdutoKitAux = New ClassProdutoKit
            
            objProdutoKitAux.dCustoStandard = objProdutoKit.dCustoStandard
            objProdutoKitAux.dPercentualPerda = objProdutoKit.dPercentualPerda
            objProdutoKitAux.dQuantidade = objProdutoKit.dQuantidade
            objProdutoKitAux.iClasseUM = objProdutoKit.iClasseUM
            objProdutoKitAux.iComposicao = objProdutoKit.iComposicao
            objProdutoKitAux.iNivel = objProdutoKit.iNivel
            objProdutoKitAux.iPosicaoArvore = objProdutoKit.iPosicaoArvore
            objProdutoKitAux.iSeq = objProdutoKit.iSeq
            objProdutoKitAux.iSeqPai = objProdutoKit.iSeqPai
            objProdutoKitAux.sProduto = objProdutoKit.sProduto
            objProdutoKitAux.sProdutoDesc = objProdutoKit.sProdutoDesc
            objProdutoKitAux.sProdutoRaiz = objProdutoKit.sProdutoRaiz
            objProdutoKitAux.sUnidadeMed = objProdutoKit.sUnidadeMed
            objProdutoKitAux.sVersao = objProdutoKit.sVersao
            objProdutoKitAux.sVersaoKitComp = objProdutoKit.sVersaoKitComp

            objKitAux.colComponentes.Add objProdutoKitAux
        
        Else
        
            Set objProduto = New ClassProduto
        
            objProduto.sCodigo = objProdutoKit.sProduto
            
            'Lê o Produto Componente do Kit
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 134705
            
            'se o produto não estiver cadastrado... erro
            If lErro <> SUCESSO Then gError 134706
               
            'Descobre o fator de conversao da UM do Roteiro p/UM do Kit do Produto Raiz
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProdutoKitAux.sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 137681
            
            objProdutoKitAux.dQuantidade = objProdutoKitAux.dQuantidade + (objProdutoKit.dQuantidade * dFator)
   
        End If
        
    Next
    
    'inicializa o contador de linhas
    iLinha = 0
    
    'para cada produto componente do kit ...
    For Each objProdutoKit In objKitAux.colComponentes
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objProdutoKit.sProduto
        
        'Lê o Produto Componente do Kit
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134705
        
        'se o produto não estiver cadastrado... erro
        If lErro <> SUCESSO Then gError 134706

        'se é o produto raiz ...
        If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
        
            'Descobre o fator de conversao da UM do Roteiro p/UM do Kit do Produto Raiz
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, UM.Text, objProdutoKit.sUnidadeMed, dFatorUMKit)
            If lErro <> SUCESSO Then gError 137681
                        
            dFatorQuantidade = StrParaDbl(Quantidade.Text) * dFatorUMKit / objProdutoKit.dQuantidade
        
        Else
            
            'se é de composição variável...
            If objProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                                        
                dQuantidade = objProdutoKit.dQuantidade * dFatorQuantidade
            
            Else
            
                dQuantidade = objProdutoKit.dQuantidade
            
            End If
            
            'Verifica se já há algum produto do kit na coleção
            For Each objOperacoes In colComponentes
                       
                For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos
                
                    'se encontrou ...
                    If objOperacaoInsumos.sProduto = objProdutoKit.sProduto Then
                    
                        'Descobre o fator de conversao da UM da Coleção p/UM do Kit do Produto Raiz
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOperacaoInsumos.sUMProduto, objProdutoKit.sUnidadeMed, dFatorUMCol)
                        If lErro <> SUCESSO Then gError 137682
                    
                        dQuantidade = dQuantidade - (objOperacaoInsumos.dQuantidade * dFatorUMCol)
                        Exit For
                        
                    End If
                
                Next
                
            Next
            
            'se ainda tem quantidade ...
            If dQuantidade > 0 Then
                
                lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 134707
                
                GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_Produto_Col) = sProdutoMascarado
                GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
                
                If objProduto.iCompras = PRODUTO_COMPRAVEL Then
                    GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_OrigemProduto_Col) = INSUMO_COMPRADO
                Else
                    GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_OrigemProduto_Col) = INSUMO_PRODUZIDO
                End If
                
                GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_QuantidadeProduto_Col) = Formata_Estoque(dQuantidade)
                GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_UMProduto_Col) = objProdutoKit.sUnidadeMed
                
                Call Carrega_ComboVersoes(objProdutoKit.sProduto)
                
                'Verifica se tem versões (então é um Kit)...
                If VersaoKitComp.ListCount > 0 Then
                
                    'Verifica se no Kit está especificada a versão
                    If Len(objProdutoKit.sVersaoKitComp) > 0 Then
                        'seleciona a versão especificada
                        Call VersaoKitComp_Seleciona(objProdutoKit.sVersaoKitComp)
                    Else
                        'senão pega a versão padrão
                        Call VersaoKitComp_SelecionaPadrao(objProdutoKit.sProduto)
                    End If
                    
                    'e põe no grid
                    GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_VersaoKitComp_Col) = VersaoKitComp.Text
                
                End If
                
                If objProdutoKit.dPercentualPerda > 0 Then
                    GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_PerdaInsumos_Col) = Format(objProdutoKit.dPercentualPerda, "Percent")
                End If
                ComposicaoInsumos.ListIndex = objProdutoKit.iComposicao
                GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_ComposicaoInsumos_Col) = ComposicaoInsumos.Text
                If objProdutoKit.dCustoStandard > 0 Then
                    GridOperacaoInsumos.TextMatrix(iLinha + 1, iGrid_CustoStandardInsumos_Col) = Format(objProdutoKit.dCustoStandard, "Standard")
                End If
                
                'incrementa o contador de linhas
                iLinha = iLinha + 1
            
            End If
    
        End If
    
    Next
    
    objGridOperacaoInsumos.iLinhasExistentes = iLinha
    
    Exit Sub

Erro_BotaoTrazerInsumosKit_Click:

    Select Case gErr

        Case 137163
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTORAIZ_ROTFABR_NAO_PREENCHIDO", gErr)

        Case 137679
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSDEFABRICACAO_NAO_PREENCHIDO", gErr)

        Case 134702
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_ROTFABR_NAO_PREENCHIDA", gErr)
        
        Case 137680
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_ROTEIROSDEFABRICACAO_NAO_PREENCHIDA", gErr)
        
        Case 134701, 134703, 134705, 134707, 137681, 137682
            'erros tratados nas rotinas chamadas
        
        Case 134706
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                        
        Case 134704
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOKIT_INEXISTENTE", gErr, objKit.sVersao, objKit.sProdutoRaiz)

        Case 181410
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOKIT_DIFERENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174232)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVerRoteiros_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelProdutoRaiz_Click

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 138320

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(Versao.Text)) <> 0 Then
        objRoteirosDeFabricacao.sVersao = Versao.Text
    End If

    Call Chama_Tela("RoteirosDeFabricacaoLista", colSelecao, objRoteirosDeFabricacao, objEventoRoteiroDeFabricacao)

    Exit Sub

Erro_LabelProdutoRaiz_Click:

    Select Case gErr

        Case 138320

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174233)

    End Select

    Exit Sub

End Sub

Private Sub CustoStandard_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CustoStandard, iAlterado)
    
End Sub

Private Sub CustoStandard_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CustoStandard_Validate

    If Len(Trim(CustoStandard.ClipText)) <> 0 Then

       'Critica a CustoStandard
       lErro = Valor_Positivo_Critica(CustoStandard.Text)
       If lErro <> SUCESSO Then gError 137155

    End If

    Exit Sub

Erro_CustoStandard_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137155

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174234)

    End Select

    Exit Sub

End Sub

Private Sub CustoStandard_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_GotFocus()

    Call MaskEdBox_TrataGotFocus(Descricao, iAlterado)

End Sub

Private Sub GridOperacaoInsumos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridOperacaoInsumos)
        
End Sub

Private Sub GridOperacaoInsumos_LostFocus()

    Call Grid_Libera_Foco(objGridOperacaoInsumos)

End Sub

Private Sub IgnoraTaxa_Click()

Dim lErro As Long
    
On Error GoTo Erro_IgnoraTaxa_Click
    
    lErro = Habilita_TaxaDeProducao()
    If lErro <> SUCESSO Then gError 137537
   
    Exit Sub
    
Erro_IgnoraTaxa_Click:

    Select Case gErr
    
        Case 137537
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174235)

    End Select

    Exit Sub
   
End Sub

Private Sub LoteMaximo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LoteMaximo, iAlterado)
    
End Sub

Private Sub LoteMaximo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteMaximo_Validate

    If Len(Trim(LoteMaximo.ClipText)) <> 0 Then

       'Critica a LoteMaximo
       lErro = Valor_Positivo_Critica(LoteMaximo.Text)
       If lErro <> SUCESSO Then gError 137156
       
       LoteMaximo.Text = Formata_Estoque(LoteMaximo.Text)

    End If

    Exit Sub

Erro_LoteMaximo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137156

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174236)

    End Select

    Exit Sub

End Sub

Private Sub LoteMaximo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteMinimo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LoteMinimo, iAlterado)
    
End Sub

Private Sub LoteMinimo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteMinimo_Validate

    If Len(Trim(LoteMinimo.ClipText)) <> 0 Then

       'Critica a LoteMinimo
       lErro = Valor_Positivo_Critica(LoteMinimo.Text)
       If lErro <> SUCESSO Then gError 137157
       
       LoteMinimo.Text = Formata_Estoque(LoteMinimo.Text)

    End If

    Exit Sub

Erro_LoteMinimo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137157

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174237)

    End Select

    Exit Sub

End Sub

Private Sub LoteMinimo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LotePadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LotePadrao, iAlterado)
    
End Sub

Private Sub LotePadrao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LotePadrao_Validate

    If Len(Trim(LotePadrao.ClipText)) <> 0 Then

       'Critica a LotePadrao
       lErro = Valor_Positivo_Critica(LotePadrao.Text)
       If lErro <> SUCESSO Then gError 137158
       
       LotePadrao.Text = Formata_Estoque(LotePadrao.Text)

    End If

    Exit Sub

Erro_LotePadrao_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174238)

    End Select

    Exit Sub

End Sub

Private Sub LotePadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCentroDeTrabalho_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCentrodeTrabalho = obj1

    CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CodigoCTPadrao_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174239)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCompetencias_evSelecao(obj1 As Object)

Dim objCompetencias As New ClassCompetencias
Dim lErro As Long

On Error GoTo Erro_objEventoCompetencia_evSelecao

    Set objCompetencias = obj1

    CodigoCompetencia.Text = objCompetencias.sNomeReduzido
    
    Call CodigoCompetencia_Validate(bSGECancelDummy)

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_objEventoCompetencia_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174240)

    End Select

    Exit Sub

End Sub

Private Sub objEventoInsumos_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim objProdutoKit As New ClassProdutoKit

On Error GoTo Erro_objEventoInsumos_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 134708

    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True
    
    'Verifica se há algum produto repetido no grid
    For iLinha = 1 To objGridOperacaoInsumos.iLinhasExistentes
        
        If iLinha <> GridOperacaoInsumos.Row Then
                                                
            If GridOperacaoInsumos.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text Then
                Produto.PromptInclude = False
                Produto.Text = ""
                Produto.PromptInclude = True
                gError 134709
                
            End If
                
        End If
                       
    Next
        
    If Not (Me.ActiveControl Is Produto) Then
        
        GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_Produto_Col) = Produto.Text
    
        GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
        
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_OrigemProduto_Col) = INSUMO_COMPRADO
        Else
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_OrigemProduto_Col) = INSUMO_PRODUZIDO
        End If
        
        GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_UMProduto_Col) = objProduto.sSiglaUMEstoque
        
        Call Carrega_ComboVersoes(objProduto.sCodigo)
        
        If VersaoKitComp.ListCount > 0 Then
        
            Call VersaoKitComp_SelecionaPadrao(objProduto.sCodigo)
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col) = VersaoKitComp.Text
        
        End If
    
        Set objProdutoKit = New ClassProdutoKit
    
        objProdutoKit.sProdutoRaiz = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_Produto_Col)
        objProdutoKit.sVersao = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col)
        
        'Lê o Produto Raiz do Kit para pegar seus dados
        lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
        If lErro <> SUCESSO And lErro <> 34875 Then gError 134710
        
        'Se não encontrou é porque não existe esta Versão do Kit
        If lErro = SUCESSO Then
                        
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_PerdaInsumos_Col) = CStr(objProdutoKit.dPercentualPerda * 100)
            ComposicaoInsumos.ListIndex = objProdutoKit.iComposicao
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_ComposicaoInsumos_Col) = ComposicaoInsumos.Text
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_CustoStandardInsumos_Col) = CStr(objProdutoKit.dCustoStandard)
        
        Else
        
            Call ComposicaoInsumos_Seleciona
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_ComposicaoInsumos_Col) = ComposicaoInsumos.Text
        
        End If
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridOperacaoInsumos.iLinhasExistentes Then
            objGridOperacaoInsumos.iLinhasExistentes = objGridOperacaoInsumos.iLinhasExistentes + 1
        End If
        
    End If
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoInsumos_evSelecao:

    Select Case gErr

        Case 1347018, 134710
        
        Case 134709
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174241)

    End Select

    Exit Sub

End Sub

Private Sub objEventoKit_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objKit As ClassKit
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_objEventoKit_evSelecao

    Set objKit = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objKit.sProdutoRaiz, sProduto)
    If lErro <> SUCESSO Then gError 134711

    ProdutoRaiz.PromptInclude = False
    ProdutoRaiz.Text = sProduto
    ProdutoRaiz.PromptInclude = True

    Call ProdutoRaiz_Validate(bSGECancelDummy)

    Versao.Text = objKit.sVersao
    
    Set objProdutoKit = New ClassProdutoKit
    
    objProdutoKit.sProdutoRaiz = objKit.sProdutoRaiz
    objProdutoKit.sVersao = objKit.sVersao
    
    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
    If lErro <> SUCESSO Then gError 134712
            
    Quantidade.Text = Formata_Estoque(objProdutoKit.dQuantidade)
    
    lErro = CarregaComboUM(objProdutoKit.iClasseUM, objProdutoKit.sUnidadeMed)
    If lErro <> SUCESSO Then gError 134713
    
    UM.Text = objProdutoKit.sUnidadeMed
        
    PercentualPerda.Text = Format(objProdutoKit.dPercentualPerda, "Percent")
    
    Call Combo_Seleciona_ItemData(Composicao, objProdutoKit.iComposicao)

    CustoStandard = Format(objProdutoKit.dCustoStandard, "Standard")

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoKit_evSelecao:

    Select Case gErr

        Case 134711 To 134713

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174242)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoRaiz_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_objEventoProdutoRaiz_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 134714

    ProdutoRaiz.PromptInclude = False
    ProdutoRaiz.Text = sProduto
    ProdutoRaiz.PromptInclude = True
    
    Call ProdutoRaiz_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProdutoRaiz_evSelecao:

    Select Case gErr

        Case 134714
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174243)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174244)

    End Select

    Exit Sub
    
End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OrigemProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OrigemProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub OrigemProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub OrigemProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = OrigemProduto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PercentualPerda_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PercentualPerda, iAlterado)
    
End Sub

Private Sub PercentualPerda_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentualPerda_Validate

    If Len(Trim(PercentualPerda.ClipText)) <> 0 Then

        lErro = Porcentagem_Critica(PercentualPerda.Text)
        If lErro <> SUCESSO Then gError 137064

    End If

    Exit Sub

Erro_PercentualPerda_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137064
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174245)

    End Select

    Exit Sub

End Sub

Private Sub PercentualPerda_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoRaiz_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoRaiz, iAlterado)
    
End Sub

Private Sub Roteiro_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objOperacoes As New ClassOperacoes
Dim objNode As Node

On Error GoTo Erro_Roteiro_NodeClick

    Set objNode = Roteiro.SelectedItem

    Set objOperacoes = colComponentes.Item(objNode.Tag)
    
    lErro = Preenche_Operacoes(objOperacoes)
    If lErro <> SUCESSO Then gError 134716

    bOperacaoNova = False

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Roteiro_NodeClick:

    Select Case gErr

        Case 134716

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174246)

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
        
    End If

End Sub

Private Sub TempoDescarga_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoDescarga, iAlterado)
    
End Sub

Private Sub TempoDescarga_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoDescarga_Validate

    If Len(Trim(TempoDescarga.ClipText)) <> 0 Then

       'Critica a TempoDescarga
       lErro = Valor_Positivo_Critica(TempoDescarga.Text)
       If lErro <> SUCESSO Then gError 137159
       
       TempoDescarga.Text = Formata_Estoque(TempoDescarga.Text)

    End If

    Exit Sub

Erro_TempoDescarga_Validate:

    Cancel = True

    Select Case gErr
    
        Case 1367159

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174247)

    End Select

    Exit Sub

End Sub

Private Sub TempoDescarga_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoMovimentacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoMovimentacao, iAlterado)
    
End Sub

Private Sub TempoMovimentacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoMovimentacao_Validate

    If Len(Trim(TempoMovimentacao.ClipText)) <> 0 Then

       'Critica a TempoMovimentacao
       lErro = Valor_Positivo_Critica(TempoMovimentacao.Text)
       If lErro <> SUCESSO Then gError 137160
       
       TempoMovimentacao.Text = Formata_Estoque(TempoMovimentacao.Text)

    End If

    Exit Sub

Erro_TempoMovimentacao_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137160

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174248)

    End Select

    Exit Sub

End Sub

Private Sub TempoMovimentacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoOperacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoOperacao, iAlterado)
    
End Sub

Private Sub TempoOperacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoOperacao_Validate

    If Len(Trim(TempoOperacao.ClipText)) <> 0 Then
    
       'Critica a TempoOperacao
       lErro = Valor_Positivo_Critica(TempoOperacao.Text)
       If lErro <> SUCESSO Then gError 137161
       
       TempoOperacao.Text = Formata_Estoque(TempoOperacao.Text)

       If IgnoraTaxa.Value = vbChecked Then Call Preenche_LabelTaxaDeProducao

    End If

    Exit Sub

Erro_TempoOperacao_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137161

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174249)

    End Select

    Exit Sub

End Sub

Private Sub TempoOperacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoPreparacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoPreparacao, iAlterado)
    
End Sub

Private Sub TempoPreparacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoPreparacao_Validate

    If Len(Trim(TempoPreparacao.ClipText)) <> 0 Then

       'Critica a TempoPreparacao
       lErro = Valor_Positivo_Critica(TempoPreparacao.Text)
       If lErro <> SUCESSO Then gError 137162
       
       TempoPreparacao.Text = Formata_Estoque(TempoPreparacao.Text)

    End If

    Exit Sub

Erro_TempoPreparacao_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137162

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174250)

    End Select

    Exit Sub

End Sub

Private Sub TempoPreparacao_Change()

    iAlterado = REGISTRO_ALTERADO

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

    Set objEventoRoteiroDeFabricacao = Nothing
    Set objEventoCompetencias = Nothing
    Set objEventoCentroDeTrabalho = Nothing
    Set objEventoInsumos = Nothing
    Set objEventoKit = Nothing
    Set objEventoVersao = Nothing
    Set objEventoProdutoRaiz = Nothing
    Set objEventoCusteio = Nothing
    Set objEventoMaquina = Nothing
    
    Set objGridOperacaoInsumos = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174251)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    bOperacaoNova = True

    Set objEventoRoteiroDeFabricacao = New AdmEvento
    Set objEventoCompetencias = New AdmEvento
    Set objEventoCentroDeTrabalho = New AdmEvento
    Set objEventoInsumos = New AdmEvento
    Set objEventoKit = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Set objEventoProdutoRaiz = New AdmEvento
    Set objEventoCusteio = New AdmEvento
    Set objEventoMaquina = New AdmEvento
    
    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    DataUltModificacao.Caption = ""
    Autor.Caption = ""
            
    Call Composicao_Seleciona
    
    'Grid de OperacaoInsumos
    Set objGridOperacaoInsumos = New AdmGrid
    
    'tela em questão
    Set objGridOperacaoInsumos.objForm = Me
    
    lErro = Inicializa_GridOperacaoInsumos(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then gError 134717
    
    lErro = CarregaComboTipo(Tipo)
    If lErro <> SUCESSO Then gError 134718
    
    lErro = Preenche_Combo_UMTempo()
    If lErro <> SUCESSO Then gError 134719

    IgnoraTaxa.Value = vbUnchecked
    
    lErro = Habilita_TaxaDeProducao()
    If lErro <> SUCESSO Then gError 137538

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 134720
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoRaiz)
    If lErro <> SUCESSO Then gError 137130
    
    Quantidade.Format = FORMATO_ESTOQUE_KIT
    CustoStandard.Format = FORMATO_CUSTO
    
    If gobjEST.iTemRepeticoesOper = MARCADO Then
        Repeticao.Visible = True
        LabelRepeticao.Visible = True
    Else
        Repeticao.Visible = False
        LabelRepeticao.Visible = False
    End If
    
    Repeticao.PromptInclude = False
    Repeticao.Text = "1"
    Repeticao.PromptInclude = True

    lErro_Chama_Tela = SUCESSO

    iFrameAtual = 1
    iFrameAtualOper = 1
    iAlterado = 0
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 134717 To 134720, 137130, 137538
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174252)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Trata_Parametros

    If Not (objRoteirosDeFabricacao Is Nothing) Then

        lErro = Traz_RoteirosDeFabricacao_Tela(objRoteirosDeFabricacao)
        If lErro <> SUCESSO And lErro <> 134974 Then gError 134721

        If lErro <> SUCESSO Then
                
            If Len(objRoteirosDeFabricacao.sProdutoRaiz) > 0 Then
                    
                lErro = Mascara_RetornaProdutoEnxuto(objRoteirosDeFabricacao.sProdutoRaiz, sProduto)
                If lErro <> SUCESSO Then gError 134721
    
                'Coloca o ProdutoRaiz e a Versao na tela
                ProdutoRaiz.PromptInclude = False
                ProdutoRaiz.Text = sProduto
                ProdutoRaiz.PromptInclude = True
                
                If Len(objRoteirosDeFabricacao.sVersao) > 0 Then
                
                    Versao.Text = objRoteirosDeFabricacao.sVersao
                
                End If
                
                Call ProdutoRaiz_Validate(bSGECancelDummy)
                    
            End If
    
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 134721

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174253)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objOperacoes As New ClassOperacoes

On Error GoTo Erro_Move_Tela_Memoria

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 137132
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
    
    End If

    objRoteirosDeFabricacao.sVersao = Versao.Text
    objRoteirosDeFabricacao.sDescricao = Descricao.Text
    objRoteirosDeFabricacao.iVersaoFormPreco = VersaoFormPreco.Value
    
    objRoteirosDeFabricacao.dtDataCriacao = StrParaDate(DataCriacao.Text)

    objRoteirosDeFabricacao.dtDataUltModificacao = gdtDataAtual
    
    If Len(Trim(Autor.Caption)) = 0 Then
       objRoteirosDeFabricacao.sAutor = gsUsuario
    Else
       objRoteirosDeFabricacao.sAutor = Autor.Caption
    End If
    
    objRoteirosDeFabricacao.dQuantidade = StrParaDbl(Quantidade.Text)
    objRoteirosDeFabricacao.sUM = UM.Text
    
    objRoteirosDeFabricacao.iNumMaxMaqPorOper = StrParaInt(NumMaxMaqPorOper.Text)
    
    'preenche objeto para geração/alteração do Kit
    objRoteirosDeFabricacao.iComposicao = Composicao.ItemData(Composicao.ListIndex)
    objRoteirosDeFabricacao.dPercentualPerda = StrParaDbl(Val(PercentualPerda.Text) / 100)
    objRoteirosDeFabricacao.dCustoStandard = StrParaDbl(CustoStandard.Text)
    
    'preenche a coleção das Operacoes
    For Each objOperacoes In colComponentes
    
       objRoteirosDeFabricacao.colOperacoes.Add objOperacoes
       
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 137132

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174254)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RoteirosDeFabricacao"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134722

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "ProdutoRaiz", objRoteirosDeFabricacao.sProdutoRaiz, STRING_PRODUTO, "ProdutoRaiz"
    colCampoValor.Add "Versao", objRoteirosDeFabricacao.sVersao, 10, "Versao"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 134722

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174255)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao

On Error GoTo Erro_Tela_Preenche

    objRoteirosDeFabricacao.sProdutoRaiz = colCampoValor.Item("ProdutoRaiz").vValor
    objRoteirosDeFabricacao.sVersao = colCampoValor.Item("Versao").vValor

    If Len(Trim(objRoteirosDeFabricacao.sProdutoRaiz)) > 0 And Len(Trim(objRoteirosDeFabricacao.sVersao)) > 0 Then
        lErro = Traz_RoteirosDeFabricacao_Tela(objRoteirosDeFabricacao)
        If lErro <> SUCESSO Then gError 134723
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 134723

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174256)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se ProdutoRaiz está preenchido
    If Len(Trim(ProdutoRaiz.Text)) = 0 Then gError 134724
    
    'Verifica se Versao está preenchida
    If Len(Trim(Versao.Text)) = 0 Then gError 134725

    'Verifica se a Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then gError 134726
    
    'Verifica se a U.M. está preenchida
    If Len(Trim(UM.Text)) = 0 Then gError 134727
    
    'Verifica se a Data de Criação está preenchida
    If Len(Trim(DataCriacao.ClipText)) = 0 Then gError 137127
    
    'Verifica se a Composicao está preenchida
    If Len(Trim(Composicao.Text)) = 0 Then gError 134728
    
    'Verifica se existe pelo menos uma Operação cadastrada
    If colComponentes.Count = 0 Then gError 134729

    'Preenche o objRoteirosDeFabricacao
    lErro = Move_Tela_Memoria(objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134730
        
    lErro = Trata_Alteracao(objRoteirosDeFabricacao, objRoteirosDeFabricacao.sProdutoRaiz, objRoteirosDeFabricacao.sVersao)
    If lErro <> SUCESSO Then gError 137683
    
    'Grava o RoteirosDeFabricacao no Banco de Dados
    lErro = CF("RoteirosDeFabricacao_Grava", objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134731
        
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134724
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTORAIZ_ROTFABR_NAO_PREENCHIDO", gErr)

        Case 134725
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSDEFABRICACAO_NAO_PREENCHIDO", gErr)

        Case 134726
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_ROTFABR_NAO_PREENCHIDA", gErr)
        
        Case 134727
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_ROTEIROSDEFABRICACAO_NAO_PREENCHIDA", gErr)
        
        Case 134728
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPOSICAO_ROTFABRICACAO_NAOPREENCHIDA", gErr)
        
        Case 134729
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACOES_ROTEIROSDEFABRICACAO_NAO_PREENCHIDA", gErr)
        
        Case 134730, 134731, 137683
        
        Case 137127
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACRIACAO_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174257)

    End Select

    Exit Function

End Function

Function Limpa_Tela_RoteirosDeFabricacao() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RoteirosDeFabricacao
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    DescricaoProd.Caption = ""
    
    UM.Clear
    
    Call Composicao_Seleciona
    
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 137539
    
    lErro = Limpa_Arvore_Roteiro()
    If lErro <> SUCESSO Then gError 137540
    
    Call Grid_Limpa(objGridOperacaoInsumos)

    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    VersaoFormPreco.Value = vbUnchecked
    
    DataUltModificacao.Caption = ""
    Autor.Caption = ""
    
    IgnoraTaxa.Value = vbUnchecked
    
    lErro = Habilita_TaxaDeProducao()
    If lErro <> SUCESSO Then gError 137541
    
    iAlterado = 0

    Limpa_Tela_RoteirosDeFabricacao = SUCESSO

    Exit Function

Erro_Limpa_Tela_RoteirosDeFabricacao:

    Limpa_Tela_RoteirosDeFabricacao = gErr

    Select Case gErr
    
        Case 137539 To 137541
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174258)

    End Select

    Exit Function

End Function

Function Traz_RoteirosDeFabricacao_Tela(objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sCodProduto As String
Dim sProdutoMascarado As String
Dim objProdutoKit As ClassProdutoKit
Dim objKit As New ClassKit

On Error GoTo Erro_Traz_RoteirosDeFabricacao_Tela

    'Lê o RoteirosDeFabricacao que está sendo Passado
    lErro = CF("RoteirosDeFabricacao_Le", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 134733
    
    If lErro <> SUCESSO Then gError 134974

    'Limpa Tela
    Call Limpa_Tela_RoteirosDeFabricacao
    
    sCodProduto = objRoteirosDeFabricacao.sProdutoRaiz
    
    lErro = Mascara_RetornaProdutoEnxuto(sCodProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 137131

    ProdutoRaiz.PromptInclude = False
    ProdutoRaiz.Text = sProdutoMascarado
    ProdutoRaiz.PromptInclude = True
        
    Versao.Text = objRoteirosDeFabricacao.sVersao
    Descricao.Text = objRoteirosDeFabricacao.sDescricao

    If objRoteirosDeFabricacao.dtDataCriacao <> DATA_NULA Then
        DataCriacao.PromptInclude = False
        DataCriacao.Text = Format(objRoteirosDeFabricacao.dtDataCriacao, "dd/mm/yy")
        DataCriacao.PromptInclude = True
    End If

    If objRoteirosDeFabricacao.dtDataUltModificacao <> DATA_NULA Then
        DataUltModificacao.Caption = Format(objRoteirosDeFabricacao.dtDataUltModificacao, "dd/mm/yyyy")
    End If
    
    If Len(objRoteirosDeFabricacao.sAutor) <> 0 Then
        Autor.Caption = objRoteirosDeFabricacao.sAutor
    End If
    
    If objRoteirosDeFabricacao.dQuantidade <> 0 Then Quantidade.Text = CStr(objRoteirosDeFabricacao.dQuantidade)
    UM.Text = objRoteirosDeFabricacao.sUM
    
    If objRoteirosDeFabricacao.iNumMaxMaqPorOper <> 0 Then
        NumMaxMaqPorOper.PromptInclude = False
        NumMaxMaqPorOper.Text = objRoteirosDeFabricacao.iNumMaxMaqPorOper
        NumMaxMaqPorOper.PromptInclude = True
    End If
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objRoteirosDeFabricacao.sProdutoRaiz
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 134735
    
    DescricaoProd.Caption = objProduto.sDescricao

    lErro = CarregaComboUM(objProduto.iClasseUM, objRoteirosDeFabricacao.sUM)
    If lErro <> SUCESSO Then gError 134736

    Set objProdutoKit = New ClassProdutoKit
    
    objProdutoKit.sProdutoRaiz = objRoteirosDeFabricacao.sProdutoRaiz
    objProdutoKit.sVersao = objRoteirosDeFabricacao.sVersao
    
    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
    If lErro <> SUCESSO Then gError 134737
    
    If lErro = SUCESSO Then
                    
        If objProdutoKit.dPercentualPerda <> 0 Then PercentualPerda.Text = CStr(objProdutoKit.dPercentualPerda * 100)
        
        Composicao.ListIndex = objProdutoKit.iComposicao
            
        If objProdutoKit.dCustoStandard <> 0 Then CustoStandard = CStr(objProdutoKit.dCustoStandard)
        
    End If
    
    objKit.sVersao = objRoteirosDeFabricacao.sVersao
    objKit.sProdutoRaiz = objRoteirosDeFabricacao.sProdutoRaiz
    
    lErro = CF("Kit_Le", objKit)
    If lErro <> SUCESSO And lErro <> 21826 Then gError 138535
    
    If lErro = SUCESSO Then
    
        VersaoFormPreco.Value = objKit.iVersaoFormPreco
    
    End If

    lErro = Carrega_Arvore(objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134738

    iAlterado = 0

    Traz_RoteirosDeFabricacao_Tela = SUCESSO

    Exit Function

Erro_Traz_RoteirosDeFabricacao_Tela:

    Traz_RoteirosDeFabricacao_Tela = gErr

    Select Case gErr

        Case 134733, 134738, 137131, 138535, 134736, 134735, 134737
            'erros tratados nas rotinas chamadas
        
        Case 134974
            'sem dados -> erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174259)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 134739

    'Limpa Tela
    Call Limpa_Tela_RoteirosDeFabricacao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 134739

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174260)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174261)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134740

    Call Limpa_Tela_RoteirosDeFabricacao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 134740

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174262)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sTextoExclusao As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(ProdutoRaiz.Text)) = 0 Then gError 134741
    If Len(Trim(Versao.Text)) = 0 Then gError 134742

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 137132
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
    
    End If

    objRoteirosDeFabricacao.sVersao = Versao.Text

    sTextoExclusao = objRoteirosDeFabricacao.sProdutoRaiz & SEPARADOR & objRoteirosDeFabricacao.sVersao

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ROTEIROSDEFABRICACAO", sTextoExclusao)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a requisição de consumo
    lErro = CF("RoteirosDeFabricacao_Exclui", objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134743

    'Limpa Tela
    Call Limpa_Tela_RoteirosDeFabricacao

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134741
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTORAIZ_ROTFABRICACAO_NAO_PREENCHIDO", gErr)

        Case 134742
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSDEFABRICACAO_NAO_PREENCHIDO", gErr)

        Case 134743

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174263)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoRaiz_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim sProdutoMascarado As String

On Error GoTo Erro_Produto_Validate
   
    If Len(Trim(ProdutoRaiz.ClipText)) = 0 Then
        DescricaoProd.Caption = ""
        Exit Sub
    End If

    sProduto = ProdutoRaiz.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 134744

    'se o produto não estiver cadastrado ==> erro
    If lErro = 25041 Then gError 134745

    'se o produto for gerencial, não pode fazer parte de um kit
    If objProduto.iGerencial = GERENCIAL Then gError 134746
    
    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 48743
    
    ProdutoRaiz.PromptInclude = False
    ProdutoRaiz.Text = sProdutoMascarado
    ProdutoRaiz.PromptInclude = True
    
    lErro = CarregaComboUM(objProduto.iClasseUM, objProduto.sSiglaUMEstoque)
    If lErro <> SUCESSO Then gError 134747

    UM.Text = objProduto.sSiglaUMEstoque

    DescricaoProd.Caption = objProduto.sDescricao
            
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 134744, 134747, 48743
            'erros tratados nas rotinas chamadas
            
        Case 134745
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
        
        Case 134746
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174264)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoRaiz_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Versao_GotFocus()

    Call MaskEdBox_TrataGotFocus(Versao, iAlterado)

End Sub

Private Sub Versao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Descricao_Change()

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
        If lErro <> SUCESSO Then gError 134750

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_DownClick:

    Select Case gErr

        Case 134750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174265)

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
        If lErro <> SUCESSO Then gError 134751

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_UpClick:

    Select Case gErr

        Case 134751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174266)

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
        If lErro <> SUCESSO Then gError 134752

    End If

    Exit Sub

Erro_DataCriacao_Validate:

    Cancel = True

    Select Case gErr

        Case 134752

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174267)

    End Select

    Exit Sub

End Sub

Private Sub DataCriacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Veifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

       'Critica a Quantidade
       lErro = Valor_Positivo_Critica(Quantidade.Text)
       If lErro <> SUCESSO Then gError 134753
       
       Quantidade.Text = Formata_Estoque(Quantidade.Text)
       
       If IgnoraTaxa.Value = vbChecked Then Call Preenche_LabelTaxaDeProducao

    End If

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 134753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174268)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UM_Validate

    'Veifica se UM está preenchida
    If Len(Trim(UM.Text)) <> 0 Then

       If IgnoraTaxa.Value = vbChecked Then Call Preenche_LabelTaxaDeProducao

    End If

    Exit Sub

Erro_UM_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174269)

    End Select

    Exit Sub

End Sub

Private Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoRoteiroDeFabricacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao

On Error GoTo Erro_objEventoRoteiroDeFabricacao_evSelecao

    Set objRoteirosDeFabricacao = obj1

    'Mostra os dados do RoteirosDeFabricacao na tela
    lErro = Traz_RoteirosDeFabricacao_Tela(objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134754

    Me.Show

    Exit Sub

Erro_objEventoRoteiroDeFabricacao_evSelecao:

    Select Case gErr

        Case 134754

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174270)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134755

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoRaiz)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 134755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174271)

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
    If lErro <> SUCESSO Then gError 134756

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
        If Len(Trim(Versao.ClipText)) > 0 Then objKit.sVersao = Versao.Text
            
        colSelecao.Add sProdutoFormatado
        
        Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
    Else
         gError 134757
         
    End If

    Exit Sub

Erro_LabelVersao_Click:

    Select Case gErr

        Case 134756
        
        Case 134757
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174272)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridOperacaoInsumos(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("C/P")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Perda")
    objGrid.colColuna.Add ("Compos.")
    objGrid.colColuna.Add ("C.Stand.")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (DescricaoProduto.Name)
    objGrid.colCampo.Add (OrigemProduto.Name)
    objGrid.colCampo.Add (UMProduto.Name)
    objGrid.colCampo.Add (QuantidadeProduto.Name)
    objGrid.colCampo.Add (VersaoKitComp.Name)
    objGrid.colCampo.Add (PerdaInsumos.Name)
    objGrid.colCampo.Add (ComposicaoInsumos.Name)
    objGrid.colCampo.Add (CustoStandardInsumos.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescricaoProduto_Col = 2
    iGrid_OrigemProduto_Col = 3
    iGrid_UMProduto_Col = 4
    iGrid_QuantidadeProduto_Col = 5
    iGrid_VersaoKitComp_Col = 6
    iGrid_PerdaInsumos_Col = 7
    iGrid_ComposicaoInsumos_Col = 8
    iGrid_CustoStandardInsumos_Col = 9

    objGrid.objGrid = GridOperacaoInsumos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1 'NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridOperacaoInsumos.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridOperacaoInsumos = SUCESSO

End Function

Private Sub GridOperacaoInsumos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOperacaoInsumos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOperacaoInsumos, iAlterado)
    End If

End Sub

Private Sub GridOperacaoInsumos_GotFocus()
    
    Call Grid_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub GridOperacaoInsumos_EnterCell()

    Call Grid_Entrada_Celula(objGridOperacaoInsumos, iAlterado)

End Sub

Private Sub GridOperacaoInsumos_LeaveCell()
    
    Call Saida_Celula(objGridOperacaoInsumos)

End Sub

Private Sub GridOperacaoInsumos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOperacaoInsumos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOperacaoInsumos, iAlterado)
    End If

End Sub

Private Sub GridOperacaoInsumos_RowColChange()

    Call Grid_RowColChange(objGridOperacaoInsumos)

End Sub

Private Sub GridOperacaoInsumos_Scroll()

    Call Grid_Scroll(objGridOperacaoInsumos)

End Sub

Private Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub QuantidadeProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub QuantidadeProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = QuantidadeProduto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub UMProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub UMProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = UMProduto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'OperacaoInsumos
        If objGridInt.objGrid.Name = GridOperacaoInsumos.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Produto_Col

                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 134758

                Case iGrid_DescricaoProduto_Col

                    lErro = Saida_Celula_DescricaoProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 134759

                Case iGrid_QuantidadeProduto_Col

                    lErro = Saida_Celula_QuantidadeProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 134760

                Case iGrid_UMProduto_Col

                    lErro = Saida_Celula_UMProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 134761

                Case iGrid_VersaoKitComp_Col

                    lErro = Saida_Celula_VersaoKitComp(objGridInt)
                    If lErro <> SUCESSO Then gError 134762
                
                Case iGrid_PerdaInsumos_Col

                    lErro = Saida_Celula_PerdaInsumos(objGridInt)
                    If lErro <> SUCESSO Then gError 134763
                
                Case iGrid_ComposicaoInsumos_Col

                    lErro = Saida_Celula_ComposicaoInsumos(objGridInt)
                    If lErro <> SUCESSO Then gError 134764
                
                Case iGrid_CustoStandardInsumos_Col

                    lErro = Saida_Celula_CustoStandardInsumos(objGridInt)
                    If lErro <> SUCESSO Then gError 134765

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 134766

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 134758 To 134765

        Case 134766
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174273)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sProduto As String
Dim objProdutos As ClassProduto
Dim objClasseUM As ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Guardo o valor do Codigo do Produto
    sProduto = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_Produto_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134767
    
    If objControl.Name = "Produto" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True
        
        End If
        
    ElseIf objControl.Name = "DescricaoProduto" Then

        objControl.Enabled = False
        
    ElseIf objControl.Name = "OrigemProduto" Then

        objControl.Enabled = False
        
    ElseIf objControl.Name = "QuantidadeProduto" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "UMProduto" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objControl.Enabled = True

            Set objProdutos = New ClassProduto

            objProdutos.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProdutos)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 134768

            Set objClasseUM = New ClassClasseUM
            
            objClasseUM.iClasse = objProdutos.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 134769

            'Se tem algum valor para UMProduto do Grid
            If Len(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_UMProduto_Col)) > 0 Then
                'Guardo o valor da UMProduto da Linha
                sUnidadeMed = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_UMProduto_Col)
            Else
                'Senão coloco o do Produto em estoque
                sUnidadeMed = objProdutos.sSiglaUMEstoque
            End If
            
            'Limpar as Unidades utilizadas anteriormente
            UMProduto.Clear

            For Each objUnidadeDeMedida In colSiglas
                UMProduto.AddItem objUnidadeDeMedida.sSigla
            Next

            'UMProduto.AddItem ""

            'Tento selecionar na Combo a Unidade anterior
            If UMProduto.ListCount <> 0 Then

                For iIndice = 0 To UMProduto.ListCount - 1

                    If UMProduto.List(iIndice) = sUnidadeMed Then
                        UMProduto.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If
            
        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "VersaoKitComp" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True
            Call Carrega_ComboVersoes(sProdutoFormatado)
            Call VersaoKitComp_Seleciona(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col))

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "PerdaInsumos" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "ComposicaoInsumos" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "CustoStandardInsumos" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
    
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 134767 To 134769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174274)

    End Select

    Exit Sub

End Sub

Private Function CarregaComboTipo(objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CarregaComboTipo

    
    objCombo.AddItem ITEM_TIPO_TAXAPRODUCAO_VARIAVEL & SEPARADOR & STRING_ITEM_TIPO_TAXAPRODUCAO_VARIAVEL
    objCombo.ItemData(objCombo.NewIndex) = ITEM_TIPO_TAXAPRODUCAO_VARIAVEL
    
    objCombo.AddItem ITEM_TIPO_TAXAPRODUCAO_FIXO & SEPARADOR & STRING_ITEM_TIPO_TAXAPRODUCAO_FIXO
    objCombo.ItemData(objCombo.NewIndex) = ITEM_TIPO_TAXAPRODUCAO_FIXO
    
    CarregaComboTipo = SUCESSO

    Exit Function

Erro_CarregaComboTipo:

    CarregaComboTipo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174275)

    End Select

    Exit Function

End Function
Private Function Preenche_Combo_UMTempo()

Dim objClasseUM As ClassClasseUM
Dim colSiglas As New Collection
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim lErro As Long

On Error GoTo Erro_Preenche_Combo_UMTempo

    Set objClasseUM = New ClassClasseUM
    
    objClasseUM.iClasse = gobjEST.iClasseUMTempo

    'Preenche a List da Combo UnidadeMed com as UM's de Tempo
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO Then gError 134770

    'Se tem algum valor para UMTempo
    If Len(UMTempo.Text) > 0 Then
        'Guardo o valor da UMTempo da Linha
        sUnidadeMed = UMTempo.Text
    Else
        'Senão coloco o Padrão UMTempo
        Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)
        'sUnidadeMed = TAXA_CONSUMO_TEMPO_PADRAO
    End If
    
    'Limpar as Unidades utilizadas anteriormente
    UMTempo.Clear

    For Each objUnidadeDeMedida In colSiglas
        UMTempo.AddItem objUnidadeDeMedida.sSigla
    Next

    UMTempo.AddItem ""

    'Tento selecionar na Combo a Unidade anterior
    If UMTempo.ListCount <> 0 Then

        For iIndice = 0 To UMTempo.ListCount - 1

            If UMTempo.List(iIndice) = sUnidadeMed Then
                UMTempo.ListIndex = iIndice
                Exit For
            End If
        Next
    End If

    Exit Function

Erro_Preenche_Combo_UMTempo:

    Select Case gErr
    
        Case 134770

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174276)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodProduto As String
Dim iLinha As Integer
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
                
    sCodProduto = Produto.Text
        
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134771
    
    'Se o campo foi preenchido
    If Len(sProdutoFormatado) > 0 Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134772
        
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridOperacaoInsumos.Row Then
                                                    
                If GridOperacaoInsumos.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado Then
                    Produto.PromptInclude = False
                    Produto.Text = ""
                    Produto.PromptInclude = True
                    gError 134773
                    
                End If
                    
            End If
                           
        Next
        
        Set objProdutos = New ClassProduto

        objProdutos.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134774
        
        'Verifica se o produto pode compor um Kit
        If objProdutos.iAtivo <> 0 And objProdutos.iGerencial <> 0 And _
            objProdutos.iKitBasico <> 1 And objProdutos.iKitInt <> 1 Then gError 134775
                
        GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_DescricaoProduto_Col) = objProdutos.sDescricao
        
        If objProdutos.iCompras = PRODUTO_COMPRAVEL Then
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_OrigemProduto_Col) = INSUMO_COMPRADO
        Else
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_OrigemProduto_Col) = INSUMO_PRODUZIDO
        End If

        If Len(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_UMProduto_Col)) = 0 Then
            GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_UMProduto_Col) = objProdutos.sSiglaUMEstoque
        End If

        Call Carrega_ComboVersoes(objProdutos.sCodigo)

        If VersaoKitComp.ListCount > 0 Then

            If Len(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col)) = 0 Then
                Call VersaoKitComp_SelecionaPadrao(objProdutos.sCodigo)
                GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col) = VersaoKitComp.Text
            End If

        End If

        Set objProdutoKit = New ClassProdutoKit

        objProdutoKit.sProdutoRaiz = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_Produto_Col)
        objProdutoKit.sVersao = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col)

        'Lê o Produto Raiz do Kit para pegar seus dados
        lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
        If lErro <> SUCESSO And lErro <> 34875 Then gError 134776

        'Se não encontrou é porque não existe esta Versão do Kit
        If lErro = SUCESSO Then

            If Len(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_PerdaInsumos_Col)) = 0 Then
                GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_PerdaInsumos_Col) = CStr(objProdutoKit.dPercentualPerda * 100)
            End If
            If Len(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_ComposicaoInsumos_Col)) = 0 Then
                Call Combo_Seleciona_ItemData(ComposicaoInsumos, objProdutoKit.iComposicao)
                GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_ComposicaoInsumos_Col) = ComposicaoInsumos.Text
            End If
            If Len(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_CustoStandardInsumos_Col)) = 0 Then
                GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_CustoStandardInsumos_Col) = CStr(objProdutoKit.dCustoStandard)
            End If

        Else

            If Len(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_ComposicaoInsumos_Col)) = 0 Then
                Call ComposicaoInsumos_Seleciona
                GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_ComposicaoInsumos_Col) = ComposicaoInsumos.Text
            End If

        End If

        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134777

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 134773
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134775
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_COMPOR_KIT", gErr, sProdutoMascarado, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134771, 134772, 134774, 134776, 134777
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174277)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoProduto

    Set objGridInt.objControle = DescricaoProduto
    
    'Se o campo foi preenchido
    If Len(Trim(DescricaoProduto.Text)) > 0 Then
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134778

    Saida_Celula_DescricaoProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoProduto:

    Saida_Celula_DescricaoProduto = gErr

    Select Case gErr

        Case 134778
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174278)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantidadeProduto

    Set objGridInt.objControle = QuantidadeProduto
    
    'Se o campo foi preenchido
    If Len(Trim(QuantidadeProduto.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(QuantidadeProduto.Text)
        If lErro <> SUCESSO Then gError 137192
        
        QuantidadeProduto.Text = Formata_Estoque(QuantidadeProduto.Text)
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134779

    Saida_Celula_QuantidadeProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeProduto:

    Saida_Celula_QuantidadeProduto = gErr

    Select Case gErr

        Case 134779, 137192
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174279)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UMProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UMProduto

    Set objGridInt.objControle = UMProduto
    
    'Se o campo foi preenchido
    If Len(Trim(UMProduto.Text)) > 0 Then
    
        GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_UMProduto_Col) = UMProduto.Text
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134780

    Saida_Celula_UMProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_UMProduto:

    Saida_Celula_UMProduto = gErr

    Select Case gErr

        Case 134780
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174280)

    End Select

    Exit Function

End Function

Private Function Inicializa_Padrao() As Long

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objCTMaquinas As New ClassCTMaquinas
Dim objTaxaDeProducao As ClassTaxaDeProducao
Dim objTaxaDeProducao_Melhor As ClassTaxaDeProducao
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Inicializa_Padrao

    Set objTaxaDeProducao_Melhor = New ClassTaxaDeProducao
    
    objTaxaDeProducao_Melhor.iTipo = 1
    objTaxaDeProducao_Melhor.dQuantidade = 1
    objTaxaDeProducao_Melhor.dTempoOperacao = 1
    
    Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)
    objTaxaDeProducao_Melhor.sUMTempo = sUnidadeMed
    'objTaxaDeProducao_Melhor.sUMTempo = TAXA_CONSUMO_TEMPO_PADRAO

    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.sNomeReduzido = CodigoCompetencia.Text
    
    'Verifica a Competencia no BD a partir do NomeReduzido
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 134780

    sProduto = ProdutoRaiz.Text
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134781
        
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Verifica se o CodigoCTPadrao existe, lendo no BD a partir do NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 134782

        If lErro = SUCESSO Then
        
            'Lê o CTMaquinas que está sendo Passado
            lErro = CF("CentrodeTrabalho_Le_CTMaquinas", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134458 Then gError 134783
            
            'Percorre os dados da coleção de CTMáquinas
            For Each objCTMaquinas In objCentrodeTrabalho.colMaquinas
                
                Set objTaxaDeProducao = New ClassTaxaDeProducao
                
                objTaxaDeProducao.lNumIntDocCompet = objCompetencias.lNumIntDoc
                objTaxaDeProducao.sProduto = sProdutoFormatado
                objTaxaDeProducao.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq
        
                lErro = CF("TaxaDeProducao_Seleciona", objTaxaDeProducao)
                If lErro <> SUCESSO And lErro <> 134600 Then gError 134784
                
                'encontrou taxa ...
                If lErro = SUCESSO Then
                
                    If objTaxaDeProducao.lNumIntDoc <> objTaxaDeProducao_Melhor.lNumIntDoc Then
                
                        lErro = Compara_TaxaDeProducao(objTaxaDeProducao, objTaxaDeProducao_Melhor)
                        If lErro <> SUCESSO Then
                        
                            Set objTaxaDeProducao_Melhor = New ClassTaxaDeProducao
                        
                            objTaxaDeProducao_Melhor.iTipo = 1
                            objTaxaDeProducao_Melhor.dTempoOperacao = 1
                            'objTaxaDeProducao_Melhor.sUMTempo = TAXA_CONSUMO_TEMPO_PADRAO
                            Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)
                            objTaxaDeProducao_Melhor.sUMTempo = sUnidadeMed
                                                
                            Exit For
                            
                        End If
                    
                    End If
                    
                End If
        
            Next
        
        End If
    
    Else
    
        Set objTaxaDeProducao = New ClassTaxaDeProducao
        
        objTaxaDeProducao.lNumIntDocCompet = objCompetencias.lNumIntDoc
        objTaxaDeProducao.sProduto = sProdutoFormatado

        lErro = CF("TaxaDeProducao_Seleciona", objTaxaDeProducao)
        If lErro <> SUCESSO And lErro <> 134600 Then gError 134785

        'encontrou taxa ...
        If lErro = SUCESSO Then
        
            Set objTaxaDeProducao_Melhor = objTaxaDeProducao
    
        End If
    
    End If
    
    Call Combo_Seleciona_ItemData(Tipo, objTaxaDeProducao_Melhor.iTipo)
    
    If objTaxaDeProducao_Melhor.iTipo = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
        dFator = 1
    Else
        dFator = StrParaDbl(Quantidade.Text) / objTaxaDeProducao_Melhor.dQuantidade
    End If
    
    If objTaxaDeProducao_Melhor.dTempoOperacao <> 0 Then TempoOperacao.Text = Formata_Estoque(objTaxaDeProducao_Melhor.dTempoOperacao * dFator)
    sUnidadeMed = objTaxaDeProducao_Melhor.sUMTempo
    If objTaxaDeProducao_Melhor.dTempoPreparacao <> 0 Then TempoPreparacao.Text = Formata_Estoque(objTaxaDeProducao_Melhor.dTempoPreparacao)
    If objTaxaDeProducao_Melhor.dTempoMovimentacao <> 0 Then TempoMovimentacao.Text = Formata_Estoque(objTaxaDeProducao_Melhor.dTempoMovimentacao)
    If objTaxaDeProducao_Melhor.dTempoDescarga <> 0 Then TempoDescarga.Text = Formata_Estoque(objTaxaDeProducao_Melhor.dTempoDescarga)
    If objTaxaDeProducao_Melhor.dLoteMax <> 0 Then LoteMaximo.Text = Formata_Estoque(objTaxaDeProducao_Melhor.dLoteMax)
    If objTaxaDeProducao_Melhor.dLoteMin <> 0 Then LoteMinimo.Text = Formata_Estoque(objTaxaDeProducao_Melhor.dLoteMin)
    If objTaxaDeProducao_Melhor.dLotePadrao <> 0 Then LotePadrao.Text = Formata_Estoque(objTaxaDeProducao_Melhor.dLotePadrao)
    
    If objTaxaDeProducao_Melhor.lNumIntDocMaq <> 0 Then
        
        objMaquinas.lNumIntDoc = objTaxaDeProducao_Melhor.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 180348
        
        Maquina.Text = objMaquinas.sNomeReduzido
        DescMaquina.Caption = objMaquinas.sDescricao
        
    End If
    
    'Tento selecionar na Combo a Unidade Padrão
    If UMTempo.ListCount <> 0 Then

        For iIndice = 0 To UMTempo.ListCount - 1

            If UMTempo.List(iIndice) = sUnidadeMed Then
                UMTempo.ListIndex = iIndice
                Exit For
            End If
            
        Next
        
    End If
    
    Inicializa_Padrao = SUCESSO
    
    Exit Function
        
Erro_Inicializa_Padrao:

    Inicializa_Padrao = gErr
    
    Select Case gErr
    
        Case 134780 To 134785, 180348
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174281)

    End Select

    Exit Function
    
End Function

Private Function CarregaComboUM(ByVal iClasseUM As Integer, ByVal sUM As String) As Long

Dim lErro As Long
Dim objClasseUM As ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_CarregaComboUM

    Set objClasseUM = New ClassClasseUM
    
    objClasseUM.iClasse = iClasseUM
    
    'Preenche a List da Combo UnidadeMed com as UM's da Competencia
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO Then gError 134786

    'Se tem algum valor para UM na Tela
    If Len(UM.Text) > 0 Then
        'Guardo o valor da UM da Tela
        sUnidadeMed = UM.Text
    Else
        'Senão coloco a do ProdutoRaiz no Kit
        sUnidadeMed = sUM
    End If
    
    'Limpar as Unidades utilizadas anteriormente
    UM.Clear

    For Each objUnidadeDeMedida In colSiglas
        UM.AddItem objUnidadeDeMedida.sSigla
    Next

    'UM.AddItem ""

    'Tento selecionar na Combo a Unidade anterior
    If UM.ListCount <> 0 Then

        For iIndice = 0 To UM.ListCount - 1

            If UM.List(iIndice) = sUnidadeMed Then
                UM.ListIndex = iIndice
                Exit For
            End If
        Next
    End If
    
    CarregaComboUM = SUCESSO
    
    Exit Function

Erro_CarregaComboUM:

    CarregaComboUM = gErr

    Select Case gErr

        Case 134786
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174282)

    End Select

    Exit Function

End Function

Private Function Preenche_GridOperacaoInsumos(objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_Preenche_GridOperacaoInsumos
    
    Call Grid_Limpa(objGridOperacaoInsumos)
    
    'Exibe os dados da coleção na tela
    For iIndice = 1 To objOperacoes.colOperacaoInsumos.Count
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objOperacoes.colOperacaoInsumos.Item(iIndice).sProduto
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134787
        
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134788
                
        'Insere no GridOperacaoInsumos
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_DescricaoProduto_Col) = objProdutos.sDescricao
        
        If objProdutos.iCompras = PRODUTO_COMPRAVEL Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_OrigemProduto_Col) = INSUMO_COMPRADO
        Else
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_OrigemProduto_Col) = INSUMO_PRODUZIDO
        End If
        
        If objOperacoes.colOperacaoInsumos.Item(iIndice).dQuantidade > 0 Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_QuantidadeProduto_Col) = Formata_Estoque(objOperacoes.colOperacaoInsumos.Item(iIndice).dQuantidade)
        End If
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_UMProduto_Col) = objOperacoes.colOperacaoInsumos.Item(iIndice).sUMProduto
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_VersaoKitComp_Col) = objOperacoes.colOperacaoInsumos.Item(iIndice).sVersaoKitComp
        If objOperacoes.colOperacaoInsumos.Item(iIndice).dPercentualPerda > 0 Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_PerdaInsumos_Col) = Format(objOperacoes.colOperacaoInsumos.Item(iIndice).dPercentualPerda, "Percent")
        End If
        Call Combo_Seleciona_ItemData(ComposicaoInsumos, objOperacoes.colOperacaoInsumos.Item(iIndice).iComposicao)
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_ComposicaoInsumos_Col) = ComposicaoInsumos.Text
        If objOperacoes.colOperacaoInsumos.Item(iIndice).dCustoStandard > 0 Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_CustoStandardInsumos_Col) = Format(objOperacoes.colOperacaoInsumos.Item(iIndice).dCustoStandard, "Standard")
        End If
        
    Next

    objGridOperacaoInsumos.iLinhasExistentes = objOperacoes.colOperacaoInsumos.Count
    
    Preenche_GridOperacaoInsumos = SUCESSO
    
    Exit Function

Erro_Preenche_GridOperacaoInsumos:

    Preenche_GridOperacaoInsumos = gErr

    Select Case gErr

        Case 134787, 134788

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174283)

    End Select

    Exit Function

End Function

Private Function Move_OperacaoInsumos_Memoria(objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim objOperacaoInsumos As ClassOperacaoInsumos
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objKit As ClassKit

On Error GoTo Erro_Move_OperacaoInsumos_Memoria

    'Ir preenchendo a colecao no objOperacoes com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridOperacaoInsumos.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then Exit For
        
        'Verifica se Quantidade está preenchida
        If Len(Trim(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_QuantidadeProduto_Col))) = 0 Then gError 134897

        'Verifica se UMProduto está preenchida
        If Len(Trim(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_UMProduto_Col))) = 0 Then gError 134898

        'Verifica se Composicao está preenchida
        If Len(Trim(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_ComposicaoInsumos_Col))) = 0 Then gError 134899
        
        Set objProdutos = New ClassProduto
        
        lErro = CF("Produto_Formata", GridOperacaoInsumos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 134789
        
        objProdutos.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134790
        
        Set objOperacaoInsumos = New ClassOperacaoInsumos
        
        objOperacaoInsumos.lNumIntDocOper = objOperacoes.lNumIntDoc
        objOperacaoInsumos.sProduto = objProdutos.sCodigo
        objOperacaoInsumos.dQuantidade = StrParaDbl(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_QuantidadeProduto_Col))
        objOperacaoInsumos.sUMProduto = GridOperacaoInsumos.TextMatrix(iIndice, iGrid_UMProduto_Col)
        objOperacaoInsumos.iComposicao = Composicao_Extrai(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_ComposicaoInsumos_Col))
        objOperacaoInsumos.dPercentualPerda = StrParaDbl(Val(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_PerdaInsumos_Col)) / 100)
        objOperacaoInsumos.dCustoStandard = StrParaDbl(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_CustoStandardInsumos_Col))
        objOperacaoInsumos.sVersaoKitComp = GridOperacaoInsumos.TextMatrix(iIndice, iGrid_VersaoKitComp_Col)
        
        If objProdutos.iCompras = PRODUTO_PRODUZIVEL Then
        
            If Len(objOperacaoInsumos.sVersaoKitComp) = 0 Then
            
                Set objKit = New ClassKit
        
                objKit.sProdutoRaiz = objProdutos.sCodigo
            
                'Le as Versoes Ativas e a Padrao
                lErro = CF("Kit_Le_Padrao", objKit)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 139046
                
                If lErro <> SUCESSO Then gError 139047
                
                objOperacaoInsumos.sVersaoKitComp = objKit.sVersao
            
            End If
        
        End If
        
        objOperacoes.colOperacaoInsumos.Add objOperacaoInsumos
    
    Next

    Move_OperacaoInsumos_Memoria = SUCESSO

    Exit Function

Erro_Move_OperacaoInsumos_Memoria:

    Move_OperacaoInsumos_Memoria = gErr

    Select Case gErr
    
        Case 134897
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_QUANTIDADE_NAO_PREENCHIDA", gErr)
        
        Case 134898
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_UMPRODUTO_NAO_PREENCHIDA", gErr)
        
        Case 134899
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_COMPOSICAO_NAO_PREENCHIDA", gErr)
        
        Case 134789, 134790, 139046
            'erros tratados nas rotinas chamadas
            
        Case 139047
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_PADRAO_NAO_LOCALIZADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174284)

    End Select

    Exit Function

End Function

Private Sub TabStrip2_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index <> iFrameAtualOper Then

        If TabStrip_PodeTrocarTab(iFrameAtualOper, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame2(TabStrip2.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame2(iFrameAtualOper).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtualOper = TabStrip2.SelectedItem.Index
        
    End If

End Sub

Private Sub Composicao_Seleciona()
Dim iIndice As Integer

    Composicao.ListIndex = -1
    For iIndice = 0 To Composicao.ListCount - 1
        If Composicao.List(iIndice) = COMPOSICAO_VARIAVEL Then
            Composicao.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Private Sub ComposicaoInsumos_Seleciona()
Dim iIndice As Integer

    ComposicaoInsumos.ListIndex = -1
    For iIndice = 0 To ComposicaoInsumos.ListCount - 1
        If ComposicaoInsumos.List(iIndice) = COMPOSICAO_VARIAVEL Then
            ComposicaoInsumos.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Private Function Composicao_Extrai(sComposicaoGrid As String) As Integer

Dim iIndice As Integer

    For iIndice = 0 To ComposicaoInsumos.ListCount - 1
        If ComposicaoInsumos.List(iIndice) = sComposicaoGrid Then
            Composicao_Extrai = iIndice
            Exit For
        End If
    Next

End Function

Private Function Habilita_TaxaDeProducao() As Long
        
Dim lErro As Long

On Error GoTo Erro_Habilita_TaxaDeProducao

    If IgnoraTaxa.Value = vbChecked Then
        
        'Frame Taxa
        LabelTipo.Enabled = True
        Tipo.Enabled = True
        
        LabelTempoOperacao.Enabled = True
        TempoOperacao.Enabled = True
        
        LabelUMTempo.Enabled = True
        UMTempo.Enabled = True
        
        LabelTaxa.Enabled = True
                
        'Frame Tempos
        LabelTempoPreparacao.Enabled = True
        TempoPreparacao.Enabled = True
        
        LabelTempoMovimentacao.Enabled = True
        TempoMovimentacao.Enabled = True
        
        LabelTempoDescarga.Enabled = True
        TempoDescarga.Enabled = True
        
        'Frame Lotes
        LabelLoteMin.Enabled = True
        LoteMinimo.Enabled = True
        
        LabelLoteMax.Enabled = True
        LoteMaximo.Enabled = True
        
        LabelPadrao.Enabled = True
        LotePadrao.Enabled = True
        
        LabelMaquina.Enabled = True
        Maquina.Enabled = True
        
        lErro = Inicializa_Padrao()
        If lErro <> SUCESSO Then gError 137546
        
        Call Preenche_LabelTaxaDeProducao
    
    Else
    
        'Frame Taxa
        LabelTipo.Enabled = False
        Tipo.Enabled = False
        Tipo.ListIndex = -1
        
        LabelTempoOperacao.Enabled = False
        TempoOperacao.Enabled = False
        TempoOperacao.Text = ""
        
        LabelUMTempo.Enabled = False
        UMTempo.Enabled = False
        UMTempo.ListIndex = -1
        
        LabelTaxa.Enabled = False
        LabelTaxaDeProducao.Caption = ""
        
        'Frame Tempos
        LabelTempoPreparacao.Enabled = False
        TempoPreparacao.Enabled = False
        TempoPreparacao.Text = ""
        
        LabelTempoMovimentacao.Enabled = False
        TempoMovimentacao.Enabled = False
        TempoMovimentacao.Text = ""
        
        LabelTempoDescarga.Enabled = False
        TempoDescarga.Enabled = False
        TempoDescarga.Text = ""
        
        'Frame Lotes
        LabelLoteMin.Enabled = False
        LoteMinimo.Enabled = False
        LoteMinimo.Text = ""
        
        LabelLoteMax.Enabled = False
        LoteMaximo.Enabled = False
        LoteMaximo.Text = ""
        
        LabelPadrao.Enabled = False
        LotePadrao.Enabled = False
        LotePadrao.Text = ""
        
        LabelMaquina.Enabled = False
        Maquina.Enabled = False
        Maquina.Text = ""
        DescMaquina.Caption = ""
    
        FrameLotes.Caption = "Lotes"
        
    End If
    
    Habilita_TaxaDeProducao = SUCESSO
    
    Exit Function
    
Erro_Habilita_TaxaDeProducao:

    Habilita_TaxaDeProducao = gErr

    Select Case gErr
    
        Case 137546
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174285)

    End Select

    Exit Function

End Function

Private Sub VersaoKitComp_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VersaoKitComp_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub VersaoKitComp_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub VersaoKitComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = VersaoKitComp
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PerdaInsumos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PerdaInsumos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub PerdaInsumos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub PerdaInsumos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = PerdaInsumos
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComposicaoInsumos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComposicaoInsumos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub ComposicaoInsumos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub ComposicaoInsumos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = ComposicaoInsumos
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoStandardInsumos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoStandardInsumos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub CustoStandardInsumos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub CustoStandardInsumos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = CustoStandardInsumos
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Sub Calcula_Proxima_Chave(sChave As String)

Dim iNumero As Integer

    iNumero = iProxChave
    iProxChave = iProxChave + 1
    sChave = "X" & Right$(CStr(100000 + iNumero), 5)

End Sub

Sub Recalcula_Nivel_Sequencial()
'(re)calcula niveis e sequencias de toda a estrutura
'deve ser chamada apos a remocao de algum node

Dim iIndice As Integer

    If Roteiro.Nodes.Count = 0 Then Exit Sub

    For iIndice = LBound(aNivelSequencial) To UBound(aNivelSequencial)
        aNivelSequencial(iIndice) = 0
    Next

    iUltimoNivel = 0

    'chamar rotina que recalcula recursivamente os campos nivel e sequencial (Nivel e SeqArvore)
    Call Calcula_Nivel_Sequencial(Roteiro.Nodes.Item(1), 0, 0)

End Sub

Sub Calcula_Nivel_Sequencial(objNode As Node, iNivel As Integer, iPosicaoAtual As Integer)
'parte recursiva do recalculo de nivel e sequencial, atuando a partir do node passado
'iNivel informa o nivel deste node

Dim objOperacoes As New ClassOperacoes
Dim sChave1 As String

    sChave1 = objNode.Tag

    Set objOperacoes = colComponentes.Item(sChave1)

    aNivelSequencial(iNivel) = aNivelSequencial(iNivel) + 1

    iPosicaoAtual = iPosicaoAtual + 1
    aSeqPai(iNivel) = iPosicaoAtual

    objOperacoes.iSeqArvore = aNivelSequencial(iNivel)

    If iNivel > 0 Then
        objOperacoes.iSeqPai = aSeqPai(iNivel - 1)
    Else
        objOperacoes.iSeqPai = 0
    End If
    
    objOperacoes.iSeq = iPosicaoAtual

    objOperacoes.iNivel = iNivel
    
    colComponentes.Remove sChave1
    colComponentes.Add objOperacoes, sChave1

    If objNode.Children > 0 Then
        Call Calcula_Nivel_Sequencial(objNode.Child, iNivel + 1, iPosicaoAtual)
    End If

    If objNode.Index <> objNode.LastSibling.Index Then Call Calcula_Nivel_Sequencial(objNode.Next, iNivel, iPosicaoAtual)

    If iNivel > iUltimoNivel Then iUltimoNivel = iNivel
   
End Sub

Function Limpa_Operacoes() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Operacoes
        
    Nivel.Caption = ""
    Sequencial.Caption = ""
    
    CodigoCompetencia.PromptInclude = False
    CodigoCompetencia.Text = ""
    CodigoCompetencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = ""
    
    CodigoCTPadrao.PromptInclude = False
    CodigoCTPadrao.Text = ""
    CodigoCTPadrao.PromptInclude = True
    
    DescricaoCTPadrao.Caption = ""
    
    NumMaxMaqPorOperItem.PromptInclude = False
    NumMaxMaqPorOperItem.Text = ""
    NumMaxMaqPorOperItem.PromptInclude = True
    
    Repeticao.PromptInclude = False
    Repeticao.Text = ""
    Repeticao.PromptInclude = True
    
    bOperacaoNova = True
    
    Observacao.Text = ""
    
    Call Grid_Limpa(objGridOperacaoInsumos)
    
    IgnoraTaxa.Value = vbUnchecked
    
    lErro = Habilita_TaxaDeProducao()
    If lErro <> SUCESSO Then gError 137547
    
    iAlterado = 0
    iCompetenciaAlterada = 0

    Limpa_Operacoes = SUCESSO

    Exit Function

Erro_Limpa_Operacoes:

    Limpa_Operacoes = gErr

    Select Case gErr
    
        Case 137547
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174286)

    End Select

    Exit Function

End Function

Private Function Move_Operacoes_Memoria(ByVal objOperacoes As ClassOperacoes, ByVal objCompetencias As ClassCompetencias, ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Operacoes_Memoria
        
    objCompetencias.sNomeReduzido = CodigoCompetencia.Text
    
    'Verifica a Competencia no BD a partir do Código
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 134792

    objOperacoes.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 134793
        
        objOperacoes.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Trim(Observacao.Text)) <> 0 Then objOperacoes.sObservacao = Observacao.Text
    
    objOperacoes.iIgnoraTaxaProducao = StrParaInt(IgnoraTaxa.Value)
    
    objOperacoes.iNumMaxMaqPorOper = StrParaInt(NumMaxMaqPorOperItem.Text)
    objOperacoes.iNumRepeticoes = StrParaInt(Repeticao.Text)
    
    If IgnoraTaxa.Value = vbChecked Then
    
        lErro = Move_OperacoesTempo_Memoria(objOperacoes)
        If lErro <> SUCESSO Then gError 134794
    
    End If
    
    lErro = Move_OperacaoInsumos_Memoria(objOperacoes)
    If lErro <> SUCESSO Then gError 134795

    Move_Operacoes_Memoria = SUCESSO

    Exit Function

Erro_Move_Operacoes_Memoria:

    Move_Operacoes_Memoria = gErr

    Select Case gErr

        Case 134792 To 134795
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174287)

    End Select

    Exit Function

End Function

Private Function Move_OperacoesTempo_Memoria(objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim objOperacoesTempo As ClassOperacoesTempo
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Move_OperacoesTempo_Memoria

    Set objOperacoesTempo = New ClassOperacoesTempo

    objOperacoesTempo.lNumIntDocOperacao = objOperacoes.lNumIntDoc
    objOperacoesTempo.iTipo = Codigo_Extrai(Tipo.Text)
    objOperacoesTempo.dTempoOperacao = StrParaDbl(TempoOperacao.Text)
    objOperacoesTempo.sUMTempo = UMTempo.Text
    objOperacoesTempo.dTempoPreparacao = StrParaDbl(TempoPreparacao.Text)
    objOperacoesTempo.dTempoMovimentacao = StrParaDbl(TempoMovimentacao.Text)
    objOperacoesTempo.dTempoDescarga = StrParaDbl(TempoDescarga.Text)
    objOperacoesTempo.dLoteMax = StrParaDbl(LoteMaximo.Text)
    objOperacoesTempo.dLoteMin = StrParaDbl(LoteMinimo.Text)
    objOperacoesTempo.dLotePadrao = StrParaDbl(LotePadrao.Text)

    If Len(Maquina.Text) > 0 Then
        
        objMaquinas.sNomeReduzido = Maquina.Text
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 180342
        
        objOperacoesTempo.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If
    
    Set objOperacoes.objOperacoesTempo = objOperacoesTempo

    Move_OperacoesTempo_Memoria = SUCESSO

    Exit Function

Erro_Move_OperacoesTempo_Memoria:

    Move_OperacoesTempo_Memoria = gErr

    Select Case gErr
    
        Case 180342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174288)

    End Select

    Exit Function

End Function

Sub Remove_Filhos(objNode As Node)
'remove objNode, seus irmaos e filhos de colComponentes

        colComponentes.Remove (objNode.Tag)

        If objNode.Children > 0 Then

            Call Remove_Filhos(objNode.Child)

        End If

        If objNode <> objNode.LastSibling Then Call Remove_Filhos(objNode.Next)

    Exit Sub

End Sub

Function Preenche_Operacoes(objOperacoes As ClassOperacoes) As Long
'preenche as tabs de Detalhes, Insumos e Produção à partir dos dados de objOperacoes

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Preenche_Operacoes

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 134796

    Nivel.Caption = objOperacoes.iNivel
    Sequencial.Caption = objOperacoes.iSeqArvore
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.lNumIntDoc = objOperacoes.lNumIntDocCompet
    
    lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134336 Then gError 134797
    
    CodigoCompetencia.PromptInclude = False
    CodigoCompetencia.Text = objCompetencias.sNomeReduzido
    CodigoCompetencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = objCompetencias.sDescricao
    
    If objOperacoes.iNumMaxMaqPorOper <> 0 Then
        NumMaxMaqPorOperItem.PromptInclude = False
        NumMaxMaqPorOperItem.Text = objOperacoes.iNumMaxMaqPorOper
        NumMaxMaqPorOperItem.PromptInclude = True
    End If
    
    If objOperacoes.iNumRepeticoes <> 0 Then
        Repeticao.PromptInclude = False
        Repeticao.Text = objOperacoes.iNumRepeticoes
        Repeticao.PromptInclude = True
    End If
    
    If objOperacoes.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = objOperacoes.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 134798
        
        CodigoCTPadrao.PromptInclude = False
        CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        CodigoCTPadrao.PromptInclude = True
        
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If
    
    Observacao.Text = objOperacoes.sObservacao
    
    If objOperacoes.iIgnoraTaxaProducao = MARCADO Then
    
        IgnoraTaxa.Value = vbChecked
        
        lErro = Habilita_TaxaDeProducao()
        If lErro <> SUCESSO Then gError 137548
        
        lErro = Preenche_OperacoesTempo(objOperacoes)
        If lErro <> SUCESSO Then gError 134799
        
    End If

    lErro = Preenche_GridOperacaoInsumos(objOperacoes)
    If lErro <> SUCESSO Then gError 134800

    iAlterado = 0
    iCompetenciaAlterada = 0

    Preenche_Operacoes = SUCESSO

    Exit Function

Erro_Preenche_Operacoes:

    Preenche_Operacoes = gErr

    Select Case gErr
    
        Case 134796 To 134800, 137548

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174289)

    End Select

    Exit Function

End Function

Private Sub Roteiro_Collapse(ByVal Node As MSComctlLib.Node)
    Roteiro_NodeClick Node
End Sub

Private Function Preenche_OperacoesTempo(objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim objOperacoesTempo As ClassOperacoesTempo
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Preenche_OperacoesTempo

    Set objOperacoesTempo = objOperacoes.objOperacoesTempo
        
    'Frame Taxa
    Call Combo_Seleciona_ItemData(Tipo, objOperacoesTempo.iTipo)
    
    TempoOperacao.Text = Formata_Estoque(objOperacoesTempo.dTempoOperacao)
    
    UMTempo.Text = objOperacoesTempo.sUMTempo
    
    LabelTaxaDeProducao.Caption = ""
    
    'Frame Tempos
    TempoPreparacao.Text = Formata_Estoque(objOperacoesTempo.dTempoPreparacao)
    
    TempoMovimentacao.Text = Formata_Estoque(objOperacoesTempo.dTempoMovimentacao)
    
    TempoDescarga.Text = Formata_Estoque(objOperacoesTempo.dTempoDescarga)
    
    'Frame Lotes
    LoteMinimo.Text = Formata_Estoque(objOperacoesTempo.dLoteMin)
    
    LoteMaximo.Text = Formata_Estoque(objOperacoesTempo.dLoteMax)
    
    LotePadrao.Text = Formata_Estoque(objOperacoesTempo.dLotePadrao)

    FrameLotes.Caption = "Lotes"
    
    If objOperacoesTempo.lNumIntDocMaq <> 0 Then
        
        objMaquinas.lNumIntDoc = objOperacoesTempo.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 180341
        
        Maquina.Text = objMaquinas.sNomeReduzido
        DescMaquina.Caption = objMaquinas.sDescricao
        
    Else
    
        Maquina.Text = ""
        DescMaquina.Caption = ""
        
    End If
    
    Call Preenche_LabelTaxaDeProducao
    
    Preenche_OperacoesTempo = SUCESSO
    
    Exit Function

Erro_Preenche_OperacoesTempo:

    Preenche_OperacoesTempo = gErr

    Select Case gErr

        Case 135254, 180341

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174290)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VersaoKitComp(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_VersaoKitComp

    Set objGridInt.objControle = VersaoKitComp
    
    GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col) = VersaoKitComp.Text
    
    'Se o campo foi preenchido
    If Len(Trim(VersaoKitComp.Text)) > 0 Then

        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134801

    Saida_Celula_VersaoKitComp = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoKitComp:

    Saida_Celula_VersaoKitComp = gErr

    Select Case gErr

        Case 134801
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174291)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PerdaInsumos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PerdaInsumos

    Set objGridInt.objControle = PerdaInsumos
    
    'Se o campo foi preenchido
    If Len(Trim(PerdaInsumos.Text)) > 0 Then
    
        lErro = Porcentagem_Critica(PerdaInsumos.Text)
        If lErro <> SUCESSO Then gError 137065
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134802

    Saida_Celula_PerdaInsumos = SUCESSO

    Exit Function

Erro_Saida_Celula_PerdaInsumos:

    Saida_Celula_PerdaInsumos = gErr

    Select Case gErr

        Case 134802, 137065
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174292)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ComposicaoInsumos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ComposicaoInsumos

    Set objGridInt.objControle = ComposicaoInsumos
    
    'Se o campo foi preenchido
    If Len(Trim(ComposicaoInsumos.Text)) > 0 Then
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134803

    Saida_Celula_ComposicaoInsumos = SUCESSO

    Exit Function

Erro_Saida_Celula_ComposicaoInsumos:

    Saida_Celula_ComposicaoInsumos = gErr

    Select Case gErr

        Case 134803
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174293)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CustoStandardInsumos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoStandardInsumos

    Set objGridInt.objControle = CustoStandardInsumos
    
    'Se o campo foi preenchido
    If Len(Trim(CustoStandardInsumos.Text)) > 0 Then
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperacaoInsumos.Row - GridOperacaoInsumos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134804

    Saida_Celula_CustoStandardInsumos = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoStandardInsumos:

    Saida_Celula_CustoStandardInsumos = gErr

    Select Case gErr

        Case 134804
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174294)

    End Select

    Exit Function

End Function

Private Sub Carrega_ComboVersoes(ByVal sProdutoRaiz As String)
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    VersaoKitComp.Enabled = True
    
    'Limpa a Combo
    VersaoKitComp.Clear
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProdutoRaiz
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 134805
    
    VersaoKitComp.AddItem ""
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        VersaoKitComp.AddItem (objKit.sVersao)
        
    Next
    
    Exit Sub
    
Erro_Carrega_ComboVersoes:

    Select Case gErr
    
        Case 134805
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174295)
    
    End Select
    
End Sub

Private Sub VersaoKitComp_Seleciona(sVersao As String)
Dim iIndice As Integer

    VersaoKitComp.ListIndex = -1
    For iIndice = 0 To VersaoKitComp.ListCount - 1
        If VersaoKitComp.List(iIndice) = sVersao Then
            VersaoKitComp.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Private Function VersaoKitComp_SelecionaPadrao(sProduto As String)

Dim lErro As Long
Dim objKit As New ClassKit
    
On Error GoTo Erro_VersaoKitComp_SelecionaPadrao
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProduto
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 134806
        
    Call VersaoKitComp_Seleciona(objKit.sVersao)
    
    VersaoKitComp_SelecionaPadrao = SUCESSO
    
    Exit Function

Erro_VersaoKitComp_SelecionaPadrao:

    VersaoKitComp_SelecionaPadrao = gErr
    
    Select Case gErr

        Case 134806
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174296)

    End Select

    Exit Function
    
End Function

Function Limpa_Arvore_Roteiro() As Long
'Limpa a Arvore do Roteiro

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Arvore_Roteiro

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 134808

    Roteiro.Nodes.Clear
    Set colComponentes = New Collection
    
    iProxChave = 1

    Limpa_Arvore_Roteiro = SUCESSO

    Exit Function

Erro_Limpa_Arvore_Roteiro:

    Limpa_Arvore_Roteiro = gErr
    
    Select Case gErr

        Case 134808

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174297)

    End Select

    Exit Function

End Function

Private Function Preenche_LabelTaxaDeProducao() As Long

Dim lErro As Long
Dim dQtdeUnitaria As Double
Dim sTaxaDeProducao As String

On Error GoTo Erro_Preenche_LabelTaxaDeProducao

    sTaxaDeProducao = ""
    
    If Codigo_Extrai(Tipo.Text) = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
    
        If Len(TempoOperacao.Text) <> 0 And Len(UMTempo.Text) <> 0 Then
            sTaxaDeProducao = Formata_Estoque(StrParaDbl(TempoOperacao.Text)) & " " & UMTempo.Text
        End If
    
    Else
    
        If Len(Quantidade.Text) <> 0 And Len(UM.Text) <> 0 Then
            
            If Len(TempoOperacao.Text) <> 0 And Len(UMTempo.Text) <> 0 Then
                
                If StrParaDbl(TempoOperacao.Text) <> 0 Then
                    dQtdeUnitaria = CDbl(Quantidade.Text) / StrParaDbl(TempoOperacao.Text)
                    sTaxaDeProducao = Formata_Estoque(dQtdeUnitaria) & " " & UM.Text
                    sTaxaDeProducao = sTaxaDeProducao & "/" & UMTempo.Text
                Else
                    sTaxaDeProducao = ""
                End If
            
            End If
    
        End If
    
    End If
            
    LabelTaxaDeProducao.Caption = sTaxaDeProducao
    
    Exit Function

Erro_Preenche_LabelTaxaDeProducao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174298)

    End Select

    Exit Function

End Function

Private Sub UMTempo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UMTempo_Validate

    'Veifica se UMTempo está preenchida
    If Len(Trim(UMTempo.Text)) <> 0 Then

       If IgnoraTaxa.Value = vbChecked Then Call Preenche_LabelTaxaDeProducao

    End If

    Exit Sub

Erro_UMTempo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174299)

    End Select

    Exit Sub

End Sub

Private Sub UMTempo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
   
On Error GoTo Erro_Tipo_Validate
    
    If Len(Tipo.Text) <> 0 Then
    
        Call Preenche_LabelTaxaDeProducao
    
    End If

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174300)

    End Select

    Exit Sub

End Sub

Function Carrega_Arvore(objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim objNode As Node
Dim lErro As Long, sChave As String, sChaveTvw As String
Dim iIndice As Integer
Dim sTexto As String
Dim objOperacoes As ClassOperacoes
Dim objCompetencias As ClassCompetencias
Dim objProduto As ClassProduto
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim sNomeReduzidoProduto As String
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Carrega_Arvore

    Set objProduto = New ClassProduto
    
    sCodProduto = ProdutoRaiz.Text
    
    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sCodProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 137532
    
    sNomeReduzidoProduto = objProduto.sNomeReduzido
    
    For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes

        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objOperacoes.lNumIntDocCompet
        
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 134809

        'prepara texto que identificará a nova Operação que está sendo incluida
        sTexto = objCompetencias.sNomeReduzido
        
        sTexto = sTexto & " (" & sNomeReduzidoProduto
        
        If objOperacoes.lNumIntDocCT > 0 Then
        
            Set objCentrodeTrabalho = New ClassCentrodeTrabalho
            
            objCentrodeTrabalho.lNumIntDoc = objOperacoes.lNumIntDocCT
            
            lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 134810
        
            sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
           
        End If
        
        sTexto = sTexto & ")"
        
        'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
        Call Calcula_Proxima_Chave(sChaveTvw)
        
        sChave = sChaveTvw
        sChaveTvw = sChaveTvw & objCompetencias.lCodigo

        If objOperacoes.iNivel = 0 Then

            Set objNode = Roteiro.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)

        Else

            Set objNode = Roteiro.Nodes.Add(objOperacoes.iSeqPai, tvwChild, sChaveTvw, sTexto)

        End If
                
        Roteiro.Nodes.Item(objNode.Index).Expanded = True
        
        colComponentes.Add objOperacoes, sChave
        
        objNode.Tag = sChave
        
    Next

    'se houver árvore ...
    If Roteiro.Nodes.Count > 0 Then
        
        'selecionar a raiz
        Set Roteiro.SelectedItem = Roteiro.Nodes.Item(1)
        Roteiro.SelectedItem.Selected = True
        
        'e carregar as operações pertinentes
        Call Roteiro_NodeClick(Roteiro.Nodes.Item(1))
        
        bOperacaoNova = False
        
    End If

    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr

        Case 134809, 134810, 137531, 137532
            'erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174301)

    End Select

    Exit Function

End Function

Private Function Compara_TaxaDeProducao(ByVal objTaxaDeProducao As ClassTaxaDeProducao, objTaxaDeProducao_Melhor As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim dFatorTempo As Double
Dim dFatorProduto As Double
Dim dFatorTempo_Melhor As Double
Dim dFatorProduto_Melhor As Double
Dim dTaxaNova As Double
Dim dTaxaMelhor As Double

On Error GoTo Erro_Compara_TaxaDeProducao
    
    'se não existe Taxa Melhor esta é a Melhor
    If objTaxaDeProducao_Melhor.lNumIntDoc = 0 Then
    
        Set objTaxaDeProducao_Melhor = objTaxaDeProducao
    
    Else
    
        'Encontra Fator Taxa - Nova
        
        'Descobre o fator de conversao da UM recebida p/UM padrão de tempo
        lErro = CF("UM_Conversao_Trans", gobjEST.iClasseUMTempo, objTaxaDeProducao.sUMTempo, TAXA_CONSUMO_TEMPO_PADRAO, dFatorTempo)
        If lErro <> SUCESSO Then gError 134811
    
        Set objProduto = New ClassProduto
    
        sProduto = ProdutoRaiz.Text
    
        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 134812
    
        'Descobre o fator de conversao da UM recebida p/UM de Estoque do Produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTaxaDeProducao.sUMProduto, objProduto.sSiglaUMEstoque, dFatorProduto)
        If lErro <> SUCESSO Then gError 134813
    
        'Encontra Fator Taxa - Melhor
        
        'Descobre o fator de conversao da UM recebida p/UM padrão de tempo
        lErro = CF("UM_Conversao_Trans", gobjEST.iClasseUMTempo, objTaxaDeProducao_Melhor.sUMTempo, TAXA_CONSUMO_TEMPO_PADRAO, dFatorTempo_Melhor)
        If lErro <> SUCESSO Then gError 134814
    
        'Descobre o fator de conversao da UM recebida p/UM de Estoque do Produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTaxaDeProducao_Melhor.sUMProduto, objProduto.sSiglaUMEstoque, dFatorProduto_Melhor)
        If lErro <> SUCESSO Then gError 134815
        
        'Calcula as Taxas
        dTaxaNova = (objTaxaDeProducao.dQuantidade * dFatorProduto) / (objTaxaDeProducao.dTempoOperacao * dFatorTempo)
    
        dTaxaMelhor = (objTaxaDeProducao_Melhor.dQuantidade * dFatorProduto_Melhor) / (objTaxaDeProducao_Melhor.dTempoOperacao * dFatorTempo_Melhor)
        
        'se a TaxaNova menor que TaxaMelhor -> é a nova TaxaMelhor
        If dTaxaMelhor < dTaxaNova Then
                
            Set objTaxaDeProducao_Melhor = objTaxaDeProducao
            
        End If
        
    End If
        
    Compara_TaxaDeProducao = SUCESSO
    
    Exit Function

Erro_Compara_TaxaDeProducao:

    Compara_TaxaDeProducao = gErr
    
    Select Case gErr
    
        Case 134811 To 134815
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174302)
        
    End Select
    
    Exit Function

End Function

Private Sub BotaoRelRoteiro_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim lNumIntRel As Long
Dim objRoteiro As New ClassRoteirosDeFabricacao
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sTsk As String

On Error GoTo Erro_BotaoImprimirOper_Click

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 138289

    If iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(Versao.Text)) <> 0 Then

        objRoteiro.sProdutoRaiz = sProdutoFormatado
        objRoteiro.sVersao = Versao.Text

        lErro = CF("RelRoteiroFabricacao_Prepara", objRoteiro, lNumIntRel)
        If lErro <> SUCESSO Then gError 138290
        
        If DetalharInsumos.Value = vbChecked Then
            sTsk = "RotFabrD"
        Else
            sTsk = "RotFabr"
        End If
                
        lErro = objRelatorio.ExecutarDireto("Roteiro de Fabricação", "", 0, sTsk, "NNUMINTREL", CStr(lNumIntRel))
        If lErro <> SUCESSO Then gError 138291
    
    End If
    
    Exit Sub
    
Erro_BotaoImprimirOper_Click:
    
    Select Case gErr
    
        Case 138289 To 138291
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174303)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoAbrirRoteiro_Click()

Dim lErro As Long
Dim objRoteiro As New ClassRoteirosDeFabricacao
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoAbrirRoteiro_Click

    'Se não tiver linha selecionada => Erro
    If GridOperacaoInsumos.Row = 0 Then gError 138310
        
    sProduto = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_Produto_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 138311
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 138312
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lê o Produto Componente do Kit
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138313
    
    'se o produto não estiver cadastrado... erro
    If lErro <> SUCESSO Then gError 138314
    
    If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 138315

    If Len(Trim(GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col))) = 0 Then gError 138316

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 138317

    objRoteiro.sProdutoRaiz = objProduto.sCodigo
    objRoteiro.sVersao = GridOperacaoInsumos.TextMatrix(GridOperacaoInsumos.Row, iGrid_VersaoKitComp_Col)

    lErro = Traz_RoteirosDeFabricacao_Tela(objRoteiro)
    If lErro <> SUCESSO Then gError 138318

    Exit Sub
    
Erro_BotaoAbrirRoteiro_Click:
    
    Select Case gErr
    
        Case 138310
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 138311, 138313, 138317, 138318

        Case 138312
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTORAIZ_ROTFABR_NAO_PREENCHIDO", gErr)
        
        Case 138314
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case 138315
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, objProduto.sCodigo)
    
        Case 138316
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSDEFABRICACAO_NAO_PREENCHIDO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174304)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoOnde_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRoteiro As New ClassRoteirosDeFabricacao
Dim sFiltro As String

On Error GoTo Erro_BotaoOnde_Click
    
    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 138320
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 138321
        
    If Len(Trim(Versao.Text)) = 0 Then gError 138322
    
    colSelecao.Add sProdutoFormatado
    colSelecao.Add Versao.Text
    
    sFiltro = "EXISTS (SELECT R.ProdutoRaiz FROM RoteirosDeFabricacao AS R, Operacoes AS O, OperacaoInsumos AS I WHERE R.NumIntDoc = O.NumIntDocRotFabr AND O.NumIntDoc = I.NumIntDocOper AND R.ProdutoRaiz = RoteirosdeFabricacao.ProdutoRaiz AND R.Versao = RoteirosdeFabricacao.Versao AND I.Produto = ? AND I.VersaoKitComp = ?)"
    
    Call Chama_Tela("RoteirosDeFabricacaoLista", colSelecao, objRoteiro, objEventoRoteiroDeFabricacao, sFiltro)
    
    Exit Sub

Erro_BotaoOnde_Click:

    Select Case gErr
    
        Case 138320

        Case 138321
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTORAIZ_ROTFABR_NAO_PREENCHIDO", gErr)

        Case 138322
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSDEFABRICACAO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174305)

    End Select

    Exit Sub
    
End Sub

Private Sub NumMaxMaqPorOper_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumMaxMaqPorOper_Validate

    'Veifica se NumMaxMaqPorOper está preenchida
    If Len(Trim(NumMaxMaqPorOper.Text)) <> 0 Then

       'Critica a NumMaxMaqPorOper
       lErro = Valor_Positivo_Critica(NumMaxMaqPorOper.Text)
       If lErro <> SUCESSO Then gError 141610
       
       If bOperacaoNova And StrParaInt(NumMaxMaqPorOperItem.Text) = 0 Then
            NumMaxMaqPorOperItem.PromptInclude = False
            NumMaxMaqPorOperItem.Text = NumMaxMaqPorOper.Text
            NumMaxMaqPorOperItem.PromptInclude = True
        End If
       
    End If

    Exit Sub

Erro_NumMaxMaqPorOper_Validate:

    Cancel = True

    Select Case gErr

        Case 141610

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174306)

    End Select

    Exit Sub

End Sub

Private Sub NumMaxMaqPorOper_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumMaxMaqPorOper, iAlterado)
    
End Sub

Private Sub NumMaxMaqPorOper_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumMaxMaqPorOperItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumMaxMaqPorOperItem_Validate

    'Veifica se NumMaxMaqPorOperItem está preenchida
    If Len(Trim(NumMaxMaqPorOperItem.Text)) <> 0 Then

       'Critica a NumMaxMaqPorOperItem
       lErro = Valor_Positivo_Critica(NumMaxMaqPorOperItem.Text)
       If lErro <> SUCESSO Then gError 141611
            
    End If

    Exit Sub

Erro_NumMaxMaqPorOperItem_Validate:

    Cancel = True

    Select Case gErr

        Case 141611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174307)

    End Select

    Exit Sub

End Sub

Private Sub NumMaxMaqPorOperItem_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumMaxMaqPorOperItem, iAlterado)
    
End Sub

Private Sub NumMaxMaqPorOperItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodigoCompetencia Then Call CompetenciaLabel_Click
        If Me.ActiveControl Is CodigoCTPadrao Then Call CTLabel_Click
        If Me.ActiveControl Is Produto Then Call BotaoProdutos_Click
        If Me.ActiveControl Is ProdutoRaiz Then Call LabelProduto_Click
        If Me.ActiveControl Is Versao Then Call LabelVersao_Click
    
    End If
    
End Sub

Private Sub BotaoMaquinasAptas_Click()

Dim lErro As Long
Dim objTaxaDeProducao As New ClassTaxaDeProducao
Dim colSelecao As New Collection
Dim objCompetencias As ClassCompetencias
Dim sFiltro As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCT As New ClassCentrodeTrabalho

On Error GoTo Erro_BotaoMaquinasAptas_Click

    If Len(Trim(CodigoCompetencia.Text)) = 0 Then gError 141968
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.sNomeReduzido = CodigoCompetencia.Text
    
    'Verifica a Competencia no BD a partir do NomeReduzido
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 141969
    
    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141971
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 141972

    If Len(Trim(CodigoCTPadrao.Text)) = 0 Then gError 141974

    objCT.sNomeReduzido = CodigoCTPadrao.Text
    
    'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCT)
    If lErro <> SUCESSO And lErro <> 134941 Then gError 141973
       
    sFiltro = "Ativo = ? And NumIntDocCompet = ? AND (Produto=? OR Produto = '') AND ((CodigoMaquinas IS NULL) OR (CodigoMaquinas IN (SELECT M.Codigo from CTMaquinas AS CTM, Maquinas AS M WHERE CTM.NumIntDocMaq = M.NumIntDoc AND CTM.NumIntDocCT = ?))) "
    
    colSelecao.Add TAXA_ATIVA
    colSelecao.Add objCompetencias.lNumIntDoc
    colSelecao.Add sProdutoFormatado
    colSelecao.Add objCT.lNumIntDoc
    
    objTaxaDeProducao.sProduto = sProdutoFormatado

    Call Chama_Tela("TaxaDeProducaoLista", colSelecao, objTaxaDeProducao, Nothing, sFiltro)

    Exit Sub

Erro_BotaoMaquinasAptas_Click:

    Select Case gErr
    
        Case 141968
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            
        Case 141969, 141971, 141973
            
        Case 141972
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
    
        Case 141974
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141975)

    End Select

    Exit Sub

End Sub

Private Sub LabelMaquina_Click()

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim colSelecao As New Collection
Dim objCT As New ClassCentrodeTrabalho
Dim sFiltro As String

On Error GoTo Erro_LabelMaquina_Click

    'Verifica se o CodigoCTPadrao foi preenchido
    If Len(Trim(CodigoCTPadrao.Text)) = 0 Then gError 180345

    'Verifica se a Maquina foi preenchida
    If Len(Trim(Maquina.Text)) <> 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Text
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 180339
            
    End If
            
    objCT.iFilialEmpresa = giFilialEmpresa
    objCT.sNomeReduzido = CodigoCTPadrao.Text
    
    'Verifica o CodigoCTPadrao, lendo no BD a partir do NomeReduzido
    lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
    If lErro <> SUCESSO And lErro <> 137210 Then gError 180346

    'Filtra as máquinas do Centro de Trabalho escolhido
    sFiltro = "NumIntDoc IN (SELECT CTM.NumIntDocMaq from CTMaquinas AS CTM WHERE CTM.NumIntDocCT = ?)"
    
    colSelecao.Add objCT.lNumIntDoc
    
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaquina, sFiltro)

    Exit Sub

Erro_LabelMaquina_Click:

    Select Case gErr
    
        Case 180339, 180346
        
        Case 180345
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180340)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMaquina_evSelecao(obj1 As Object)

Dim objMaquinas As New ClassMaquinas
Dim lErro As Long

On Error GoTo Erro_objEventoMaquina_evSelecao

    Set objMaquinas = obj1

    Maquina.Text = objMaquinas.sNomeReduzido
    DescMaquina.Caption = objMaquinas.sDescricao
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMaquina_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180338)

    End Select

    Exit Sub

End Sub

Private Sub Maquina_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMaq As ClassCTMaquinas
Dim bAchou As Boolean

On Error GoTo Erro_Maquina_Validate

    'Verifica se o CodigoCTPadrao foi preenchido
    If Len(Trim(CodigoCTPadrao.Text)) = 0 Then gError 180347
    
    'Verifica sua existencia
    lErro = CF("TP_Maquina_Le", Maquina, objMaquinas)
    If lErro <> SUCESSO Then gError 180336
    
    objCT.iFilialEmpresa = giFilialEmpresa
    objCT.sNomeReduzido = CodigoCTPadrao.Text
    
    'Verifica o CodigoCTPadrao, lendo no BD a partir do NomeReduzido
    lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
    If lErro <> SUCESSO And lErro <> 137210 Then gError 180343
    
    bAchou = False
    For Each objCTMaq In objCT.colMaquinas
        If objCTMaq.lNumIntDocMaq = objMaquinas.lNumIntDoc Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then gError 180344
    
    DescMaquina.Caption = objMaquinas.sDescricao
    
    Exit Sub

Erro_Maquina_Validate:

    Cancel = True

    Select Case gErr

        Case 180336, 180343
        
        Case 180344
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINAS_NAO_CADASTRADA_CT", gErr, objCT.lCodigo, objCT.iFilialEmpresa)

        Case 180347
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180337)

    End Select

    Exit Sub

End Sub

Private Sub Maquina_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Maquina, iAlterado)
    
End Sub

Private Sub Maquina_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Trata_TaxaDeProducao() As Long

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMaq As ClassCTMaquinas
Dim objTaxa As ClassTaxaDeProducao
Dim objCompetencias As New ClassCompetencias
Dim bAchou As Boolean
Dim vbResult As VbMsgBoxResult
Dim bTemCT As Boolean

On Error GoTo Erro_Trata_TaxaDeProducao
    
    If Len(Trim(CodigoCompetencia.Text)) <> 0 Then
    
        objCompetencias.sNomeReduzido = CodigoCompetencia.Text
        
        'Verifica a Competencia no BD a partir do NomeReduzido
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 180351
        
    End If

    lErro = CF("Produto_Formata", ProdutoRaiz.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 180352

    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
    
        objCT.iFilialEmpresa = giFilialEmpresa
        objCT.sNomeReduzido = CodigoCTPadrao.Text
        
        'Verifica se o CodigoCTPadrao existe, lendo no BD a partir do NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137210 Then gError 180353
        
    End If
    
    'Se alterou um dos campos que defina a taxa de produção
    If glNumIntCompetAnt <> objCompetencias.lNumIntDoc Or glNumIntCTAnt <> objCT.lNumIntDoc Or gsProdutoAnt <> sProduto Then
        
        If lErro = SUCESSO Then
        
            bTemCT = True
        
            bAchou = False
            For Each objCTMaq In objCT.colMaquinas
            
                Set objTaxa = New ClassTaxaDeProducao
            
                objTaxa.lNumIntDocCompet = objCompetencias.lNumIntDoc
                objTaxa.sProduto = sProduto
                objTaxa.lNumIntDocMaq = objCTMaq.lNumIntDocMaq
        
                lErro = CF("TaxaDeProducao_Seleciona", objTaxa)
                If lErro <> SUCESSO And lErro <> 134600 Then gError 180354
                
                'encontrou taxa ...
                If lErro = SUCESSO Then
                    bAchou = True
                    Exit For
                End If
            
            Next
        
            Set objTaxa = New ClassTaxaDeProducao
        
            objTaxa.lNumIntDocCompet = objCompetencias.lNumIntDoc
            objTaxa.sProduto = sProduto
        
            lErro = CF("TaxaDeProducao_Seleciona", objTaxa)
            If lErro <> SUCESSO And lErro <> 134600 Then gError 180354
            
            'encontrou taxa ...
            If lErro = SUCESSO Then bAchou = True
        
        
        Else
            bTemCT = False
            
            Set objTaxa = New ClassTaxaDeProducao
        
            objTaxa.lNumIntDocCompet = objCompetencias.lNumIntDoc
            objTaxa.sProduto = sProduto
            
            bAchou = False
            lErro = CF("TaxaDeProducao_Seleciona", objTaxa)
            If lErro <> SUCESSO And lErro <> 134600 Then gError 180355
    
            'encontrou taxa ...
            If lErro = SUCESSO Then bAchou = True
        
        End If
        
        If Not bAchou Then
        
            If bTemCT Then
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_SEM_TAXA_DE_PRODUCAO", objCompetencias.sNomeReduzido, ProdutoRaiz.Text, objCT.sNomeReduzido, objCT.iFilialEmpresa)
            Else
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_SEM_TAXA_DE_PRODUCAO2", objCompetencias.sNomeReduzido, ProdutoRaiz.Text)
            End If
            
            If vbResult = vbYes Then
            
                If objTaxa Is Nothing Then
                
                    Set objTaxa = New ClassTaxaDeProducao
                    
                    objTaxa.lNumIntDocCompet = objCompetencias.lNumIntDoc
                
                End If
            
                objTaxa.sProduto = sProduto
                objTaxa.lNumIntDocMaq = 0
                
                Call ComandoSeta_Fechar(Me.Name)
                
                Call Chama_Tela("TaxaDeProducao", objTaxa)
            
            End If
        
        End If
        
        glNumIntCompetAnt = objCompetencias.lNumIntDoc
        glNumIntCTAnt = objCT.lNumIntDoc
        gsProdutoAnt = sProduto
        
    End If
        
    Trata_TaxaDeProducao = SUCESSO
    
    Exit Function

Erro_Trata_TaxaDeProducao:

    Trata_TaxaDeProducao = gErr
    
    Select Case gErr
    
        Case 180351 To 180355

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180350)
        
    End Select
    
    Exit Function

End Function

Private Sub Repeticao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Repeticao_Validate

    'Veifica se Repeticao está preenchida
'    If Len(Trim(Repeticao.Text)) <> 0 Then

       'Critica a Repeticao
       lErro = Valor_Positivo_Critica(Repeticao.Text)
       If lErro <> SUCESSO Then gError 141611
            
'    End If

    Exit Sub

Erro_Repeticao_Validate:

    Cancel = True

    Select Case gErr

        Case 141611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174307)

    End Select

    Exit Sub

End Sub

Private Sub Repeticao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Repeticao, iAlterado)
    
End Sub

Private Sub Repeticao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
