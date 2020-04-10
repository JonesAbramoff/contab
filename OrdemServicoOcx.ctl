VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl OrdemServicoOcx 
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "'"
      Height          =   4875
      Index           =   1
      Left            =   45
      TabIndex        =   60
      Top             =   930
      Width           =   9255
      Begin VB.Frame Frame7 
         Caption         =   "Solicitação do Serviço"
         Height          =   2100
         Left            =   195
         TabIndex        =   110
         Top             =   1695
         Width           =   8955
         Begin VB.ComboBox FilialOP 
            Height          =   315
            Left            =   4905
            TabIndex        =   8
            Top             =   1455
            Width           =   1905
         End
         Begin MSMask.MaskEdBox CodigoSolSrv 
            Height          =   300
            Left            =   2055
            TabIndex        =   5
            Top             =   360
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoSolServ 
            Height          =   315
            Left            =   2055
            TabIndex        =   6
            Top             =   915
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Lote 
            Height          =   315
            Left            =   2055
            TabIndex        =   7
            Top             =   1455
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   1305
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   115
            Top             =   390
            Width           =   645
         End
         Begin VB.Label ProdutoLabel1 
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
            Left            =   1260
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   114
            Top             =   960
            Width           =   735
         End
         Begin VB.Label DescProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3585
            TabIndex        =   113
            Top             =   915
            Width           =   3210
         End
         Begin VB.Label LoteLabel 
            AutoSize        =   -1  'True
            Caption         =   "Lote:"
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
            Left            =   1545
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   112
            Top             =   1485
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FilialOP:"
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
            Index           =   3
            Left            =   4110
            TabIndex        =   111
            Top             =   1530
            Width           =   720
         End
      End
      Begin VB.CheckBox ImprimeAoGravar 
         Caption         =   "Imprimir ao Gravar uma O.S."
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
         Left            =   375
         TabIndex        =   9
         Top             =   4245
         Width           =   2715
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relatórios"
         Height          =   930
         Left            =   150
         TabIndex        =   64
         Top             =   3855
         Width           =   8985
         Begin VB.Frame Frame4 
            Caption         =   "Ordem de Serviço"
            Height          =   660
            Left            =   4485
            TabIndex        =   65
            Top             =   180
            Width           =   2775
            Begin VB.OptionButton OpcaoDetalhado 
               Caption         =   "Detalhada"
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
               Left            =   1335
               TabIndex        =   11
               Top             =   315
               Width           =   1215
            End
            Begin VB.OptionButton OpcaoSimples 
               Caption         =   "Normal"
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
               TabIndex        =   10
               Top             =   300
               Value           =   -1  'True
               Width           =   1065
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificação"
         Height          =   1485
         Left            =   165
         TabIndex        =   61
         Top             =   135
         Width           =   8985
         Begin VB.ComboBox Tipo 
            Height          =   315
            Left            =   2115
            TabIndex        =   4
            Top             =   900
            Width           =   4215
         End
         Begin VB.TextBox CodigoOS 
            Height          =   285
            Left            =   2115
            MaxLength       =   6
            TabIndex        =   0
            Top             =   345
            Width           =   1350
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   3450
            Picture         =   "OrdemServicoOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Numeração Automática"
            Top             =   345
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   6045
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   345
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   4980
            TabIndex        =   2
            Top             =   345
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1605
            TabIndex        =   109
            Top             =   945
            Width           =   450
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   6720
            TabIndex        =   99
            Top             =   435
            Width           =   615
         End
         Begin VB.Label StatusOS 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7395
            TabIndex        =   98
            Top             =   375
            Width           =   1305
         End
         Begin VB.Label CodigoOSLabel 
            AutoSize        =   -1  'True
            Caption         =   "Código O.S.:"
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
            Left            =   960
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   63
            Top             =   390
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4425
            TabIndex        =   62
            Top             =   390
            Width           =   480
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   3
      Left            =   45
      TabIndex        =   21
      Top             =   930
      Visible         =   0   'False
      Width           =   9360
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Operação"
         Height          =   2775
         Index           =   1
         Left            =   75
         TabIndex        =   73
         Top             =   2190
         Width           =   9195
         Begin VB.TextBox Observacao 
            Height          =   630
            Left            =   1380
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   74
            Top             =   1680
            Width           =   7635
         End
         Begin MSMask.MaskEdBox CodigoCTPadrao 
            Height          =   315
            Left            =   1380
            TabIndex        =   75
            Top             =   1275
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoCompetencia 
            Height          =   315
            Left            =   1380
            TabIndex        =   76
            Top             =   885
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label DescricaoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   71
            Top             =   1275
            Width           =   5145
         End
         Begin VB.Label DescricaoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   72
            Top             =   885
            Width           =   5145
         End
         Begin VB.Label LabelNivel 
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
            Left            =   720
            TabIndex        =   93
            Top             =   2430
            Width           =   540
         End
         Begin VB.Label Sequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3420
            TabIndex        =   92
            Top             =   2400
            Width           =   420
         End
         Begin VB.Label Nivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   91
            Top             =   2400
            Width           =   420
         End
         Begin VB.Label LabelSeq 
            AutoSize        =   -1  'True
            Caption         =   "Seqüêncial:"
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
            Left            =   2325
            TabIndex        =   90
            Top             =   2430
            Width           =   1020
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
            Height          =   330
            Left            =   180
            TabIndex        =   89
            Top             =   1710
            Width           =   1155
         End
         Begin VB.Label CTLabel 
            Caption         =   "C. Trabalho:"
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
            Left            =   225
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   88
            Top             =   1305
            Width           =   1110
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   87
            Top             =   915
            Width           =   1155
         End
         Begin VB.Label LabelCodigoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   86
            Top             =   2400
            Width           =   420
         End
         Begin VB.Label LabelCodigoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4320
            TabIndex        =   85
            Top             =   2400
            Width           =   420
         End
         Begin VB.Label VersaoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   84
            Top             =   480
            Width           =   1665
         End
         Begin VB.Label LabelDetVersao 
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
            Left            =   615
            TabIndex        =   83
            Top             =   540
            Width           =   690
         End
         Begin VB.Label ProdutoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   82
            Top             =   75
            Width           =   7635
         End
         Begin VB.Label LabelDetProduto 
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
            Left            =   555
            TabIndex        =   81
            Top             =   105
            Width           =   810
         End
         Begin VB.Label QtdeLabel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6015
            TabIndex        =   80
            Top             =   480
            Width           =   1470
         End
         Begin VB.Label LabelDetQtde 
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
            Height          =   210
            Left            =   4890
            TabIndex        =   79
            Top             =   540
            Width           =   1050
         End
         Begin VB.Label UMLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8235
            TabIndex        =   78
            Top             =   480
            Width           =   780
         End
         Begin VB.Label LabelDetUM 
            Caption         =   "U.M.:"
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
            Left            =   7725
            TabIndex        =   77
            Top             =   540
            Width           =   480
         End
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
         Height          =   435
         Left            =   8040
         Picture         =   "OrdemServicoOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Exclui a Operação da Árvore do Roteiro"
         Top             =   990
         Width           =   1275
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
         Height          =   435
         Left            =   8040
         Picture         =   "OrdemServicoOcx.ctx":1A10
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Inclui a Operação na Árvore do Roteiro"
         Top             =   555
         Width           =   1275
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   2
         Left            =   75
         TabIndex        =   34
         Top             =   2190
         Visible         =   0   'False
         Width           =   9195
         Begin VB.CommandButton BotaoPecas 
            Caption         =   "Peças"
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
            Left            =   75
            TabIndex        =   106
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   2385
            Width           =   1305
         End
         Begin VB.TextBox UMPecaPrev 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   4710
            MaxLength       =   50
            TabIndex        =   105
            Top             =   1080
            Width           =   600
         End
         Begin VB.ComboBox VersaoPecaPrev 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OrdemServicoOcx.ctx":325E
            Left            =   7425
            List            =   "OrdemServicoOcx.ctx":3260
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1020
            Width           =   1245
         End
         Begin VB.TextBox OrigemPecaPrev 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   100
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox DescricaoPecaPrev 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1950
            TabIndex        =   35
            Top             =   1050
            Width           =   3060
         End
         Begin MSMask.MaskEdBox CodigoPecaPrev 
            Height          =   315
            Left            =   435
            TabIndex        =   36
            Top             =   1035
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadePecaPrev 
            Height          =   315
            Left            =   5565
            TabIndex        =   37
            Top             =   1065
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridPecasPrev 
            Height          =   2130
            Left            =   75
            TabIndex        =   38
            Top             =   225
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   3757
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   4
         Left            =   75
         TabIndex        =   22
         Top             =   2190
         Visible         =   0   'False
         Width           =   9195
         Begin MSMask.MaskEdBox HorasMOPrev 
            Height          =   315
            Left            =   6030
            TabIndex        =   26
            Top             =   1440
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoMO 
            Caption         =   "Tipo de Mão-de-Obra"
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
            Left            =   75
            TabIndex        =   25
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   2385
            Width           =   2505
         End
         Begin VB.TextBox CodigoMOPrev 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   195
            TabIndex        =   24
            Top             =   1410
            Width           =   1665
         End
         Begin VB.TextBox NomeRedMOPrev 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1425
            TabIndex        =   23
            Top             =   1425
            Width           =   4665
         End
         Begin MSFlexGridLib.MSFlexGrid GridMOPrev 
            Height          =   2085
            Left            =   75
            TabIndex        =   27
            Top             =   225
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   3678
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   3
         Left            =   75
         TabIndex        =   28
         Top             =   2190
         Visible         =   0   'False
         Width           =   9195
         Begin VB.CommandButton BotaoMaquinas 
            Caption         =   "Máquinas"
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
            Left            =   75
            TabIndex        =   30
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   2385
            Width           =   1305
         End
         Begin VB.TextBox NomeRedMaquinaPrev 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1665
            TabIndex        =   29
            Top             =   1065
            Width           =   4020
         End
         Begin MSMask.MaskEdBox HorasMaquinaPrev 
            Height          =   315
            Left            =   4905
            TabIndex        =   31
            Top             =   1470
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantMaquinaPrev 
            Height          =   315
            Left            =   3120
            TabIndex        =   32
            Top             =   1455
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaquinasPrev 
            Height          =   2055
            Left            =   75
            TabIndex        =   33
            Top             =   225
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   3625
            _Version        =   393216
         End
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
         Height          =   435
         Left            =   8040
         Picture         =   "OrdemServicoOcx.ctx":3262
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Altera a Operação da Árvore do Roteiro"
         Top             =   1410
         Width           =   1275
      End
      Begin VB.Frame FrameRoteiro 
         Caption         =   "Roteiro de Serviço:"
         Height          =   1875
         Left            =   30
         TabIndex        =   43
         Top             =   -15
         Width           =   7965
         Begin MSComctlLib.TreeView Roteiro 
            Height          =   1485
            Left            =   120
            TabIndex        =   44
            Top             =   270
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   2619
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
      Begin VB.Frame FrameItemOS 
         Caption         =   "Item da OS"
         Height          =   570
         Left            =   8040
         TabIndex        =   40
         Top             =   -15
         Width           =   1275
         Begin MSComCtl2.UpDown UpDownItemOS 
            Height          =   315
            Left            =   780
            TabIndex        =   41
            Top             =   210
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ItemOS 
            Height          =   315
            Left            =   270
            TabIndex        =   42
            Top             =   210
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
      End
      Begin VB.CommandButton BotaoImprimirOper 
         Caption         =   "Relatório de Operações"
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
         TabIndex        =   39
         ToolTipText     =   "Abre o Relatório de Operações"
         Top             =   1995
         Visible         =   0   'False
         Width           =   2235
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   3165
         Left            =   30
         TabIndex        =   46
         Top             =   1860
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   5583
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhe"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsão de Peças"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsão de Máquinas"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsão de Mão de Obra"
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
      Height          =   5160
      Index           =   2
      Left            =   45
      TabIndex        =   47
      Top             =   930
      Visible         =   0   'False
      Width           =   9330
      Begin VB.TextBox ObsSRV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4320
         TabIndex        =   104
         Top             =   2550
         Width           =   480
      End
      Begin MSMask.MaskEdBox NumIntDoc 
         Height          =   225
         Left            =   7365
         TabIndex        =   103
         Top             =   1935
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.ComboBox Versao 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "OrdemServicoOcx.ctx":4B88
         Left            =   2010
         List            =   "OrdemServicoOcx.ctx":4B8A
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   1545
         Width           =   1875
      End
      Begin VB.TextBox Obs 
         Height          =   825
         Left            =   90
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   96
         Top             =   3750
         Width           =   9150
      End
      Begin VB.CommandButton BotaoAcompanhamento 
         Caption         =   "Acompanhamento"
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
         Left            =   5805
         TabIndex        =   95
         Top             =   4680
         Width           =   1680
      End
      Begin VB.ComboBox StatusItem 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OrdemServicoOcx.ctx":4B8C
         Left            =   4035
         List            =   "OrdemServicoOcx.ctx":4B8E
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   1005
         Width           =   1920
      End
      Begin VB.TextBox UnidadeMed 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6195
         MaxLength       =   50
         TabIndex        =   51
         Top             =   1095
         Width           =   600
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   5910
         MaxLength       =   50
         TabIndex        =   50
         Top             =   630
         Width           =   2600
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
         Height          =   360
         Left            =   3630
         TabIndex        =   49
         Top             =   4680
         Width           =   1680
      End
      Begin VB.CommandButton BotaoServicos 
         Caption         =   "Serviços"
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
         Left            =   1425
         TabIndex        =   48
         Top             =   4680
         Width           =   1680
      End
      Begin MSMask.MaskEdBox Prioridade 
         Height          =   255
         Left            =   6600
         TabIndex        =   52
         Top             =   1605
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   285
         Left            =   6960
         TabIndex        =   53
         Top             =   1140
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Servico 
         Height          =   255
         Left            =   4140
         TabIndex        =   54
         Top             =   660
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   270
         Left            =   6105
         TabIndex        =   55
         Top             =   2010
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
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
      Begin MSMask.MaskEdBox DataPrevisaoFim 
         Height          =   255
         Left            =   4290
         TabIndex        =   56
         Top             =   1560
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataPrevisaoInicio 
         Height          =   255
         Left            =   4575
         TabIndex        =   57
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
         Height          =   3030
         Left            =   75
         TabIndex        =   58
         Top             =   375
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   5345
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label MensagemLabel 
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
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   97
         Top             =   3495
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serviços"
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
         TabIndex        =   59
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Pedido de Serviço"
      Height          =   615
      Index           =   0
      Left            =   45
      TabIndex        =   68
      Top             =   -15
      Width           =   6390
      Begin VB.CommandButton BotaoPedSRV 
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
         Left            =   3855
         TabIndex        =   19
         Top             =   225
         Width           =   1350
      End
      Begin VB.CommandButton BotaoVerPedSRV 
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
         Left            =   5280
         TabIndex        =   20
         Top             =   225
         Width           =   1005
      End
      Begin VB.ComboBox FilialPedSRV 
         Height          =   315
         Left            =   2310
         TabIndex        =   18
         Top             =   240
         Width           =   1500
      End
      Begin MSMask.MaskEdBox PedSRV 
         Height          =   300
         Left            =   855
         TabIndex        =   17
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label PedSRVLabel 
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
         Height          =   255
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   70
         Top             =   285
         Width           =   720
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
         Height          =   195
         Index           =   0
         Left            =   1845
         TabIndex        =   69
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6810
      ScaleHeight     =   495
      ScaleWidth      =   2610
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   30
      Width           =   2670
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   105
         Picture         =   "OrdemServicoOcx.ctx":4B90
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   615
         Picture         =   "OrdemServicoOcx.ctx":4C92
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1125
         Picture         =   "OrdemServicoOcx.ctx":4DEC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "OrdemServicoOcx.ctx":4F76
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2100
         Picture         =   "OrdemServicoOcx.ctx":54A8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5535
      Left            =   15
      TabIndex        =   67
      Top             =   600
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   9763
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "OrdemServicoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iLinhaAntiga As Integer
Dim iCodigoAlterado As Integer

Dim iPrestServAlterado As Integer

Dim gcolItemOS As Collection
'criado por causa da forma como foi construido a tela romaneiograde
Dim gobjOS As ClassOS
Dim gobjItemOS As ClassItemOS

Dim objGrid As AdmGrid
Dim iGrid_Servico_Col As Integer
Dim iGrid_Versao_Col  As Integer
Dim iGrid_DescricaoItem_Col  As Integer
Dim iGrid_UnidadeMed_Col  As Integer
Dim iGrid_Quantidade_Col  As Integer
Dim iGrid_Ccl_Col  As Integer
Dim iGrid_DataPrevInicio_Col  As Integer
Dim iGrid_DataPrevFim_Col  As Integer
Dim iGrid_Prioridade_Col  As Integer
Dim iGrid_StatusItem_Col  As Integer
Dim iGrid_NumIntDoc_Col As Integer
Dim iGrid_ObsSRV_Col As Integer

Dim iFrameAtual As Integer
Dim glTipoPadrao As Long

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoSolSRV As AdmEvento
Attribute objEventoSolSRV.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoProdutoSolicSRV As AdmEvento
Attribute objEventoProdutoSolicSRV.VB_VarHelpID = -1
Private WithEvents objEventoMaquinas As AdmEvento
Attribute objEventoMaquinas.VB_VarHelpID = -1
Private WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1
Private WithEvents objEventoPeca As AdmEvento
Attribute objEventoPeca.VB_VarHelpID = -1
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoPedido As AdmEvento
Attribute objEventoPedido.VB_VarHelpID = -1

'---------------------------------------
'Inserido por Jorge Specian - 29/04/2005
Dim gbOPValidaParaMRP As Boolean
Dim iFrameAtualOper As Integer

'Grid de OperacaoPecas
Dim objGridPecasPrev As AdmGrid
Dim iGrid_CodigoPecaPrev_Col As Integer
Dim iGrid_DescricaoPecaPrev_Col As Integer
Dim iGrid_OrigemPecaPrev_Col As Integer
Dim iGrid_QuantidadePecaPrev_Col As Integer
Dim iGrid_UMPecaPrev_Col As Integer
Dim iGrid_VersaoPecaPrev_Col As Integer

Dim colPMPItens As Collection
Dim gcolOSItens As Collection
Dim colComponentes As Collection
Dim iProxChave As Integer

'variaveis auxiliares para recalculo de nivel e sequencial
Dim aNivelSequencial(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o maior sequencial
Dim aSeqPai(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o SeqPai

Dim iUltimoNivel As Integer

Dim iCompetenciaAlterada As Integer
Dim iQtdeAlterada As Integer
Dim giItemOS As Integer

Private WithEvents objEventoCompetencias As AdmEvento
Attribute objEventoCompetencias.VB_VarHelpID = -1
Private WithEvents objEventoCentroDeTrabalho As AdmEvento
Attribute objEventoCentroDeTrabalho.VB_VarHelpID = -1

Private Const TAB_Operacoes = 3
'---------------------------------------

'Grid de Maquinas
Dim objGridMaquinasPrev As AdmGrid
Dim iGrid_NomeRedMaquinaPrev_Col As Integer
Dim iGrid_QuantMaquinaPrev_Col As Integer
Dim iGrid_HorasMaquinaPrev_Col As Integer

Dim objGridMOPrev As AdmGrid
Dim iGrid_NomeRedMOPrev_Col As Integer
Dim iGrid_CodigoMOPrev_Col As Integer
Dim iGrid_HorasMOPrev_Col As Integer


Private Function Inicializa_GridMovimentos() As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Serviço")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Ccl")
    objGrid.colColuna.Add ("Previsão Início")
    objGrid.colColuna.Add ("Previsão Fim")
    objGrid.colColuna.Add ("Prioridade")
    objGrid.colColuna.Add ("Status")
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Servico.Name)
    objGrid.colCampo.Add (Versao.Name)
    objGrid.colCampo.Add (DescricaoItem.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (Ccl.Name)
    objGrid.colCampo.Add (DataPrevisaoInicio.Name)
    objGrid.colCampo.Add (DataPrevisaoFim.Name)
    objGrid.colCampo.Add (Prioridade.Name)
    objGrid.colCampo.Add (StatusItem.Name)
    objGrid.colCampo.Add (NumIntDoc.Name)
    objGrid.colCampo.Add (ObsSRV.Name)

    'Colunas do Grid
    iGrid_Servico_Col = 1
    iGrid_Versao_Col = 2
    iGrid_DescricaoItem_Col = 3
    iGrid_UnidadeMed_Col = 4
    iGrid_Quantidade_Col = 5
    iGrid_Ccl_Col = 6
    iGrid_DataPrevInicio_Col = 7
    iGrid_DataPrevFim_Col = 8
    iGrid_Prioridade_Col = 9
    iGrid_StatusItem_Col = 10
    iGrid_NumIntDoc_Col = 11
    iGrid_ObsSRV_Col = 12

    objGrid.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridMovimentos.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    NumIntDoc.Width = 0
    ObsSRV.Width = 0
    
    Call Grid_Inicializa(objGrid)

    GridMovimentos.ColWidth(iGrid_NumIntDoc_Col) = 0
    GridMovimentos.ColWidth(iGrid_ObsSRV_Col) = 0

    Inicializa_GridMovimentos = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sVersaoAnt As String
Dim sNomeMaqAnt As String
Dim objRotFab As New ClassRoteirosDeFabricacao
Dim colRotFab As New Collection
Dim sPeca As String
Dim sPecaFormatado As String
Dim iPecaPreenchido As Integer
Dim objKit As New ClassKit
Dim colKits As New Collection


On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then

        'Verifica se produto está preenchido
        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 193804

        'Pesquisa o controle da coluna em questão
        Select Case objControl.Name
    
            'Servico
            Case Servico.Name
                
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    Servico.Enabled = False
                Else
                    Servico.Enabled = True
                End If

            Case Versao.Name
                
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_NumIntDoc_Col))) = 0 Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                
                Call Carrega_ComboVersoes(sProdutoFormatado)
                
'                sVersaoAnt = Versao.Text
'
'                Versao.Clear
'
'                If iProdutoPreenchido <> PRODUTO_VAZIO Then
'
'                    'Armazena o Produto Raiz do kit
'                    objKit.sProdutoRaiz = sProdutoFormatado
'
'                    'Le as Versoes Ativas e a Padrao
'                    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
'                    If lErro <> SUCESSO And lErro <> 106333 Then gError 131215
'
'                    'Carrega a Combo com os Dados da Colecao
'                    For Each objKit In colKits
'
'                        Versao.AddItem (objKit.sVersao)
'
'                    Next
'
'                    'Tento selecionar na Combo a Unidade anterior
'                    If Versao.ListCount <> 0 Then
'
'                        For iIndice = 0 To Versao.ListCount - 1
'
'                            If Versao.List(iIndice) = sVersaoAnt Then
'                                Versao.ListIndex = iIndice
'                                Exit For
'                            End If
'                        Next
'                    End If
'                End If

            Case UnidadeMed.Name, DescricaoItem.Name, ObsSRV.Name, NumIntDoc.Name
                'ficam sempre desabilitadas
    
            Case StatusItem.Name
    
                'habilita atualização da situacao da OP
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case Quantidade.Name
            
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                            
                                        
                            
            Case VersaoPecaPrev.Name
                
                'Verifica se produto está preenchido
                sPeca = GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_CodigoPecaPrev_Col)
        
                lErro = CF("Produto_Formata", sPeca, sPecaFormatado, iPecaPreenchido)
                If lErro <> SUCESSO Then gError 193804
                
                sVersaoAnt = VersaoPecaPrev.Text
                
                VersaoPecaPrev.Clear
    
                If iPecaPreenchido <> PRODUTO_VAZIO Then

                    'Armazena o Produto Raiz do kit
                    objKit.sProdutoRaiz = sPecaFormatado

                    'Le as Versoes Ativas e a Padrao
                    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
                    If lErro <> SUCESSO And lErro <> 106333 Then gError 195482

                    'Carrega a Combo com os Dados da Colecao
                    For Each objKit In colKits

                        VersaoPecaPrev.AddItem (objKit.sVersao)

                    Next

                    'Tento selecionar na Combo a Versao anterior
                    If VersaoPecaPrev.ListCount <> 0 Then

                        For iIndice = 0 To VersaoPecaPrev.ListCount - 1

                            If VersaoPecaPrev.List(iIndice) = sVersaoAnt Then
                                VersaoPecaPrev.ListIndex = iIndice
                                Exit For
                            End If
                        Next
                    End If
                End If
                            
            Case DescricaoPecaPrev.Name, UMPecaPrev.Name, OrigemPecaPrev.Name, NomeRedMOPrev.Name
            
                
'            'se são campos do Grid de uso de máquinas real
'            Case HorasReal.Name, QuantidadeMaquinaReal.Name
'
'                If StrParaDate(GridMaquinasPrev.TextMatrix(GridMaquinasPrev.Row, iGrid_DataReal_Col)) = DATA_NULA Then
'                    objControl.Enabled = False
'                Else
'                    objControl.Enabled = True
'                End If
'
'            Case NomeMaquinaReal.Name
'
'                'Se ainda não tiver data ou já tiver preechido a máquina então não deixa alterar
'                If StrParaDate(GridMaquinasPrev.TextMatrix(GridMaquinasPrev.Row, iGrid_DataReal_Col)) = DATA_NULA Or Len(Trim(GridMaquinasPrev.TextMatrix(GridMaquinasPrev.Row, iGrid_NomeMaquinaReal_Col))) > 0 Then
'                    objControl.Enabled = False
'                Else
'                    objControl.Enabled = True
'                End If
                    
'            'se são campos do Grid de uso de mão de obra real
'            Case HorasMOPrev.Name, QuantidadeMOPrev.Name, _
'                    TipoMaoDeObraReal.Name
'
'                If StrParaDate(GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_CodigoMOPrev_Col)) = DATA_NULA Or Len(Trim(GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_NomeRedMOPrev_Col))) = 0 Then
'                    objControl.Enabled = False
'                Else
'                    objControl.Enabled = True
'                End If
'
'            Case DataReal.Name
'
'                If MRP.Value = vbChecked Then
'                    If StrParaDate(GridMaquinasPrev.TextMatrix(GridMaquinasPrev.Row, iGrid_DataReal_Col)) = DATA_NULA Then
'                        objControl.Enabled = True
'                    Else
'                        objControl.Enabled = False
'                    End If
'                Else
'                    objControl.Enabled = False
'                End If
'
'            Case CodigoMOPrev.Name
'
'                If MRP.Value = vbChecked Then
'                    If StrParaDate(GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_CodigoMOPrev_Col)) = DATA_NULA Then
'                        objControl.Enabled = True
'                    Else
'                        objControl.Enabled = False
'                    End If
'                Else
'                    objControl.Enabled = False
'                End If
'
'            Case DescricaoTipoMO.Name
'
'                objControl.Enabled = False
'
'            Case NomeRedMOPrev.Name
'
'                If MRP.Value = vbChecked Then
'
'                    sNomeMaqAnt = NomeRedMOPrev.Text
'
'                    NomeRedMOPrev.Clear
'                    For iLinha = 1 To objGridMaquinasPrev.iLinhasExistentes
'                        If StrParaDate(GridMaquinasPrev.TextMatrix(iLinha, iGrid_DataReal_Col)) = StrParaDate(GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_CodigoMOPrev_Col)) Then
'                            NomeRedMOPrev.AddItem GridMaquinasPrev.TextMatrix(iLinha, iGrid_NomeMaquinaReal_Col)
'                        End If
'                    Next
'
'                    'Tento selecionar na Combo a Máquina anterior
'                    If NomeRedMOPrev.ListCount <> 0 Then
'
'                        For iIndice = 0 To NomeRedMOPrev.ListCount - 1
'
'                            If NomeRedMOPrev.List(iIndice) = sNomeMaqAnt Then
'                                NomeRedMOPrev.ListIndex = iIndice
'                                Exit For
'                            End If
'                        Next
'                    End If
'
'                    If StrParaDate(GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_CodigoMOPrev_Col)) = DATA_NULA Then
'                        objControl.Enabled = False
'                    Else
'                        objControl.Enabled = True
'                    End If
'                Else
'                    objControl.Enabled = False
'                End If
'
    
            Case Else
                
                'Impede a atualização dos demais dados da OP
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                        
        End Select

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 193804, 193805, 195482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193806)

    End Select

    Exit Sub

End Sub

Private Sub BotaoPecas_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoPecas_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridPecasPrev.Row = 0 Then gError 195320

    'Verifica se o Produto está preenchido
    If Len(Trim(GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_CodigoPecaPrev_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_CodigoPecaPrev_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 195321
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza<>?"

    'Chama a tela de browse
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoPeca, sSelecaoSQL)
        
        
    Exit Sub

Erro_BotaoPecas_Click:

    Select Case gErr
    
        Case 195320
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195321
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195322)
    
    End Select
    
    Exit Sub

End Sub


Private Sub BotaoPedSRV_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoPedSRV_Click

    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(PedSRV.Text)) = 0 Then gError 195721
    If Len(Trim(FilialPedSRV.Text)) = 0 Then gError 195722

    objPedidoVenda.lCodigo = StrParaLong(PedSRV.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(FilialPedSRV.Text)

    'Lê o Pedido de Servico a partir do Código passado (incluindo tributacao, itens de pedido, parcelas, Comissoes, bloqueios)
    lErro = CF("PedidoServico_Le_Completo", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 193245 Then gError 195723

    If lErro <> SUCESSO Then gError 195724

    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 195725

    Exit Sub

Erro_BotaoPedSRV_Click:

    Select Case gErr

        Case 195721
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSRV_NAO_INFORMADO", gErr)

        Case 195722
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)

        Case 195723, 195725

        Case 195724
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSRV_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195726)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVerPedSRV_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoVerPedSRV_Click

    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(PedSRV.Text)) = 0 Then gError 195746
    If Len(Trim(FilialPedSRV.Text)) = 0 Then gError 195747

    objPedidoVenda.lCodigo = StrParaLong(PedSRV.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(FilialPedSRV.Text)

    'Lê o Pedido de Servico a partir do Código passado (incluindo tributacao, itens de pedido, parcelas, Comissoes, bloqueios)
    lErro = CF("PedidoServico_Le_Completo", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 193245 Then gError 195748

    If lErro <> SUCESSO Then gError 195749
    
    'Chama  a tela de Pedido de Venda passando o pedido de venda da tela
    Call Chama_Tela("PedidoServico", objPedidoVenda)
    
    Exit Sub

Erro_BotaoVerPedSRV_Click:

    Select Case gErr
    
        Case 195746
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSRV_NAO_INFORMADO", gErr)

        Case 195747
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)

        Case 195748

        Case 195749
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSRV_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195750)

    End Select

    Exit Sub

End Sub

Private Sub FilialOP_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialOP_Validate

    'Se não estiver preenchida ou alterada pula a crítica
    If Len(Trim(FilialOP.Text)) = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialOP, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195736

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 195737

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 195738

    Exit Sub

Erro_FilialOP_Validate:

    Cancel = True

    Select Case gErr

        Case 195736

        Case 195737
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 195738
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195739)

    End Select

    Exit Sub

End Sub

Public Sub FilialOP_Click()

Dim lErro As Long

On Error GoTo Erro_FilialOP_Click

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_FilialOP_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195740)

    End Select

    Exit Sub

End Sub

Private Sub FilialPedSRV_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialPedSRV_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialPedSRV_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialPedSRV_Validate

    'Se não estiver preenchida, sai.
    If Len(Trim(FilialPedSRV.Text)) = 0 Then Exit Sub
    'Se não foi selecoinada, sai.
    If FilialPedSRV.ListIndex <> -1 Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialPedSRV, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195751
    
    If lErro <> SUCESSO Then gError 195752
    
    Exit Sub

Erro_FilialPedSRV_Validate:

    Cancel = True

    Select Case gErr

        Case 195751

        Case 195752
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialPedSRV.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195753)

    End Select

    Exit Sub

End Sub

Private Sub GridMaquinasPrev_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridMaquinasPrev)

End Sub

Private Sub GridMOPrev_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridMOPrev)

End Sub

Private Sub GridPecasPrev_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridPecasPrev)

End Sub

Private Sub LabelCodigo_Click()

Dim objSolicSRV As New ClassSolicSRV
Dim colSelecao As New Collection

    objSolicSRV.lCodigo = StrParaLong(CodigoSolSrv.Text)
    
    Call Chama_Tela("SolicitacaoSRVLista", colSelecao, objSolicSRV, objEventoSolSRV)
    
End Sub

Private Sub Lote_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LoteLabel_Click()

Dim colSelecao As New Collection
Dim objRastroLote As New ClassRastreamentoLote
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sSelecao As String
Dim lErro As Long

On Error GoTo Erro_LoteLabel_Click

    objRastroLote.sCodigo = Lote.Text

    sProduto = ProdutoSolServ.Text

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195741

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195742

    'Selecao
    colSelecao.Add sProdutoFormatado

    sSelecao = "Produto = ?"

    'Chama tela de Browse de RastreamentoLote
    Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoLote, sSelecao)

    Exit Sub

Erro_LoteLabel_Click:

    Select Case gErr

        Case 195741

        Case 195742
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195743)

    End Select

    Exit Sub

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As ClassRastreamentoLote
Dim iCodigo As Integer

On Error GoTo Erro_objEventoLote_evSelecao

    Set objRastroLote = obj1

    Lote.Text = objRastroLote.sCodigo

    If objRastroLote.iFilialOP <> 0 Then
    
        'Tenta selecionar na combo
        lErro = Combo_Seleciona(FilialOP, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195744

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr

        Case 195744

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195745)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPeca_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoPeca_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridPecasPrev.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 195475

    CodigoPecaPrev.PromptInclude = False
    CodigoPecaPrev.Text = sProduto
    CodigoPecaPrev.PromptInclude = True

    If Not (Me.ActiveControl Is CodigoPecaPrev) Then

        GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_CodigoPecaPrev_Col) = CodigoPecaPrev.Text

        'Faz o Tratamento do produto
        lErro = CodigoPecaPrev_Saida_Celula()
        If lErro <> SUCESSO Then gError 195476

    End If

    Me.Show

    Exit Sub

Erro_objEventoPeca_evSelecao:

    Select Case gErr

        Case 195475
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 195476
            GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_CodigoPecaPrev_Col) = ""

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195477)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSolSRV_evSelecao(obj1 As Object)

Dim objSolicSRV As ClassSolicSRV

    Set objSolicSRV = obj1
    
    CodigoSolSrv.Text = objSolicSRV.lCodigo

End Sub

Private Sub Obs_Validate(Cancel As Boolean)

    If GridMovimentos.Row > 0 And GridMovimentos.Row <= objGrid.iLinhasExistentes Then
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ObsSRV_Col) = Obs.Text
    End If
    
End Sub

Private Sub PedSRV_Change()
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedSRV_GotFocus()
    Call MaskEdBox_TrataGotFocus(PedSRV, iAlterado)
End Sub

Private Sub PedSRVLabel_Click()

Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection

    'Chama a tela de browse dos pedidos que podem gerar OS
    Call Chama_Tela("PedSRV_LiberadosLista", colSelecao, objPedidoVenda, objEventoPedido)

End Sub

Private Sub objEventoPedido_evSelecao(obj1 As Object)

Dim objPedidoVenda As ClassPedidoDeVenda

    Set objPedidoVenda = obj1

    'Coloca o código dopedido na tela
    PedSRV.PromptInclude = False
    PedSRV.Text = objPedidoVenda.lCodigo
    'Coloca a filial do pedido na tela
    PedSRV.PromptInclude = True
    
    If objPedidoVenda.iFilialEmpresa <> 0 Then FilialPedSRV.Text = objPedidoVenda.iFilialEmpresa
    Call FilialPedSRV_Validate(bSGECancelDummy)
    
    'Traz os dados do pedido selecionado para a tela
    Call BotaoPedSRV_Click
    
    Me.Show

End Sub


'Private Sub BotaoAbrirOP_Click()
'
'Dim lErro As Long
'Dim objOrdemDeProducao As New ClassOrdemDeProducao
'
'On Error GoTo Erro_BotaoVerEtapas_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    lErro = Teste_Salva(Me, iAlterado)
'    If lErro <> SUCESSO Then gError 139048
'
'    objOrdemDeProducao.sCodigo = OPCodigoMRP
'    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
'
'    'traz OP para a tela
'    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
'    If lErro <> SUCESSO And lErro <> 21966 Then gError 139049
'
'    'Torna Frame atual invisível
'    Frame1(TabStrip1.SelectedItem.Index).Visible = False
'    iFrameAtual = 1
'    'Torna Frame atual visível
'    Frame1(iFrameAtual).Visible = True
'    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
'
'    'Torna Frame Operacoes atual invisível
'    Frame2(TabStrip2.SelectedItem.Index).Visible = False
'    iFrameAtualOper = 1
'    'Torna Frame Operacoes atual visível
'    Frame2(iFrameAtualOper).Visible = True
'    TabStrip2.Tabs.Item(iFrameAtualOper).Selected = True
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoVerEtapas_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 139048, 139049
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163801)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub BotaoCompetencia_Click()
'
'Dim lErro As Long
'Dim objCompetencias As New ClassCompetencias
'
'On Error GoTo Erro_BotaoVerEtapas_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    objCompetencias.sNomeReduzido = LabelCodigoCompetencia.Caption
'
'    'Lê a Competencia pelo NomeReduzido
'    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
'    If lErro <> SUCESSO And lErro <> 134937 Then gError 139060
'
'    'Chama a tela de ordem de produção
'    Call Chama_Tela("Competencias", objCompetencias)
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoVerEtapas_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 139060
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163802)
'
'    End Select
'
'    Exit Sub
'
'
'End Sub

'Private Sub BotaoCT_Click()
'
'Dim lErro As Long
'Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
'
'On Error GoTo Erro_BotaoVerEtapas_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    objCentrodeTrabalho.sNomeReduzido = LabelCodigoCTPadrao.Caption
'
'    'Chama a tela de ordem de produção
'    Call Chama_Tela("CentrodeTrabalho", objCentrodeTrabalho)
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoVerEtapas_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163803)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub BotaoGrade_Click()
'
'Dim lErro  As Long
'Dim objRomaneioGrade As ClassRomaneioGrade
'Dim objItemOP As ClassItemOP
'
'On Error GoTo Erro_BotaoGrade_Click
'
'    If GridMovimentos.Row > 0 And GridMovimentos.Row <= objGrid.iLinhasExistentes Then
'
'        Set objItemOP = gobjOP.colItens(GridMovimentos.Row)
'
'        If objItemOP.iPossuiGrade = DESMARCADO Then gError 126517
'
'        objItemOP.sAlmoxarifadoNomeRed = AlmoxPadrao.Text
'
'        Set objRomaneioGrade = New ClassRomaneioGrade
'
'        objRomaneioGrade.sNomeTela = Me.Name
'        Set objRomaneioGrade.objObjetoTela = objItemOP
'        Set objRomaneioGrade.objTela = Me
'
'        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
'
'        Call Atualiza_Grid_Movimentos(objItemOP)
'
'    End If
'
'    Exit Sub
'
'Erro_BotaoGrade_Click:
'
'    Select Case gErr
'
'        Case 126517
'            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_GRADE", gErr, GridMovimentos.Row)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163804)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub BotaoGrafico_Click()
'
'Dim lErro As Long
'Dim objTelaGrafico As New ClassTelaGrafico
'
'On Error GoTo Erro_BotaoGrafico_Click:
'
'    lErro = Atualiza_Cronograma(objTelaGrafico)
'    If lErro <> SUCESSO Then gError 139061
'
'    Call Chama_Tela_Nova_Instancia("TelaGrafico", objTelaGrafico)
'
'    Exit Sub
'
'Erro_BotaoGrafico_Click:
'
'    Select Case gErr
'
'        Case 139061
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163805)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub BotaoMaquinaPMP_Click()
'
'Dim lErro As Long
'Dim objMaquinas As New ClassMaquinas
'
'On Error GoTo Erro_BotaoVerEtapas_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Se não tiver linha selecionada => Erro
'    If GridMaquinas.Row = 0 Then gError 139058
'
'    'Verifica se a linha selecionada está preenchida
'    If Len(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_NomeMaquina_Col)) = 0 Then gError 139059
'
'    objMaquinas.sNomeReduzido = GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_NomeMaquina_Col)
'
'    'Chama a tela de ordem de produção
'    Call Chama_Tela("Maquinas", objMaquinas)
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoVerEtapas_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 139058
'            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case 139059
'            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163806)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub BotaoMaquinas_Click()
''Chama o Browser de Maquinas...
'
'Dim lErro As Long
'Dim objMaquina As New ClassMaquinas
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim colSelecao As Collection
'
'On Error GoTo Erro_BotaoMaquinas_Click
'
'    'Verifica se tem alguma linha selecionada no Grid
'    If GridMovimentos.Row = 0 Then gError 106320
'
'    'Se o equipamento foi preenchido => armazena no obj
'    If Len(Trim(Maquina.Text)) > 0 Then
'
'        If IsNumeric(Maquina.Text) Then
'
'            objMaquina.iCodigo = StrParaInt(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col))
'
'        Else
'
'            objMaquina.sNomeReduzido = CStr(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col))
'
'        End If
'
'    End If
'
'    'Lista de Equipamentos
'    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquina)
'
'    Exit Sub
'
'Erro_BotaoMaquinas_Click:
'
'    Select Case gErr
'
'        Case 106320
'            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case 106433
'            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NOME_INEXISTENTE", gErr, Maquina.Text)
'
'        Case 106434
'            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_CODIGO_INEXISTENTE", gErr, Maquina.Text, giFilialEmpresa)
'
'        Case 55325, 106432
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163807)
'
'    End Select
'
'End Sub

Private Sub Prioridade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Prioridade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoLabel1_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sSelecaoSQL As String
Dim objProdutoSolicSRV As New ClassProdutoSolicSRV


On Error GoTo Erro_ProdutoLabel1_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(ProdutoSolServ.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoSolServ.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195374

        objProduto.sCodigo = sProdutoFormatado
        objProdutoSolicSRV.sCodigo = sProdutoFormatado

    End If

    If Len(Trim(CodigoSolSrv.Text)) <> 0 Then
    
        colSelecao.Add StrParaLong(CodigoSolSrv.Text)

        sSelecaoSQL = "CodSolic= ?"
    
        Call Chama_Tela("ProdutosSolicSRVLista", colSelecao, objProdutoSolicSRV, objEventoProdutoSolicSRV, sSelecaoSQL)
    
    Else
    
        Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto, sSelecaoSQL)
    
    End If



    Exit Sub

Erro_ProdutoLabel1_Click:

    Select Case gErr

        Case 195374

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195375)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoSolicSRV_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProdutoSolicSRV As ClassProdutoSolicSRV
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoProdutoSolicSRV_evSelecao

    Set objProdutoSolicSRV = obj1

    objProduto.sCodigo = objProdutoSolicSRV.sCodigo

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 195376

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 195377

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoSolServ, DescProd)
    If lErro <> SUCESSO Then gError 195378

    Lote.Text = objProdutoSolicSRV.sLote
    
    If objProdutoSolicSRV.iFilialOP <> 0 Then FilialOP.Text = objProdutoSolicSRV.iFilialOP

    DescProd.Caption = objProdutoSolicSRV.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoProdutoSolicSRV_evSelecao:

    Select Case gErr

        Case 195376, 195378

        Case 195377
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195379)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 195376

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 195377

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoSolServ, DescProd)
    If lErro <> SUCESSO Then gError 195378

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 195376, 195378

        Case 195377
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195379)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoSolServ_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoSolServ_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_ProdutoSolServ_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoSolServ, DescProd)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 202440
    
    If lErro <> SUCESSO Then gError 202441

    Exit Sub

Erro_ProdutoSolServ_Validate:

    Cancel = True

    Select Case gErr

        Case 202440

        Case 202441
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202442)

    End Select

    Exit Sub
End Sub

Private Sub Servico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Servico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Servico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Servico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Servico
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Versao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Versao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Versao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Versao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Versao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO
    iQtdeAlterada = REGISTRO_ALTERADO
    
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

Private Sub DataPrevisaoFim_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub DataPrevisaoFim_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DataPrevisaoFim_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DataPrevisaoFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DataPrevisaoFim
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataPrevisaoInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataPrevisaoInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DataPrevisaoInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DataPrevisaoInicio_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = DataPrevisaoInicio
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub StatusItem_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub StatusItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub StatusItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub StatusItem_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = StatusItem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridMovimentos_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGrid, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGrid, iAlterado)
        End If
End Sub

Private Sub GridMovimentos_GotFocus()
    Call Grid_Recebe_Foco(objGrid)
End Sub

Private Sub GridMovimentos_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridMovimentos_LeaveCell()
    Call Saida_Celula(objGrid)
End Sub

Private Sub GridMovimentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridMovimentos_RowColChange()

Dim lErro As Long

On Error GoTo Erro_GridMovimentos_RowColChange

    Call Grid_RowColChange(objGrid)
    
    If GridMovimentos.Row = 0 Then Exit Sub
    
    Obs.Text = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ObsSRV_Col)
    
    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 193937)

    End Select

    Exit Sub

End Sub

Private Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub GridMovimentos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer
Dim lErro As Long
Dim iLinhasExistentes As Integer 'm

Dim sServicoRaiz As String
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer
Dim objItOScol As New ClassItemOS

On Error GoTo Erro_GridMovimentos_KeyDown

    'Guarda iLinhasExistentes
    iLinhasExistentesAnterior = objGrid.iLinhasExistentes

    'Verifica se a Tecla apertada foi Del
    If KeyCode = vbKeyDelete Then

        'Guarda o índice da Linha a ser Excluída
        iLinhaAnterior = GridMovimentos.Row
        
        sServicoRaiz = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col)
        
    End If

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    'Verifica se a Linha foi realmente excluída
    If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then

        gcolItemOS.Remove (iLinhaAnterior)
        gobjOS.colItens.Remove iLinhaAnterior
        
        lErro = CF("Produto_Formata", sServicoRaiz, sServicoFormatado, iServicoPreenchido)
        If lErro <> SUCESSO Then gError 193938
                    
        'ajusta a coleção
        For Each objItOScol In gcolOSItens
            If objItOScol.sServico = sServicoFormatado Then
                gcolOSItens.Remove (sServicoFormatado)
                Exit For
            End If
        Next
        '---------------------------------------

        For iLinhasExistentes = 1 To objGrid.iLinhasExistentes 'm
'            If gobjOS.colItens(iLinhasExistentes).iPossuiGrade = MARCADO Then
'                GridMovimentos.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
'            Else
                GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
'            End If
            
        Next

        GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes

    End If

    Exit Sub
    
Erro_GridMovimentos_KeyDown:

    Select Case gErr
    
        Case 193938
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 193939)
    
    End Select
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridMovimentos
        If objGridInt.objGrid.Name = GridMovimentos.Name Then

            Select Case GridMovimentos.Col
    
                Case iGrid_Servico_Col
    
                    lErro = Saida_Celula_Servico(objGridInt)
                    If lErro <> SUCESSO Then gError 193940
                    
                Case iGrid_Versao_Col
                    lErro = Saida_Celula_Versao(objGridInt)
                    If lErro <> SUCESSO Then gError 195217
                    
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 193942
                
                Case iGrid_Ccl_Col
    
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 193943
    
                Case iGrid_DataPrevInicio_Col
    
                    lErro = Saida_Celula_DataPrevInicio(objGridInt)
                    If lErro <> SUCESSO Then gError 193944
    
                Case iGrid_DataPrevFim_Col
    
                    lErro = Saida_Celula_DataPrevFim(objGridInt)
                    If lErro <> SUCESSO Then gError 193945
    
                Case iGrid_StatusItem_Col
    
                    lErro = Saida_Celula_StatusItem(objGridInt)
                    If lErro <> SUCESSO Then gError 193946
    
                Case iGrid_Prioridade_Col
    
                    lErro = Saida_Celula_Prioridade(objGridInt)
                    If lErro <> SUCESSO Then gError 193947
    
            End Select
                   
        End If
        
        
        If objGridInt.objGrid.Name = GridPecasPrev.Name Then
        
            Select Case GridPecasPrev.Col
        
                Case iGrid_CodigoPecaPrev_Col
        
                    lErro = Saida_Celula_CodigoPecaPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195177
        
                Case iGrid_QuantidadePecaPrev_Col
        
                    lErro = Saida_Celula_QuantidadePecaPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195177
        
                Case iGrid_VersaoPecaPrev_Col
        
                    lErro = Saida_Celula_VersaoPecaPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195481
        
            End Select
        
        End If
        
        If objGridInt.objGrid.Name = GridMaquinasPrev.Name Then

            Select Case GridMaquinasPrev.Col

                Case iGrid_NomeRedMaquinaPrev_Col

                    lErro = Saida_Celula_NomeRedMaquinaPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195177

                Case iGrid_QuantMaquinaPrev_Col

                    lErro = Saida_Celula_QuantMaquinaPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195178

                Case iGrid_HorasMaquinaPrev_Col

                    lErro = Saida_Celula_HorasMaquinaPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195179

            End Select

        End If

        If objGridInt.objGrid.Name = GridMOPrev.Name Then

            Select Case GridMOPrev.Col

                Case iGrid_HorasMOPrev_Col

                    lErro = Saida_Celula_HorasMOPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195461

                Case iGrid_CodigoMOPrev_Col

                    lErro = Saida_Celula_CodigoMOPrev(objGridInt)
                    If lErro <> SUCESSO Then gError 195462

            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 193948

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 193940 To 193947, 195177 To 195179, 195217, 195461, 195462, 195481

        Case 193948
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193949)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Servico(objGridInt As AdmGrid) As Long
'faz a critica da celula de Servico do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sServico As String
Dim iServicoPreenchido As Integer
Dim sServicoFormatado As String
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objMaquina As New ClassMaquinas
Dim objKit As New ClassKit
Dim objProdutoFilial As New ClassProdutoFilial
Dim iPossuiGrade As Integer
Dim iIndice As Integer
Dim objItemOS As New ClassItemOS
Dim sProdutoMascarado As String

On Error GoTo Erro_Saida_Celula_Servico

    Set objGridInt.objControle = Servico

    lErro = CF("Produto_Formata", Servico.Text, sServicoFormatado, iServicoPreenchido)
    If lErro <> SUCESSO Then gError 193950
    
    'se o servico foi preenchido
    If Len(Trim(Servico.ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica2", Servico.Text, objProduto, iServicoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 193951
        
        'se produto estiver preenchido
        If iServicoPreenchido = PRODUTO_PREENCHIDO Then
                
            'se é um produto gerencial ==> erro
            If lErro = 25043 Then gError 193952
            
            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 193953
    
            'verifica se este servico já foi usado na OS
            lErro = VerificaUso_Produto(objProduto)
            If lErro <> SUCESSO And lErro <> 193956 Then gError 193954

            If lErro <> SUCESSO Then gError 193958

'            'Verifica se é um kit
'            objKit.sProdutoRaiz = sProdutoFormatado
'            lErro = CF("Kit_Le_Padrao", objKit)
'            If lErro <> SUCESSO And lErro <> 106304 Then gError 106430
'
'            'Se encontrou => É UM KIT => Carrega a Combo com as Versoes
'            If lErro <> 106304 Then Call Carrega_ComboVersoes(objProduto.sCodigo)

            Call Carrega_ComboVersoes(objProduto.sCodigo)

            'se o produto nao controla estoque ==> erro
            If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 193989
    
            If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 193990

            'Preenche a linha do grid
            lErro = ProdutoLinha_Preenche(objProduto, objItemOS)
            If lErro <> SUCESSO Then gError 193992
        
            'mascara produto escolhido
            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 195380

            Servico.PromptInclude = False
            Servico.Text = sProdutoMascarado
            Servico.PromptInclude = True

        End If

        If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) <> 0 Then

            lErro = Trata_Arvore(Servico.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)))
            If lErro <> SUCESSO Then gError 195383

        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 193993


    
    Saida_Celula_Servico = SUCESSO

    Exit Function

Erro_Saida_Celula_Servico:

    Saida_Celula_Servico = gErr

    Select Case gErr

        Case 193950, 193951, 193954, 193992, 193993, 195380, 195383
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 193952
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 193953
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERVICO", Servico.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = Servico.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 193958
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, Servico.Text, CodigoOS.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 193989
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_NATUREZA_SERVICO", gErr, objProduto.sCodigo)
        
        Case 193990
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, Servico.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193994)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Versao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Versao do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Saida_Celula_Versao

    Set objGridInt.objControle = Versao

    If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) <> 0 Then

        lErro = Trata_Arvore(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), Versao.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError 195381

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195218

    Saida_Celula_Versao = SUCESSO

    Exit Function

Erro_Saida_Celula_Versao:

    Saida_Celula_Versao = gErr

    Select Case gErr

        Case 195218, 195381
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195219)

    End Select

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    
    'se a quantidade foi preenchida
    If Len(Quantidade.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 193995
    
        Quantidade.Text = Formata_Estoque(Quantidade.Text)
        
        lErro = Trata_Arvore(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col), StrParaDbl(Quantidade.Text))
        If lErro <> SUCESSO Then gError 195382
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 193996

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 193995, 193996, 195382
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193997)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Ccl.ClipText) > 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 193998

        If lErro = 5703 Then gError 193999

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195000

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 193998, 195000
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 193999
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            
            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195001)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevInicio

    Set objGridInt.objControle = DataPrevisaoInicio

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoInicio.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoInicio.Text)
        If lErro <> SUCESSO Then gError 195002

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195003

    Saida_Celula_DataPrevInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevInicio:

    Saida_Celula_DataPrevInicio = gErr

    Select Case gErr

        Case 195002, 195003
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195004)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevFim(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevFim

    Set objGridInt.objControle = DataPrevisaoFim

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoFim.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoFim.Text)
        If lErro <> SUCESSO Then gError 195005

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195006

    Saida_Celula_DataPrevFim = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevFim:

    Saida_Celula_DataPrevFim = gErr

    Select Case gErr

        Case 195005, 195006
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195007)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Prioridade(objGridInt As AdmGrid) As Long
'faz a critica da celula de Prioridade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Prioridade

    Set objGridInt.objControle = Prioridade

    If Len(Prioridade.Text) > 0 Then
        lErro = Inteiro_Critica(Prioridade.Text)
        If lErro <> SUCESSO Then gError 195369
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195008

    Saida_Celula_Prioridade = SUCESSO

    Exit Function

Erro_Saida_Celula_Prioridade:

    Saida_Celula_Prioridade = gErr

    Select Case gErr

        Case 195008, 195369
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195009)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_StatusItem(objGridInt As AdmGrid) As Long
'faz a critica da celula de StatusItem do grid que está deixando de ser a corrente
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_StatusItem

    Set objGridInt.objControle = StatusItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195010

    Saida_Celula_StatusItem = SUCESSO

    Exit Function

Erro_Saida_Celula_StatusItem:

    Saida_Celula_StatusItem = gErr

    Select Case gErr

        Case 195010
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195011)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodigoPecaPrev(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto Data que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_CodigoPecaPrev

    Set objGridInt.objControle = CodigoPecaPrev

    If Len(Trim(CodigoPecaPrev.ClipText)) > 0 Then

        lErro = CodigoPecaPrev_Saida_Celula()
        If lErro <> SUCESSO Then gError 195462
        
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_CodigoPecaPrev_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195463

    Saida_Celula_CodigoPecaPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigoPecaPrev:

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_CodigoPecaPrev_Col) = ""

    Saida_Celula_CodigoPecaPrev = gErr

    Select Case gErr

        Case 195462, 195463
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195464)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadePecaPrev(objGridInt As AdmGrid) As Long
'faz a critica da celula de QuantidadeMaquina do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim dHoras As Double
Dim iQtd As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As New ClassMaquinas
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_QuantidadePecaPrev

    Set objGridInt.objControle = QuantidadePecaPrev

    'verifica se a QuantidadePecaPrev está preenchida
    If Len(Trim(QuantidadePecaPrev.Text)) > 0 Then

        'verifica se a QuantidadeMaquina é válida
        lErro = Valor_Positivo_Critica(QuantidadePecaPrev.Text)
        If lErro <> SUCESSO Then gError 195472
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195473

    Saida_Celula_QuantidadePecaPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadePecaPrev:

    Saida_Celula_QuantidadePecaPrev = gErr

    Select Case gErr

        Case 195472, 195473
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195474)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VersaoPecaPrev(objGridInt As AdmGrid) As Long
'faz a critica da celula de QuantidadeMaquina do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim dHoras As Double
Dim iQtd As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As New ClassMaquinas
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_VersaoPecaPrev

    Set objGridInt.objControle = VersaoPecaPrev

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195473

    Saida_Celula_VersaoPecaPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoPecaPrev:

    Saida_Celula_VersaoPecaPrev = gErr

    Select Case gErr

        Case 195472, 195473
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195474)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function CodigoPecaPrev_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CodigoPecaPrev_Saida_Celula

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", CodigoPecaPrev.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 195465
    
    'Se o produto é gerencial  ==> Erro
    If lErro = 86295 Then
        gError 195466
    End If

    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 195467

    If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then gError 195468

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 195469

        CodigoPecaPrev.PromptInclude = False
        CodigoPecaPrev.Text = sProduto
        CodigoPecaPrev.PromptInclude = True

    End If

    For iIndice = 1 To objGridPecasPrev.iLinhasExistentes
        If iIndice <> GridPecasPrev.Row Then
            If GridPecasPrev.TextMatrix(iIndice, iGrid_CodigoPecaPrev_Col) = CodigoPecaPrev.Text Then gError 195470
        End If
    Next

    'Unidade de Medida
    GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_UMPecaPrev_Col) = objProduto.sSiglaUMEstoque

    'Descricao Produto
    GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_DescricaoPecaPrev_Col) = objProduto.sDescricao


    If objProduto.iCompras = PRODUTO_COMPRAVEL Then
        GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_OrigemPecaPrev_Col) = INSUMO_COMPRADO
    Else
        GridPecasPrev.TextMatrix(GridPecasPrev.Row, iGrid_OrigemPecaPrev_Col) = INSUMO_PRODUZIDO
    End If

    'Acrescenta uma linha no Grid se for o caso
    If GridPecasPrev.Row - GridPecasPrev.FixedRows = objGridPecasPrev.iLinhasExistentes Then
        objGridPecasPrev.iLinhasExistentes = objGridPecasPrev.iLinhasExistentes + 1
    End If

    CodigoPecaPrev_Saida_Celula = SUCESSO

    Exit Function

Erro_CodigoPecaPrev_Saida_Celula:

    CodigoPecaPrev_Saida_Celula = gErr

    Select Case gErr

        Case 195465

        Case 195466
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 195467
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", CodigoPecaPrev.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridPecasPrev)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridPecasPrev)
            End If

        Case 195468
            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_NAO_PODE_TER_NATUREZA_SERVICO", gErr, objProduto.sCodigo)

        Case 195469
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, CodigoPecaPrev.Text)

        Case 195470
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO_GRID", gErr, CodigoPecaPrev.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195471)

    End Select

    Exit Function

End Function


Public Sub Form_Activate()

    'Carrega índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Function CargaCombo_StatusItem(objStatusItem As Object)
'Carga dos itens da combo Situação

Dim lErro As Long
Dim bSelecionaPadrao As Boolean

On Error GoTo Erro_CargaCombo_StatusItem

    bSelecionaPadrao = False

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_OS_STATUSITEM, objStatusItem, bSelecionaPadrao, False)
    If lErro <> SUCESSO Then gError 195012

    CargaCombo_StatusItem = SUCESSO

    Exit Function

Erro_CargaCombo_StatusItem:

    CargaCombo_StatusItem = gErr

    Select Case gErr
    
        Case 195012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195013)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim sItem As String
Dim sMascaraCclPadrao As String
Dim objFiliais As AdmFiliais
Dim colModulo As New AdmColModulo
Dim tModulo As typeModulo
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set gcolItemOS = New Collection

    Set objEventoCodigo = New AdmEvento
    Set objEventoServico = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoSolSRV = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoProdutoSolicSRV = New AdmEvento
    Set objEventoMaquinas = New AdmEvento
    Set objEventoMO = New AdmEvento
    Set objEventoPeca = New AdmEvento
    Set objEventoLote = New AdmEvento
    Set objEventoPedido = New AdmEvento
    Set objEventoCompetencias = New AdmEvento
    Set objEventoCentroDeTrabalho = New AdmEvento

    Set gcolOSItens = New Collection
    
    'Carrega Ítens das Combos
    lErro = CargaCombo_StatusItem(StatusItem)
    If lErro <> SUCESSO Then gError 195014

    lErro = Carrega_Tipo(Tipo)
    If lErro <> SUCESSO Then gError 195014

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Servico)
    If lErro <> SUCESSO Then gError 195015

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoSolServ)
    If lErro <> SUCESSO Then gError 195373

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoPecaPrev)
    If lErro <> SUCESSO Then gError 195480

    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 195733

    iFrameAtualOper = 1
    
    'Grid de OperacaoInsumos
    Set objGridPecasPrev = New AdmGrid
    
    'tela em questão
    Set objGridPecasPrev.objForm = Me
    
    lErro = Inicializa_GridPecasPrev(objGridPecasPrev)
    If lErro <> SUCESSO Then gError 195016
    
    Call Habilita_Operacoes(Roteiro.Nodes.Count)
    
    iQtdeAlterada = 0
    
    'Grid de uso de Maquinas real
    Set objGridMaquinasPrev = New AdmGrid
    
    'tela em questão
    Set objGridMaquinasPrev.objForm = Me
    
    lErro = Inicializa_GridMaquinasPrev(objGridMaquinasPrev)
    If lErro <> SUCESSO Then gError 195017
    
    'Grid de Mão de Obra
    Set objGridMOPrev = New AdmGrid
    
    'tela em questão
    Set objGridMOPrev.objForm = Me
    
    lErro = Inicializa_GridMOPrev(objGridMOPrev)
    If lErro <> SUCESSO Then gError 195018
    
    HorasMaquinaPrev.Format = FORMATO_ESTOQUE
    HorasMOPrev.Format = FORMATO_ESTOQUE
    
    QuantidadePecaPrev.Format = FORMATO_ESTOQUE
    Quantidade.Format = FORMATO_ESTOQUE

    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    lErro = Carrega_FilialPedSRV()
    If lErro <> SUCESSO Then gError 195754

    'inicializa Grid
    lErro = Inicializa_GridMovimentos
    If lErro <> SUCESSO Then gError 195019

    Set gobjOS = New ClassOS
    
    iAlterado = 0
    iCodigoAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 195014 To 195019, 195373, 195480, 195754

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195020)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Set objEventoServico = Nothing
    Set objEventoCcl = Nothing
    Set objEventoSolSRV = Nothing
    Set objEventoProduto = Nothing
    Set objEventoProdutoSolicSRV = Nothing
    Set objEventoMaquinas = Nothing
    Set objEventoMO = Nothing
    Set objEventoPeca = Nothing
    Set objEventoLote = Nothing
    Set objEventoPedido = Nothing

    Set gcolItemOS = Nothing
    Set gobjOS = Nothing
    Set gobjItemOS = Nothing
    
    Set objGrid = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    If lErro <> SUCESSO Then gError 195021

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case 195021

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195022)

    End Select

    Exit Sub

End Sub

Private Function VerificaUso_Produto(ByVal objProduto As ClassProduto) As Long
'Verifica se existem produtos repetidos na OS

Dim lErro As Long
Dim iIndice As Integer
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer

On Error GoTo Erro_VerificaUso_Produto

    If objGrid.iLinhasExistentes > 0 Then

        For iIndice = 1 To objGrid.iLinhasExistentes

            If GridMovimentos.Row <> iIndice Then

                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
                If lErro <> SUCESSO Then gError 193955

                If sServicoFormatado = objProduto.sCodigo Then gError 193956
                
            End If

        Next

    End If

    VerificaUso_Produto = SUCESSO

    Exit Function

Erro_VerificaUso_Produto:

    VerificaUso_Produto = gErr

    Select Case gErr

        Case 193955, 193956

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193957)

    End Select

End Function

Sub Limpa_GridMovimentos()

    Call Grid_Limpa(objGrid)

End Sub

Function Preenche_GridMovimentos(colItensOS As Collection) As Long
'preenche o grid com os dados contidos na coleção colItensOS

Dim lErro As Long, sCclMascarado As String
Dim iIndice As Integer, iIndice1 As Integer, sProdutoMascarado As String
Dim objItemOS As New ClassItemOS, objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridMovimentos

    'Remove os ítens de gcolItemOS
    Set gcolItemOS = New Collection
    
    'Remove os itens de gcolOSItens
    Set gcolOSItens = New Collection

    iIndice = 1

    'preenche o grid com os dados retornados na coleção colItensOS
    For Each objItemOS In colItensOS

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_RetornaProdutoTela(objItemOS.sServico, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 195028

        Servico.PromptInclude = False
        Servico.Text = sProdutoMascarado
        Servico.PromptInclude = True

        GridMovimentos.TextMatrix(iIndice, iGrid_Servico_Col) = sProdutoMascarado

        'le o produto para obter sua descricao
        objProduto.sCodigo = objItemOS.sServico
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195029
        
        'Se não encontrar o produto ==> Erro
        If lErro <> SUCESSO Then gError 195030

        GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objProduto.sDescricao
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemOS.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemOS.dQuantidade)

        'mascara Ccl , se estiver informada
        If objItemOS.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemOS.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 195031

        Else
            sCclMascarado = ""
        End If

        GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        'preenche datas
        If objItemOS.dtDataInicio <> DATA_NULA Then
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col) = Format(objItemOS.dtDataInicio, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col) = ""
        End If

        If objItemOS.dtDataFim <> DATA_NULA Then
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col) = Format(objItemOS.dtDataFim, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col) = ""
        End If

        'preenche StatusItem
        For iIndice1 = 0 To StatusItem.ListCount - 1
            If StatusItem.ItemData(iIndice1) = objItemOS.iStatusItem Then
                StatusItem.ListIndex = iIndice1
                Exit For
            End If
        Next
            
        GridMovimentos.TextMatrix(iIndice, iGrid_StatusItem_Col) = StatusItem.Text
        
        'preenche prioridade
        GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col) = CStr(objItemOS.iPrioridade)
        
        'Preenche com a Versão do Roteiro
        GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col) = objItemOS.sVersao
        
        GridMovimentos.TextMatrix(iIndice, iGrid_NumIntDoc_Col) = objItemOS.lNumIntDoc
        
        GridMovimentos.TextMatrix(iIndice, iGrid_ObsSRV_Col) = objItemOS.sObservacao
        
        'adiciona item à coleção
        gcolItemOS.Add objItemOS.iItem
        gobjOS.colItens.Add objItemOS

        'Inclui itens de colOPItens
        gcolOSItens.Add objItemOS, objItemOS.sServico

        iIndice = iIndice + 1

    Next

    objGrid.iLinhasExistentes = colItensOS.Count
    
    Call Grid_Refresh_Checkbox(objGrid)
    
    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 195028
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
        
        Case 195029, 195031
        
        Case 195030
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195032)

    End Select

    Exit Function

End Function

Function Move_Grid_Memoria(objOS As ClassOS) As Long
'move itens do Grid para objOrdemDeProducao

Dim lErro As Long
Dim iIndice As Integer, iCount As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCcl As String, sCclFormatada As String, iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim objItemOS As ClassItemOS
Dim sStatusItem As String
Dim objItemOS1 As ClassItemOS
Dim objItemOSOperacoes As ClassItemOSOperacoes

On Error GoTo Erro_Move_Grid_Memoria

    objOS.iNumItens = 0
    objOS.iNumItensBaixados = 0

    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objItemOS = New ClassItemOS

        If gobjOS.colItens.Count >= iIndice Then

            objItemOS.lNumIntDoc = gobjOS.colItens(iIndice).lNumIntDoc
            Set objItemOS.colItemOSOperacoes = gobjOS.colItens(iIndice).colItemOSOperacoes
        
        End If

        objItemOS.sCodigo = objOS.sCodigo
        objItemOS.iFilialEmpresa = objOS.iFilialEmpresa

        sProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Servico_Col)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195047

        objItemOS.sServico = sProdutoFormatado
        objItemOS.sSiglaUM = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemOS.sVersao = GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col)
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
            objItemOS.dQuantidade = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        Else
            objItemOS.dQuantidade = 0
        End If

        sCcl = GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col)

        If Len(Trim(sCcl)) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 195048

        Else
            sCclFormatada = ""
        End If

        objItemOS.sCcl = sCclFormatada

        If Len(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col)) > 0 Then
            objItemOS.dtDataInicio = CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col))
        Else
            objItemOS.dtDataInicio = DATA_NULA
        End If

        If Len(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col)) > 0 Then
            objItemOS.dtDataFim = CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col))
        Else
            objItemOS.dtDataFim = DATA_NULA
        End If

        'Seleciona o status
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_StatusItem_Col))) > 0 Then
            sStatusItem = GridMovimentos.TextMatrix(iIndice, iGrid_StatusItem_Col)
            For iCount = 0 To StatusItem.ListCount - 1
                If StatusItem.List(iCount) = sStatusItem Then
                    objItemOS.iStatusItem = StatusItem.ItemData(iCount)
                    Exit For
                End If
            Next
        End If


        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col))) > 0 Then objItemOS.iPrioridade = CInt(GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col))

        objItemOS.lNumIntDoc = StrParaLong(GridMovimentos.TextMatrix(iIndice, iGrid_NumIntDoc_Col))
        objItemOS.sObservacao = GridMovimentos.TextMatrix(iIndice, iGrid_ObsSRV_Col)
        objItemOS.iItem = iIndice
        objOS.colItens.Add objItemOS
        objOS.iNumItens = objOS.iNumItens + 1
        If objItemOS.iStatusItem = STATUSITEMOS_BAIXADA Then objOS.iNumItensBaixados = objOS.iNumItensBaixados + 1

    Next

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = gErr

    Select Case gErr

        Case 195047, 195048

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195049)

    End Select

    Exit Function

End Function

Function Limpa_Tela_OS(Optional iFechaSetas As Integer = FECHAR_SETAS) As Long
'Limpa a Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_OS

    If iFechaSetas = FECHAR_SETAS Then
    'Fecha o comando das setas se estiver aberto
     lErro = ComandoSeta_Fechar(Me.Name)
     If lErro <> SUCESSO Then gError 21801
    End If
    Call Limpa_Tela(Me)

    FilialOP.ListIndex = -1

    StatusOS.Caption = ""
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Call Limpa_GridMovimentos

    Set gcolItemOS = New Collection

    Set gobjOS = New ClassOS

    Tipo.ListIndex = glTipoPadrao
    
    ImprimeAoGravar.Value = DESMARCADO
    
    Set gcolOSItens = New Collection
    Set gobjItemOS = Nothing
    
    Call Limpa_Operacoes
        
    Roteiro.Nodes.Clear
        
    Call Habilita_Operacoes(Roteiro.Nodes.Count)
    
    iQtdeAlterada = 0
    
    iAlterado = 0
    iCodigoAlterado = 0

    Obs.Text = ""

    FilialPedSRV.Text = giFilialEmpresa
    
    Call FilialPedSRV_Validate(bSGECancelDummy)

    Limpa_Tela_OS = SUCESSO

    Exit Function

Erro_Limpa_Tela_OS:

    Limpa_Tela_OS = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195048)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOS As ClassOS) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Trata_Parametros

    If Not (objOS Is Nothing) Then

        'traz OP para a tela
        lErro = Traz_Tela_OS(objOS)
        If lErro <> SUCESSO And lErro <> 195054 Then gError 195049

        If lErro = 21966 Then

            'Se não existe exibe apenas o código
            CodigoOS.Text = objOS.sCodigo

        End If
                
        Call ComandoSeta_Fechar(Me.Name)
                
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 195049

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195050)

    End Select

    iAlterado = 0

    Exit Function

End Function


'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objOS As New ClassOS

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OS"

    objOS.sCodigo = CodigoOS.Text
    objOS.dtDataEmissao = CDate(Data.Text)
    objOS.iFilialEmpresa = giFilialEmpresa
    
    If IsNumeric(CodigoOS.Text) Then
        objOS.lCodigoNumerico = StrParaLong(CodigoOS.Text)
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objOS.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "CodigoNumerico", objOS.lCodigoNumerico, 0, "CodigoNumerico"
    colCampoValor.Add "Codigo", objOS.sCodigo, STRING_OS_CODIGO, "Codigo"
    colCampoValor.Add "DataEmissao", objOS.dtDataEmissao, 0, "DataEmissao"
    
    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195051)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objOS As ClassOS) As Long

Dim lErro As Long, objPrestServ As New ClassPrestServ
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_Move_Tela_Memoria

    objOS.sCodigo = CodigoOS.Text
    objOS.dtDataEmissao = CDate(Data.Text)
    objOS.iFilialEmpresa = giFilialEmpresa
    objOS.lCodSolSRV = StrParaLong(CodigoSolSrv.Text)
    
    objOS.lCodPedSRV = StrParaLong(PedSRV.Text)
    objOS.iFilialPedSRV = Codigo_Extrai(FilialPedSRV.Text)
    
    objOS.sLote = Lote.Text
    objOS.iFilialOP = Codigo_Extrai(FilialOP.Text)
    
    'Critica o formato do Produto
    lErro = CF("Produto_Formata", ProdutoSolServ.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195512
    
    objOS.sProduto = sProdutoFormatado
    
    If Tipo.ListIndex <> -1 Then
        objOS.lTipo = Tipo.ItemData(Tipo.ListIndex)
    End If
    
    lErro = Move_Grid_Memoria(objOS)
    If lErro <> SUCESSO Then gError 195052

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 195052, 195512
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195053)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOS As New ClassOS

On Error GoTo Erro_Tela_Preenche

    objOS.sCodigo = colCampoValor.Item("Codigo").vValor
    objOS.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objOS.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor

    'Traz dados da Ordem de Produção para a Tela
    lErro = Traz_Tela_OS(objOS)
    If lErro <> SUCESSO And lErro <> 195054 Then gError 195055

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195056)

    End Select

    Exit Sub

End Sub

Function Traz_Tela_OS(objOS As ClassOS) As Long
'preenche a tela com os dados da OS

Dim lErro As Long
'Dim sProdutoMascarado As String
Dim iIndice As Integer

On Error GoTo Erro_Traz_Tela_OS

    lErro = Limpa_Tela_OS(NAO_FECHAR_SETAS)
    If lErro <> SUCESSO Then gError 195057

    CodigoOS.Text = objOS.sCodigo

    lErro = CF("OS_Le_ComItens", objOS)
    If lErro <> SUCESSO And lErro <> 195060 Then gError 195064

    If lErro <> SUCESSO Then gError 195054

    Call DateParaMasked(Data, objOS.dtDataEmissao)
    
    CodigoSolSrv.Text = objOS.lCodSolSRV
    
'    sProdutoMascarado = String(STRING_PRODUTO, 0)
    
    If objOS.lCodPedSRV <> 0 Then PedSRV.Text = objOS.lCodPedSRV
    
    If objOS.iFilialPedSRV <> 0 Then FilialPedSRV.Text = objOS.iFilialPedSRV
    Call FilialPedSRV_Validate(bSGECancelDummy)

'    'Mascara produto
'    lErro = Mascara_RetornaProdutoTela(objOS.sProduto, sProdutoMascarado)
'    If lErro <> SUCESSO Then gError 195028
'
'    ProdutoSolServ.PromptInclude = False
'    ProdutoSolServ.Text = sProdutoMascarado
'    ProdutoSolServ.PromptInclude = True
'    lErro = CF("Produto_Perde_Foco", ProdutoSolServ, DescProd)
'    If lErro <> SUCESSO And lErro <> 27095 Then gError 195064
    lErro = CF("Traz_Produto_MaskEd", objOS.sProduto, ProdutoSolServ, DescProd)
    If lErro <> SUCESSO Then gError 195064
   
    Lote.Text = objOS.sLote
    
    FilialOP.Text = objOS.iFilialOP
    
    'Se o Rastreamento possui FilialOP (Rastro Por Ordem de Produção)
    If objOS.iFilialOP <> 0 Then

        For iIndice = 0 To FilialOP.ListCount - 1
            If FilialOP.ItemData(iIndice) = objOS.iFilialOP Then
                FilialOP.ListIndex = iIndice
                Exit For
            End If
        Next

    End If
    
    
    If objOS.iNumItensBaixados = objOS.iNumItens Then
        StatusOS.Caption = STRING_STATUS_BAIXADO
    Else
        StatusOS.Caption = "ABERTO"
    End If
    
    iQtdeAlterada = 0
    
    'preenche o grid
    lErro = Preenche_GridMovimentos(objOS.colItens)
    If lErro <> SUCESSO Then gError 195081
    
    If objOS.lTipo <> 0 Then
        Call Combo_Seleciona_ItemData(Tipo, objOS.lTipo)
    Else
        Tipo.ListIndex = -1
    End If
    
    iAlterado = 0
    iCodigoAlterado = 0
    
    Traz_Tela_OS = SUCESSO

    Exit Function

Erro_Traz_Tela_OS:

    Traz_Tela_OS = gErr

    Select Case gErr

        Case 195054, 195057, 195064, 195081

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195082)

    End Select

    Exit Function

End Function

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col))) <> 0 And GridMovimentos.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 195084

        'Coloca o valor do Ccl na coluna correspondente
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 195084

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195085)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOS As ClassOS

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objOS = obj1

    'traz OP para a tela
    lErro = Traz_Tela_OS(objOS)
    If lErro <> SUCESSO And lErro <> 195054 Then gError 195086

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 195086

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195087)
            
    End Select

    Exit Sub

End Sub

Private Sub UpDownData_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 195088

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 195088

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195089)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 195090

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 195090

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195091)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_gcolItemOS(gcolItemOS As Collection)
'limpa gcolItemOP e a coluna de qtde das linhas que estavam na tela apos a troca do codigo da OP

Dim lErro As Long
Dim iCount As Integer
Dim iIndice As Integer

On Error GoTo Erro_Limpa_gcolItemOS

    iCount = gcolItemOS.Count
    Set gcolItemOS = New Collection

    For iIndice = 1 To iCount

        gcolItemOS.Add 0

    Next

    Exit Sub

Erro_Limpa_gcolItemOS:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195092)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto, objItemOS As ClassItemOS) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sAlmoxarifadoPadrao As String

On Error GoTo Erro_ProdutoLinha_Preenche

    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_StatusItem_Col) = StatusItem.List(0)

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
    
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        gcolItemOS.Add 0
        gobjOS.colItens.Add objItemOS
        
        gobjOS.colItens(GridMovimentos.Row).sSiglaUM = objProduto.sSiglaUMEstoque
        gobjOS.colItens(GridMovimentos.Row).iItem = GridMovimentos.Row
        gobjOS.colItens(GridMovimentos.Row).sServico = objProduto.sCodigo
    
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 193991)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Ordem de Serviço"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OrdemServico"
    
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

Private Sub UpDownItemOS_UpClick()

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_UpDownItemOS_UpClick

    ItemOS.SetFocus

    If Len(ItemOS.ClipText) > 0 Then

        iItem = StrParaInt(ItemOS.Text)
        
        iItem = iItem + 1
        
        If iItem > objGrid.iLinhasExistentes Then
        
            iItem = objGrid.iLinhasExistentes
            
        End If
        
    Else
    
        iItem = 1
        
    End If

    ItemOS.PromptInclude = False
    ItemOS.Text = CStr(iItem)
    ItemOS.PromptInclude = True
    
    Call ItemOS_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownItemOS_UpClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195093)

    End Select

    Exit Sub


End Sub

Private Sub UpDownItemOS_DownClick()

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_UpDownItemOS_DownClick

    ItemOS.SetFocus

    If Len(Trim(ItemOS.ClipText)) > 0 Then

        iItem = StrParaInt(ItemOS.Text)
        
        iItem = iItem - 1
        
        If iItem < 1 Then
        
            iItem = 1
        
        End If
    
    Else
               
        iItem = 1
        
    End If

    ItemOS.PromptInclude = False
    ItemOS.Text = CStr(iItem)
    ItemOS.PromptInclude = True
    
    Call ItemOS_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownItemOS_DownClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195094)

    End Select

    Exit Sub

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is CodigoOS Then
            Call CodigoOSLabel_Click
        ElseIf Me.ActiveControl Is Servico Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is CodigoCompetencia Then
            Call CompetenciaLabel_Click
        ElseIf Me.ActiveControl Is CodigoCTPadrao Then
            Call CTLabel_Click
        ElseIf Me.ActiveControl Is CodigoSolSrv Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is CodigoPecaPrev Then
            Call BotaoPecas_Click
        ElseIf Me.ActiveControl Is NomeRedMaquinaPrev Then
            Call BotaoMaquinas_Click
        ElseIf Me.ActiveControl Is CodigoMOPrev Then
            Call BotaoMO_Click
        End If

    End If

End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub CodigoOSLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoOSLabel, Source, X, Y)
End Sub

Private Sub CodigoOSLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoOSLabel, Button, Shift, X, Y)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'Incluido por Jorge Specian - 25/05/2005
        '---------------------------------------
        'Se Frame selecionado foi o de Operacoes
        If TabStrip1.SelectedItem.Index = TAB_Operacoes Then
            
            If IsNumeric(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) Then
                dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
            Else
                dQuantidade = 0
            End If
        
            lErro = Mostra_Arvore(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), dQuantidade)
            If lErro <> SUCESSO Then gError 195095
            
        End If
        '---------------------------------------
    
    End If
    
    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr
    
        Case 195095
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195096)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)
'Critica se a Data da OP está preenchida corretamente
Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 195097

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 195097

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195098)

    End Select

    Exit Sub

End Sub

Private Sub CodigoOS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOS As New ClassOS
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_CodigoOS_Validate

    'Se houve alteração nos dados da tela
    If (iCodigoAlterado = REGISTRO_ALTERADO) Then

        If Len(Trim(CodigoOS.Text)) > 0 Then

            'limpa a coleção global
            Call Limpa_gcolItemOS(gcolItemOS)

            objOS.sCodigo = CodigoOS.Text
            objOS.iFilialEmpresa = giFilialEmpresa

            'tenta ler a OS desejada
            lErro = CF("OS_Le", objOS)
            If lErro <> SUCESSO And lErro <> 195102 Then gError 195104

            'se existir
            If lErro = SUCESSO Then

                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")

                If vbMsg = vbNo Then gError 195105

                'traz a OP para a tela
                lErro = Traz_Tela_OS(objOS)
                If lErro <> SUCESSO And lErro <> 195054 Then gError 195106

                Call ComandoSeta_Fechar(Me.Name)

            End If

        End If

    End If

    Exit Sub

Erro_CodigoOS_Validate:

    Cancel = True

    Select Case gErr

        Case 195104, 195105, 195106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195107)

    End Select

    Exit Sub

End Sub

'###################################################
'Inserido por Wagner
Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigoOS As String

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("OS_Automatico", sCodigoOS, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 195108

    CodigoOS.Text = sCodigoOS

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 195108
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195109)
    
    End Select

    Exit Sub

End Sub

Function Trata_Arvore(ByVal sProdutoRaiz As String, ByVal sVersao As String, ByVal sUMedida As String, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim bTrazDefault As Boolean
Dim objItemOS As ClassItemOS, objItemOSAux As ClassItemOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Arvore

    If Len(Trim(sProdutoRaiz)) <> 0 Then
    
        lErro = CF("Produto_Formata", sProdutoRaiz, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137126
        
        bTrazDefault = True
        
        'verifica se o item da OP já está na coleção
        For Each objItemOS In gobjOS.colItens
        
            'se está ...
            If objItemOS.sServico = sProdutoFormatado And objItemOS.colItemOSOperacoes.Count > 0 Then
            
                'Se alterou a Quantidade
                If iQtdeAlterada = REGISTRO_ALTERADO Then
                    
                    'Muda a Quantidades da coleção de Operações do objItemOS
                    lErro = AlteraQuantidade(objItemOS, dQuantidade)
                    If lErro <> SUCESSO Then gError 137072
                    
                    iQtdeAlterada = 0
                    
                End If
                
                bTrazDefault = False
                Exit For
                
            End If
            
        Next
        
        'se não está -> traz a árvore default
        If bTrazDefault Then
        
            For Each objItemOS In gobjOS.colItens
                If objItemOS.sServico = sProdutoFormatado Then
                    Exit For
                End If
            Next
    
            objItemOS.sServico = sProdutoFormatado
            objItemOS.sVersao = sVersao
            objItemOS.sSiglaUM = sUMedida
            objItemOS.dQuantidade = dQuantidade
           
            lErro = CF("ItemOS_Gera_Operacoes", objItemOS)
            If lErro <> SUCESSO Then gError 137083
            
            iQtdeAlterada = 0
            
        End If
            
    End If
        
    Trata_Arvore = SUCESSO
    
    Exit Function
    
Erro_Trata_Arvore:

    Trata_Arvore = gErr
    
    Select Case gErr
        
        Case 137072, 137083, 137126, 139268
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163879)
            
    End Select
    
    Exit Function

End Function

Function Mostra_Arvore(ByVal sProdutoRaiz As String, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim objItemOS As New ClassItemOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Mostra_Arvore

    'limpa Operacoes
    Call Limpa_Operacoes

    LabelCodigoCompetencia.Caption = ""
    LabelCodigoCTPadrao.Caption = ""

    'limpa a Arvore
    Call Limpa_Arvore_Roteiro

    If GridMovimentos.Row = 0 Or GridMovimentos.Row > objGrid.iLinhasExistentes Then

        GridMovimentos.Row = 1
        sProdutoRaiz = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col)
        dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))

    End If

'    If gbOPValidaParaMRP = True And Len(Trim(sProdutoRaiz)) <> 0 Then
    If Len(Trim(sProdutoRaiz)) <> 0 Then

        If dQuantidade = 0 Then gError 195149

        If GridMovimentos.Row <> StrParaInt(ItemOS.Text) Then

            ItemOS.PromptInclude = False
            ItemOS.Text = GridMovimentos.Row
            ItemOS.PromptInclude = True

            giItemOS = GridMovimentos.Row

        End If

        lErro = CF("Produto_Formata", sProdutoRaiz, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195150

        'verifica se o item da OS está na coleção
        For Each objItemOS In gobjOS.colItens

            'se está ...
            If objItemOS.sServico = sProdutoFormatado Then

                'Monta a árvore de operações
                lErro = Carrega_Arvore(objItemOS)
                If lErro <> SUCESSO Then gError 195151

                Exit For

            End If

        Next

    End If

    Mostra_Arvore = SUCESSO

    Exit Function

Erro_Mostra_Arvore:

    Mostra_Arvore = gErr

    Select Case gErr

        Case 195149
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, GridMovimentos.Row)

        Case 195150, 195151

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195152)

    End Select

    Exit Function

End Function

Private Sub BotaoAlterar_Click()

Dim lErro As Long
Dim objNode As Node
Dim sChave As String
Dim objItemOSOperacoes As ClassItemOSOperacoes
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objProduto As ClassProduto
Dim objItemOS As New ClassItemOS
Dim sTexto As String
Dim objItemOSOperacoes1 As New ClassItemOSOperacoes
Dim sProdutoRaiz As String

On Error GoTo Erro_BotaoAlterar_Click

    Set objNode = Roteiro.SelectedItem

    If objNode Is Nothing Then gError 195122
    If objNode.Selected = False Then gError 195123
    
    If Len(Trim(CodigoCompetencia.ClipText)) = 0 Then gError 195124
        
    Set objCompetencias = New ClassCompetencias
    Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
    Set objItemOSOperacoes = colComponentes.Item(objNode.Tag)
    
    'preenche objItemOSOperacoes à partir dos dados da tela
    lErro = Move_Operacoes_Memoria(objItemOSOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 195125

    sChave = objNode.Tag
        
    'prepara texto que identificará a nova Operação que está sendo incluida
    sTexto = objCompetencias.sNomeReduzido
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objItemOSOperacoes.sServico
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 195126
    
    sTexto = sTexto & " (" & objProduto.sNomeReduzido

    If Len(Trim(CodigoCTPadrao.ClipText)) <> 0 Then
       sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
    End If
        
    sTexto = sTexto & ")"

    objNode.Text = sTexto

    colComponentes.Remove (sChave)
    colComponentes.Add objItemOSOperacoes, sChave
    
    For Each objItemOS In gobjOS.colItens
    
        If objItemOS.sServico = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col) Then
                    
            sProdutoRaiz = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col)
            
            For Each objItemOSOperacoes1 In objItemOS.colItemOSOperacoes
            
                If objItemOSOperacoes1.iSeqRoteiro = objItemOSOperacoes.iSeqRoteiro Then
                
                    objItemOSOperacoes1.lNumIntDocCompet = objItemOSOperacoes.lNumIntDocCompet
                    objItemOSOperacoes1.lNumIntDocCT = objItemOSOperacoes.lNumIntDocCT
                    objItemOSOperacoes1.sObservacao = objItemOSOperacoes.sObservacao
                    objItemOSOperacoes1.iNivel = objItemOSOperacoes.iNivel
                    objItemOSOperacoes1.iSeq = objItemOSOperacoes.iSeq
                    objItemOSOperacoes1.iSeqArvore = objItemOSOperacoes.iSeqArvore
                    objItemOSOperacoes1.iSeqPai = objItemOSOperacoes.iSeqPai
                    objItemOSOperacoes1.iNivelRoteiro = objItemOSOperacoes.iNivelRoteiro
                    objItemOSOperacoes1.iSeqRoteiro = objItemOSOperacoes.iSeqRoteiro
                    objItemOSOperacoes1.iSeqRoteiroPai = objItemOSOperacoes.iSeqRoteiroPai
                    objItemOSOperacoes1.sServico = objItemOSOperacoes.sServico
                    objItemOSOperacoes1.sVersao = objItemOSOperacoes.sVersao
                    
                    Set objItemOSOperacoes1.colPecas = objItemOSOperacoes.colPecas
                    Set objItemOSOperacoes1.colMaquinas = objItemOSOperacoes.colMaquinas
                    Set objItemOSOperacoes1.colMO = objItemOSOperacoes.colMO
                
                    Exit For
            
                End If
                
            Next
            
            Exit For
        
        End If
        
    Next

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoAlterar_Click:

    Select Case gErr

        Case 195122, 195123
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 195124
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)

        Case 195125, 195126
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195127)

    End Select

    Exit Sub

End Sub

Private Sub CodigoCompetencia_Change()

    iAlterado = REGISTRO_ALTERADO
    iCompetenciaAlterada = REGISTRO_ALTERADO

End Sub

Private Sub CodigoCompetencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoCompetencia, iAlterado)
    
End Sub

Private Sub CodigoCompetencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho

On Error GoTo Erro_CodigoCompetencia_Validate

    DescricaoCompetencia.Caption = ""
    
    If GridMovimentos.Row = 0 Then Exit Sub

    'Verifica se CodigoCompetencia não está preenchida
    If Len(Trim(CodigoCompetencia.Text)) = 0 Then

        Exit Sub
    
    End If

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195128

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195129

    Set objCompetencias = New ClassCompetencias
    
    'Verifica sua existencia
    lErro = CF("TP_Competencia_Le", CodigoCompetencia, objCompetencias)
    If lErro <> SUCESSO Then gError 195129
    
    DescricaoCompetencia.Caption = objCompetencias.sDescricao
    
    'Verifica se CodigoCompetencia foi alterado
    If iCompetenciaAlterada = REGISTRO_ALTERADO Then
            
        CodigoCTPadrao.Text = ""
        DescricaoCTPadrao.Caption = ""
        
        'Verifica se existe CTPadrao cadastrado na Competencia e traz seus dados
        lErro = CF("Competencias_Le_CTPadrao", objCompetencias, objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134909 Then gError 195130
        
        If lErro = SUCESSO Then
        
           CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
           
           Call CodigoCTPadrao_Validate(bSGECancelDummy)
        
        End If
        
        iCompetenciaAlterada = 0
    
    End If
       
    Exit Sub

Erro_CodigoCompetencia_Validate:

    Cancel = True

    Select Case gErr

        Case 195128, 195129, 195130

        Case 195129
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACOES_SERVICO_NAO_PREENCHIDO", gErr)
            CodigoCompetencia.Text = ""
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195131)

    End Select

    Exit Sub

End Sub

Private Sub CodigoCTPadrao_Change()

    iAlterado = REGISTRO_ALTERADO
    
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
    
    If GridMovimentos.Row = 0 Then Exit Sub

    'Verifica se CodigoCTPadrao não está preenchido
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then

        'Verifica se CodigoCompetencia está preenchido
        If Len(Trim(CodigoCompetencia.Text)) = 0 Then gError 195132
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CodigoCTPadrao, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 195133
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = CodigoCompetencia.Text
        
        'Lê a Competencia pelo NomeReduzido para verificar seu NumIntDoc
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 195134
    
        If lErro <> SUCESSO Then gError 195135
        
        lErro = CF("CentrodeTrabalho_Le_CTCompetencias", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134453 Then gError 195136
    
        bCompetenciaCadastrada = False
        
        For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
        
            If objCTCompetencias.lNumIntDocCompet = objCompetencias.lNumIntDoc Then
            
                bCompetenciaCadastrada = True
                Exit For
                
            End If
        
        Next
            
        If bCompetenciaCadastrada = False Then gError 195137
            
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
        
    End If
       
    Exit Sub

Erro_CodigoCTPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 195132
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_NAO_PREENCHIDA", gErr)
            CodigoCTPadrao.Text = ""
                        
        Case 195133, 195134, 195136
        
        Case 195135, 195137
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_NAO_CADASTRADA_CT", gErr, objCentrodeTrabalho.lCodigo)
            CodigoCTPadrao.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195138)

    End Select

    Exit Sub

End Sub

Private Sub CompetenciaLabel_Click()

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CompetenciaLabel_Click

    If GridMovimentos.Row = 0 Then Exit Sub
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195139

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        'Verifica se a Competencia foi preenchida
        If Len(Trim(CodigoCompetencia.Text)) <> 0 Then
                    
            objCompetencias.sNomeReduzido = CodigoCompetencia.Text
            
            'Verifica a Competencia no BD a partir do NomeReduzido
            lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
            If lErro <> SUCESSO And lErro <> 134937 Then gError 195140
    
        End If
    
        Call Chama_Tela("CompetenciasLista", colSelecao, objCompetencias, objEventoCompetencias)

    End If
    
    Exit Sub

Erro_CompetenciaLabel_Click:

    Select Case gErr
    
        Case 195139, 195140

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195141)

    End Select

    Exit Sub

End Sub

Private Sub CTLabel_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CTLabel

    If GridMovimentos.Row = 0 Then Exit Sub
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195142

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        'Verifica se o CodigoCTPadrao foi preenchido
        If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
            
            objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
            
            'Verifica o CodigoCTPadrao, lendo no BD a partir do NomeReduzido
            lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134941 Then gError 195143
            
        End If
    
        Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCentroDeTrabalho)

    End If

    Exit Sub

Erro_CTLabel:

    Select Case gErr
    
        Case 195142, 195143

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195144)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCentroDeTrabalho_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCentroDeTrabalho_evSelecao

    Set objCentrodeTrabalho = obj1

    CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CodigoCTPadrao_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCentroDeTrabalho_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195145)

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

    Me.Show

    Exit Sub

Erro_objEventoCompetencia_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195146)

    End Select

    Exit Sub

End Sub

Private Sub Roteiro_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objItemOSOperacoes As New ClassItemOSOperacoes
Dim objNode As Node
Dim sVersao As String

On Error GoTo Erro_Roteiro_NodeClick

    Set objNode = Roteiro.SelectedItem

    Set objItemOSOperacoes = colComponentes.Item(objNode.Tag)
    
    'Preenche as Operações
    lErro = Preenche_Operacoes(objItemOSOperacoes)
    If lErro <> SUCESSO Then gError 195147

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Roteiro_NodeClick:

    Select Case gErr

        Case 195147
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195148)

    End Select

    Exit Sub

End Sub

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

Private Function Inicializa_GridPecasPrev(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Peça")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("C/P")
    objGrid.colColuna.Add ("Versão")

    'Controles que participam do Grid
    objGrid.colCampo.Add (CodigoPecaPrev.Name)
    objGrid.colCampo.Add (DescricaoPecaPrev.Name)
    objGrid.colCampo.Add (UMPecaPrev.Name)
    objGrid.colCampo.Add (QuantidadePecaPrev.Name)
    objGrid.colCampo.Add (OrigemPecaPrev.Name)
    objGrid.colCampo.Add (VersaoPecaPrev.Name)

    'Colunas do Grid
    iGrid_CodigoPecaPrev_Col = 1
    iGrid_DescricaoPecaPrev_Col = 2
    iGrid_UMPecaPrev_Col = 3
    iGrid_QuantidadePecaPrev_Col = 4
    iGrid_OrigemPecaPrev_Col = 5
    iGrid_VersaoPecaPrev_Col = 6

    objGrid.objGrid = GridPecasPrev

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridPecasPrev.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridPecasPrev = SUCESSO

End Function


Private Sub GridPecasPrev_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPecasPrev, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPecasPrev, iAlterado)
    End If

End Sub

Private Sub GridPecasPrev_GotFocus()
    
    Call Grid_Recebe_Foco(objGridPecasPrev)

End Sub

Private Sub GridPecasPrev_EnterCell()

    Call Grid_Entrada_Celula(objGridPecasPrev, iAlterado)

End Sub

Private Sub GridPecasPrev_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridPecasPrev)
End Sub

Private Sub GridPecasPrev_LeaveCell()
    
    Call Saida_Celula(objGridPecasPrev)

End Sub

Private Sub GridPecasPrev_RowColChange()
    Call Grid_RowColChange(objGridPecasPrev)
End Sub


Private Sub GridPecasPrev_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPecasPrev, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPecasPrev, iAlterado)
    End If

End Sub

Private Sub GridPecasPrev_Scroll()

    Call Grid_Scroll(objGridPecasPrev)

End Sub


Private Sub DescricaoPecaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoPecaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPecasPrev)

End Sub

Private Sub DescricaoPecaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPecasPrev)

End Sub

Private Sub DescricaoPecaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPecasPrev.objControle = DescricaoPecaPrev
    lErro = Grid_Campo_Libera_Foco(objGridPecasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadePecaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadePecaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPecasPrev)

End Sub

Private Sub QuantidadePecaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPecasPrev)

End Sub

Private Sub QuantidadePecaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPecasPrev.objControle = QuantidadePecaPrev
    lErro = Grid_Campo_Libera_Foco(objGridPecasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMPecaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMPecaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPecasPrev)

End Sub

Private Sub UMPecaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPecasPrev)

End Sub

Private Sub UMPecaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPecasPrev.objControle = UMPecaPrev
    lErro = Grid_Campo_Libera_Foco(objGridPecasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CodigoPecaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoPecaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPecasPrev)

End Sub

Private Sub CodigoPecaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPecasPrev)

End Sub

Private Sub CodigoPecaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPecasPrev.objControle = CodigoPecaPrev
    lErro = Grid_Campo_Libera_Foco(objGridPecasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VersaoPecaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VersaoPecaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPecasPrev)

End Sub

Private Sub VersaoPecaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPecasPrev)

End Sub

Private Sub VersaoPecaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPecasPrev.objControle = VersaoPecaPrev
    lErro = Grid_Campo_Libera_Foco(objGridPecasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Carrega_Arvore(ByVal objItemOS As ClassItemOS) As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim objNode As Node
Dim lErro As Long, sChave As String, sChaveTvw As String
Dim iIndice As Integer
Dim sTexto As String
Dim objItemOSOperacoes As New ClassItemOSOperacoes
Dim objCompetencias As ClassCompetencias
Dim objProduto As ClassProduto
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Carrega_Arvore

    Set gobjItemOS = objItemOS

    For Each objItemOSOperacoes In objItemOS.colItemOSOperacoes

        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objItemOSOperacoes.lNumIntDocCompet
        
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 195150

        'prepara texto que identificará a nova Operação que está sendo incluida
        sTexto = objCompetencias.sNomeReduzido
        
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objItemOSOperacoes.sServico
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195151
        
        sTexto = sTexto & " (" & objProduto.sNomeReduzido

        If objItemOSOperacoes.lNumIntDocCT > 0 Then
        
            Set objCentrodeTrabalho = New ClassCentrodeTrabalho
            
            objCentrodeTrabalho.lNumIntDoc = objItemOSOperacoes.lNumIntDocCT
            
            lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 195152
            
            If lErro = SUCESSO Then
        
                sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
                
            End If
           
        End If
        
        sTexto = sTexto & ")"
        
        'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
        Call Calcula_Proxima_Chave(sChaveTvw)
        
        sChave = sChaveTvw
        sChaveTvw = sChaveTvw & objCompetencias.lCodigo

        If objItemOSOperacoes.iNivel = 0 Then

            Set objNode = Roteiro.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)

        Else

            Set objNode = Roteiro.Nodes.Add(objItemOSOperacoes.iSeqPai, tvwChild, sChaveTvw, sTexto)

        End If
                
        Roteiro.Nodes.Item(objNode.Index).Expanded = True
        
        colComponentes.Add objItemOSOperacoes, sChave
        
        objNode.Tag = sChave
        
    Next
    
    Call Habilita_Operacoes(Roteiro.Nodes.Count)
    
    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr

        Case 195150 To 195152
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195153)

    End Select

    Exit Function

End Function

Private Sub Carrega_ComboVersoes(ByVal sProdutoRaiz As String)
    
Dim lErro As Long
Dim objRot As New ClassRoteiroSRV
Dim colRot As New Collection
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    Versao.Enabled = True
    
    'Limpa a Combo
    Versao.Clear

    'Le as Versoes Ativas e a Padrao
    lErro = CF("RoteiroSRV_Le_Versoes", sProdutoRaiz, colRot)
    If lErro <> SUCESSO Then gError 106321

    'Carrega a Combo com os Dados da Colecao
    For Each objRot In colRot

        Versao.AddItem (objRot.sVersao)

    Next

    If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)) > 0 Then

        Versao.Text = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)

    ElseIf colRot.Count = 1 Then

        'Seleciona a Padrao na Combo
        Versao.ListIndex = 0

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col) = Versao.Text

    End If

    Exit Sub
    
Erro_Carrega_ComboVersoes:

    Select Case gErr
    
        Case 106321
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163872)
    
    End Select
    
End Sub

Function Preenche_Operacoes(objItemOSOperacoes As ClassItemOSOperacoes) As Long
'preenche as tabs de Pecas, Maquinas e MO à partir dos dados de objItemOSOperacoes

Dim lErro As Long
Dim iAlteradoAnterior As Integer
Dim objProduto As ClassProduto
Dim sProdutoMascarado As String
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Preenche_Operacoes
    
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 195161
    
    iAlteradoAnterior = iAlterado

    Nivel.Caption = objItemOSOperacoes.iNivel
    Sequencial.Caption = objItemOSOperacoes.iSeqArvore
    
    Set objProduto = New ClassProduto
    
    'objItemOSOperacoes.sServico = gobjOS.colItens.Item(1).sServico
    objProduto.sCodigo = objItemOSOperacoes.sServico
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 195154
    
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 195155
    
    ProdutoLabel.Caption = sProdutoMascarado & SEPARADOR & objProduto.sDescricao
    VersaoLabel.Caption = objItemOSOperacoes.sVersao
    
    QtdeLabel.Caption = Formata_Estoque(objItemOSOperacoes.dQuantidade)
    
    UMLabel.Caption = objItemOSOperacoes.sUMedida
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.lNumIntDoc = objItemOSOperacoes.lNumIntDocCompet
    
    lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134336 Then gError 195156
    
    CodigoCompetencia.PromptInclude = False
    CodigoCompetencia.Text = objCompetencias.sNomeReduzido
    CodigoCompetencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = objCompetencias.sDescricao
    
    If objItemOSOperacoes.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = objItemOSOperacoes.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 195157
        
        CodigoCTPadrao.PromptInclude = False
        CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        CodigoCTPadrao.PromptInclude = True
        
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If
    
    Observacao.Text = objItemOSOperacoes.sObservacao
    
    lErro = Preenche_GridPecasPrev(objItemOSOperacoes)
    If lErro <> SUCESSO Then gError 195158
    
    lErro = Preenche_GridMaquinasPrev(objItemOSOperacoes)
    If lErro <> SUCESSO Then gError 195159
    
    lErro = Preenche_GridMOPrev(objItemOSOperacoes)
    If lErro <> SUCESSO Then gError 195160
    
    Call AlteraHabilitacao_Operacoes
    
    iCompetenciaAlterada = 0
    iAlterado = iAlteradoAnterior

    Preenche_Operacoes = SUCESSO
    
    Exit Function

Erro_Preenche_Operacoes:

    Preenche_Operacoes = gErr

    Select Case gErr
    
        Case 195154 To 195161
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195162)

    End Select

    Exit Function

End Function

Function Limpa_Operacoes() As Long

Dim lErro As Long
Dim iAlteradoAnterior As Integer

On Error GoTo Erro_Limpa_Operacoes
        
    iAlteradoAnterior = iAlterado
    
    Nivel.Caption = ""
    Sequencial.Caption = ""
    
    ProdutoLabel.Caption = ""
    VersaoLabel.Caption = ""
    QtdeLabel.Caption = ""
    UMLabel.Caption = ""
    
    CodigoCompetencia.PromptInclude = False
    CodigoCompetencia.Text = ""
    CodigoCompetencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = ""
    
    CodigoCTPadrao.PromptInclude = False
    CodigoCTPadrao.Text = ""
    CodigoCTPadrao.PromptInclude = True
    
    DescricaoCTPadrao.Caption = ""
    
    Observacao.Text = ""
    
    Call Grid_Limpa(objGridPecasPrev)
    Call Grid_Limpa(objGridMaquinasPrev)
    Call Grid_Limpa(objGridMOPrev)
        
    iCompetenciaAlterada = 0
    iAlterado = iAlteradoAnterior

    Limpa_Operacoes = SUCESSO

    Exit Function

Erro_Limpa_Operacoes:

    Limpa_Operacoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195163)

    End Select

    Exit Function

End Function

Private Function Preenche_GridPecasPrev(objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_Preenche_GridPecasPrev
    
    Call Grid_Limpa(objGridPecasPrev)
    
    'Exibe os dados da coleção na tela
    For iIndice = 1 To objItemOSOperacoes.colPecas.Count
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objItemOSOperacoes.colPecas.Item(iIndice).sProduto
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195164
        
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 195165
                
        'Insere no GridOrdemProducaoInsumos
        GridPecasPrev.TextMatrix(iIndice, iGrid_CodigoPecaPrev_Col) = sProdutoMascarado
        GridPecasPrev.TextMatrix(iIndice, iGrid_DescricaoPecaPrev_Col) = objProdutos.sDescricao
        
        If objProdutos.iCompras = PRODUTO_COMPRAVEL Then
            GridPecasPrev.TextMatrix(iIndice, iGrid_OrigemPecaPrev_Col) = INSUMO_COMPRADO
        Else
            GridPecasPrev.TextMatrix(iIndice, iGrid_OrigemPecaPrev_Col) = INSUMO_PRODUZIDO
        End If
        
        If objItemOSOperacoes.colPecas.Item(iIndice).dQuantidade > 0 Then
            GridPecasPrev.TextMatrix(iIndice, iGrid_QuantidadePecaPrev_Col) = Formata_Estoque(objItemOSOperacoes.colPecas.Item(iIndice).dQuantidade)
        End If
        
        GridPecasPrev.TextMatrix(iIndice, iGrid_UMPecaPrev_Col) = objItemOSOperacoes.colPecas.Item(iIndice).sUMProduto

    Next

    objGridPecasPrev.iLinhasExistentes = objItemOSOperacoes.colPecas.Count
    
    Preenche_GridPecasPrev = SUCESSO
    
    Exit Function

Erro_Preenche_GridPecasPrev:

    Preenche_GridPecasPrev = gErr

    Select Case gErr

        Case 195164, 195165

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195166)

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
    If lErro <> SUCESSO Then gError 195167

    Roteiro.Nodes.Clear
    Set colComponentes = New Collection
    
    iProxChave = 1

    Limpa_Arvore_Roteiro = SUCESSO

    Exit Function

Erro_Limpa_Arvore_Roteiro:

    Limpa_Arvore_Roteiro = gErr
    
    Select Case gErr

        Case 195167

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195168)

    End Select

    Exit Function

End Function

Sub Calcula_Proxima_Chave(sChave As String)

Dim iNumero As Integer

    iNumero = iProxChave
    iProxChave = iProxChave + 1
    sChave = "X" & right$(CStr(100000 + iNumero), 5)

End Sub

Private Function Move_Operacoes_Memoria(ByVal objItemOSOperacoes As ClassItemOSOperacoes, ByVal objCompetencias As ClassCompetencias, ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Operacoes_Memoria
        
    objCompetencias.sNomeReduzido = CodigoCompetencia.Text
    
    'Verifica a Competencia no BD a partir do Código
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 195121

    objItemOSOperacoes.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 195122
        
        objItemOSOperacoes.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Trim(Observacao.Text)) <> 0 Then objItemOSOperacoes.sObservacao = Observacao.Text
    
    lErro = Move_OperacaoPecas_Memoria(objItemOSOperacoes)
    If lErro <> SUCESSO Then gError 195123
    
    lErro = Move_OperacaoMaquinas_Memoria(objItemOSOperacoes)
    If lErro <> SUCESSO Then gError 195124
    
    lErro = Move_OperacaoMO_Memoria(objItemOSOperacoes)
    If lErro <> SUCESSO Then gError 195125
    
    Move_Operacoes_Memoria = SUCESSO

    Exit Function

Erro_Move_Operacoes_Memoria:

    Move_Operacoes_Memoria = gErr

    Select Case gErr

        Case 195121 To 195125
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195126)

    End Select

    Exit Function

End Function

Private Function Move_OperacaoPecas_Memoria(objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim objItemOSOperacoesPecas As ClassItemOSOperacoesPecas
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer

On Error GoTo Erro_Move_OperacaoPecas_Memoria

    'Limpa coleção para reiniciar
    Set objItemOSOperacoes.colPecas = New Collection

    'Ir preenchendo a colecao no objItemOSOperacoes com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridPecasPrev.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridPecasPrev.TextMatrix(iIndice, iGrid_CodigoPecaPrev_Col))) = 0 Then Exit For
                
        Set objProdutos = New ClassProduto
        
        lErro = CF("Produto_Formata", GridPecasPrev.TextMatrix(iIndice, iGrid_CodigoPecaPrev_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195110
        
        objProdutos.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195111
        
        Set objItemOSOperacoesPecas = New ClassItemOSOperacoesPecas
        
        objItemOSOperacoesPecas.lNumIntDocOper = objItemOSOperacoes.lNumIntDoc
        objItemOSOperacoesPecas.sProduto = objProdutos.sCodigo
        objItemOSOperacoesPecas.dQuantidade = StrParaDbl(GridPecasPrev.TextMatrix(iIndice, iGrid_QuantidadePecaPrev_Col))
        objItemOSOperacoesPecas.sUMProduto = GridPecasPrev.TextMatrix(iIndice, iGrid_UMPecaPrev_Col)
        objItemOSOperacoesPecas.sVersaoKitComp = GridPecasPrev.TextMatrix(iIndice, iGrid_VersaoPecaPrev_Col)
    
        objItemOSOperacoes.colPecas.Add objItemOSOperacoesPecas
    
    Next

    Move_OperacaoPecas_Memoria = SUCESSO

    Exit Function

Erro_Move_OperacaoPecas_Memoria:

    Move_OperacaoPecas_Memoria = gErr

    Select Case gErr
            
        Case 195110, 195111

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195112)

    End Select

    Exit Function

End Function

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

Dim objItemOSOperacoes As New ClassItemOSOperacoes
Dim sChave1 As String

    sChave1 = objNode.Tag

    Set objItemOSOperacoes = colComponentes.Item(sChave1)

    aNivelSequencial(iNivel) = aNivelSequencial(iNivel) + 1

    iPosicaoAtual = iPosicaoAtual + 1
    aSeqPai(iNivel) = iPosicaoAtual

    objItemOSOperacoes.iSeqArvore = aNivelSequencial(iNivel)

    If iNivel > 0 Then
        'objItemOSOperacoes.iSeqRoteiroPai = aSeqPai(iNivel - 1)
        objItemOSOperacoes.iSeqPai = aSeqPai(iNivel - 1)
    Else
        'objItemOSOperacoes.iSeqRoteiroPai = 0
        objItemOSOperacoes.iSeqPai = 0
    End If
    
    'objItemOSOperacoes.iSeqRoteiro = iPosicaoAtual
    objItemOSOperacoes.iSeq = iPosicaoAtual

    'objItemOSOperacoes.iNivelRoteiro = iNivel
    objItemOSOperacoes.iNivel = iNivel
    
    colComponentes.Remove sChave1
    colComponentes.Add objItemOSOperacoes, sChave1

    If objNode.Children > 0 Then
        Call Calcula_Nivel_Sequencial(objNode.Child, iNivel + 1, iPosicaoAtual)
    End If

    If objNode.Index <> objNode.LastSibling.Index Then Call Calcula_Nivel_Sequencial(objNode.Next, iNivel, iPosicaoAtual)

    If iNivel > iUltimoNivel Then iUltimoNivel = iNivel
   
End Sub

Private Sub GridMaquinasPrev_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaquinasPrev)

End Sub

Private Sub GridMaquinasPrev_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaquinasPrev, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinasPrev, iAlterado)
    End If

End Sub

Private Sub GridMaquinasPrev_Scroll()

    Call Grid_Scroll(objGridMaquinasPrev)

End Sub

Private Sub GridMaquinasPrev_EnterCell()

    Call Grid_Entrada_Celula(objGridMaquinasPrev, iAlterado)

End Sub

Private Sub GridMaquinasPrev_LeaveCell()
    
    Call Saida_Celula(objGridMaquinasPrev)

End Sub

Private Sub GridMaquinasPrev_RowColChange()

    Call Grid_RowColChange(objGridMaquinasPrev)

End Sub

Private Sub GridMaquinasPrev_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridMaquinasPrev)
    
End Sub


'Function AlteraQuantidade(ByVal objItemOP As ClassItemOP, ByVal dQuantidade As Double) As Long
'
'Dim lErro As Long
'Dim dFatorQtdeOP As Double
'Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
'Dim objOrdemProducaoInsumos As New ClassOrdemProducaoInsumos
'Dim objOperacoesTempo As ClassOperacoesTempo
'Dim dTempoCorrigido As Double
'Dim dQuantidadeCorrigida As Double
'
'On Error GoTo Erro_AlteraQuantidade
'
'    dFatorQtdeOP = dQuantidade / objItemOP.dQuantidade
'
'    For Each objOrdemProducaoOperacoes In objItemOP.colOrdemProducaoOperacoes
'
'        If objOrdemProducaoOperacoes.iIgnoraTaxaProducao = MARCADO Then
'
'            Set objOperacoesTempo = objOrdemProducaoOperacoes.objOperacoesTempo
'
'            If objOperacoesTempo.iTipo <> ITEM_TIPO_TAXAPRODUCAO_FIXO Then
'
'                dTempoCorrigido = objOperacoesTempo.dTempoOperacao * dFatorQtdeOP
'                objOperacoesTempo.dTempoOperacao = dTempoCorrigido
'
'            End If
'
'            Set objOrdemProducaoOperacoes.objOperacoesTempo = objOperacoesTempo
'
'        End If
'
'        For Each objOrdemProducaoInsumos In objOrdemProducaoOperacoes.colOPInsumos
'
'            If objOrdemProducaoInsumos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
'
'                dQuantidadeCorrigida = objOrdemProducaoInsumos.dQuantidade * dFatorQtdeOP
'                objOrdemProducaoInsumos.dQuantidade = dQuantidadeCorrigida
'
'            End If
'
'        Next
'
'    Next
'
'    objItemOP.dQuantidade = dQuantidade
'
'    AlteraQuantidade = SUCESSO
'
'    Exit Function
'
'Erro_AlteraQuantidade:
'
'    AlteraQuantidade = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163897)
'
'    End Select
'
'    Exit Function
'
'End Function



Function Habilita_Operacoes(ByVal lArvore As Long) As Long

Dim lErro As Long

On Error GoTo Erro_Habilita_Operacoes

    'se houver árvore ...
    If lArvore > 0 Then
        
        'selecionar a raiz
        Set Roteiro.SelectedItem = Roteiro.Nodes.Item(1)
        Roteiro.SelectedItem.Selected = True
        
        'e carregar as operações pertinentes
        Call Roteiro_NodeClick(Roteiro.Nodes.Item(1))
        
        LabelDetProduto.Enabled = True
        LabelDetVersao.Enabled = True
        LabelDetQtde.Enabled = True
        LabelDetUM.Enabled = True
        
        BotaoAlterar.Enabled = True
        BotaoImprimirOper.Enabled = True
        'só habilita se tiver mais de um item na OP
        If objGrid.iLinhasExistentes > 1 Then
            FrameItemOS.Enabled = True
            ItemOS.Enabled = True
            UpDownItemOS.Enabled = True
        Else
            FrameItemOS.Enabled = False
            ItemOS.Enabled = False
            UpDownItemOS.Enabled = False
        End If
        
        CompetenciaLabel.Enabled = True
        CodigoCompetencia.Enabled = True
        
        CTLabel.Enabled = True
        CodigoCTPadrao.Enabled = True
        
        LabelObservacao.Enabled = True
        Observacao.Enabled = True
        
        LabelNivel.Enabled = True
        
        LabelSeq.Enabled = True
        BotaoPecas.Enabled = True
        BotaoMO.Enabled = True
        BotaoMaquinas.Enabled = True
        GridPecasPrev.Enabled = True
        GridMOPrev.Enabled = True
        GridMaquinasPrev.Enabled = True
        
        
    Else
    
        LabelDetProduto.Enabled = False
        LabelDetVersao.Enabled = False
        LabelDetQtde.Enabled = False
        LabelDetUM.Enabled = False
    
        BotaoAlterar.Enabled = False
        BotaoImprimirOper.Enabled = False
        FrameItemOS.Enabled = False
        ItemOS.Enabled = False
        UpDownItemOS.Enabled = False
        
        CompetenciaLabel.Enabled = False
        CodigoCompetencia.Text = ""
        CodigoCompetencia.Enabled = False
        
        DescricaoCompetencia.Caption = ""
        
        CTLabel.Enabled = False
        CodigoCTPadrao.Text = ""
        CodigoCTPadrao.Enabled = False
        
        DescricaoCTPadrao.Caption = ""
        
        LabelObservacao.Enabled = False
        Observacao.Text = ""
        Observacao.Enabled = False
        
        LabelNivel.Enabled = False
        Nivel.Caption = ""
        
        LabelSeq.Enabled = False
        Sequencial.Caption = ""
        
        BotaoPecas.Enabled = False
        BotaoMO.Enabled = False
        BotaoMaquinas.Enabled = False
        
        GridPecasPrev.Enabled = False
        GridMOPrev.Enabled = False
        GridMaquinasPrev.Enabled = False
        
    End If
    
    Call AlteraHabilitacao_Operacoes
    
    Habilita_Operacoes = SUCESSO
    
    Exit Function
    
Erro_Habilita_Operacoes:
    
    Habilita_Operacoes = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195169)

    End Select
    
    Exit Function

End Function

Function AlteraHabilitacao_Operacoes() As Long

Dim lErro As Long

On Error GoTo Erro_AlteraHabilitacao_Operacoes

    CompetenciaLabel.ForeColor = &H80&
    CompetenciaLabel.MousePointer = 14       'Arrow and Question
    LabelCodigoCompetencia.Caption = ""
    LabelCodigoCompetencia.Visible = False
    
    CTLabel.ForeColor = &H80&
    CTLabel.MousePointer = 14                'Arrow and Question
    LabelCodigoCTPadrao.Caption = ""
    LabelCodigoCTPadrao.Visible = False

    If Roteiro.Nodes.Count > 0 Then
        
        CodigoCompetencia.Enabled = True
        CodigoCompetencia.Visible = True
    
        CodigoCTPadrao.Enabled = True
        CodigoCTPadrao.Visible = True

    End If
    
    
    AlteraHabilitacao_Operacoes = SUCESSO
    
    Exit Function
    
Erro_AlteraHabilitacao_Operacoes:

    AlteraHabilitacao_Operacoes = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195170)
    
    End Select
    
    Exit Function

End Function

Private Sub Roteiro_Collapse(ByVal Node As MSComctlLib.Node)
    Roteiro_NodeClick Node
End Sub

Private Function Inicializa_GridMaquinasPrev(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Nome")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Horas")

    'Controles que participam do Grid
    objGrid.colCampo.Add (NomeRedMaquinaPrev.Name)
    objGrid.colCampo.Add (QuantMaquinaPrev.Name)
    objGrid.colCampo.Add (HorasMaquinaPrev.Name)

    iGrid_NomeRedMaquinaPrev_Col = 1
    iGrid_QuantMaquinaPrev_Col = 2
    iGrid_HorasMaquinaPrev_Col = 3

    objGrid.objGrid = GridMaquinasPrev

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridMaquinasPrev.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaquinasPrev = SUCESSO

End Function

Private Sub NomeRedMaquinaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeRedMaquinaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinasPrev)

End Sub

Private Sub NomeRedMaquinaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinasPrev)

End Sub

Private Sub NomeRedMaquinaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinasPrev.objControle = NomeRedMaquinaPrev
    lErro = Grid_Campo_Libera_Foco(objGridMaquinasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantMaquinaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub QuantMaquinaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinasPrev)

End Sub

Private Sub QuantMaquinaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinasPrev)

End Sub

Private Sub QuantMaquinaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinasPrev.objControle = QuantMaquinaPrev
    lErro = Grid_Campo_Libera_Foco(objGridMaquinasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub HorasMaquinaPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub HorasMaquinaPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinasPrev)

End Sub

Private Sub HorasMaquinaPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinasPrev)

End Sub

Private Sub HorasMaquinaPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinasPrev.objControle = HorasMaquinaPrev
    lErro = Grid_Campo_Libera_Foco(objGridMaquinasPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Inicializa_GridMOPrev(objGrid As AdmGrid) As Long
'Inserido por Jorge Specian - 10/05/2005

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Código")
    objGrid.colColuna.Add ("Nome")
    objGrid.colColuna.Add ("Horas")

    'Controles que participam do Grid
    objGrid.colCampo.Add (CodigoMOPrev.Name)
    objGrid.colCampo.Add (NomeRedMOPrev.Name)
    objGrid.colCampo.Add (HorasMOPrev.Name)

    'Colunas do Grid
    iGrid_CodigoMOPrev_Col = 1
    iGrid_NomeRedMOPrev_Col = 2
    iGrid_HorasMOPrev_Col = 3

    objGrid.objGrid = GridMOPrev

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridMOPrev.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridMOPrev = SUCESSO

End Function

Private Sub CodigoMOPrev_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoMOPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOPrev)

End Sub

Private Sub CodigoMOPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOPrev)

End Sub

Private Sub CodigoMOPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOPrev.objControle = CodigoMOPrev
    lErro = Grid_Campo_Libera_Foco(objGridMOPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NomeRedMOPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeRedMOPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOPrev)

End Sub

Private Sub NomeRedMOPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOPrev)

End Sub

Private Sub NomeRedMOPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOPrev.objControle = NomeRedMOPrev
    lErro = Grid_Campo_Libera_Foco(objGridMOPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub HorasMOPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub HorasMOPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOPrev)

End Sub

Private Sub HorasMOPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOPrev)

End Sub

Private Sub HorasMOPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOPrev.objControle = HorasMOPrev
    lErro = Grid_Campo_Libera_Foco(objGridMOPrev)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoMO_Click()

Dim lErro As Long
Dim objTiposDeMaodeObras As New ClassTiposDeMaodeObra
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMO_Click

    If Me.ActiveControl Is CodigoMOPrev Then
    
        objTiposDeMaodeObras.iCodigo = StrParaInt(CodigoMOPrev.Text)
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridMOPrev.Row = 0 Then gError 134380

        objTiposDeMaodeObras.iCodigo = StrParaInt(GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_CodigoMOPrev_Col))
        
    End If

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTiposDeMaodeObras, objEventoMO)

    Exit Sub

Erro_BotaoMO_Click:

    Select Case gErr
        
        Case 134380
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162584)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMO_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoMO_evSelecao

    Set objTiposDeMaodeObra = obj1

    'Verifica se há algum produto repetido no grid
    For iLinha = 1 To objGridMOPrev.iLinhasExistentes
        
        If iLinha < GridMOPrev.Row Then
                                                
            If GridMOPrev.TextMatrix(iLinha, iGrid_CodigoMOPrev_Col) = objTiposDeMaodeObra.iCodigo Then
                CodigoMOPrev.Text = ""
                gError 134121
                
            End If
                
        End If
                       
    Next
       
    CodigoMOPrev.Text = CStr(objTiposDeMaodeObra.iCodigo)
    
    If Not (Me.ActiveControl Is CodigoMOPrev) Then
    
        GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_CodigoMOPrev_Col) = CStr(objTiposDeMaodeObra.iCodigo)
        GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_NomeRedMOPrev_Col) = objTiposDeMaodeObra.sDescricao
    
    End If

    'verifica se precisa preencher o grid com uma nova linha
    If GridMOPrev.Row - GridMOPrev.FixedRows = objGridMOPrev.iLinhasExistentes Then
        objGridMOPrev.iLinhasExistentes = objGridMOPrev.iLinhasExistentes + 1
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMO_evSelecao:

    Select Case gErr
    
        Case 134121
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, objTiposDeMaodeObra.iCodigo, iLinha)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195172)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMaquinas_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMaquinas_Click

    If GridMaquinasPrev.Row <= objGridMaquinasPrev.iLinhasExistentes Then

        If Me.ActiveControl Is NomeRedMaquinaPrev Then
                
            objMaquinas.sNomeReduzido = NomeRedMaquinaPrev.Text
            
        Else
        
            'Verifica se tem alguma linha selecionada no Grid
            If GridMaquinasPrev.Row = 0 Then gError 195173
    
            objMaquinas.sNomeReduzido = GridMaquinasPrev.TextMatrix(GridMaquinasPrev.Row, iGrid_NomeRedMaquinaPrev_Col)
            
        End If
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 195174
        
    End If
        
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaquinas)

    Exit Sub

Erro_BotaoMaquinas_Click:

    Select Case gErr

        Case 195173
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 195174

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195175)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMaquinas_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim iLinha As Integer
Dim iIndice As Integer

On Error GoTo Erro_objEventoMaquinas_evSelecao

    Set objMaquinas = obj1
        
    For iIndice = 1 To objGridMaquinasPrev.iLinhasExistentes
        If iIndice <> GridMaquinasPrev.Row Then
            'Tem a mesma máquina, erro
            If objMaquinas.sNomeReduzido = GridMaquinasPrev.TextMatrix(iIndice, iGrid_NomeRedMaquinaPrev_Col) Then
                gError 195180
            End If
        End If
    Next
        
    'Mostra os dados da Maquina na tela
    GridMaquinasPrev.TextMatrix(GridMaquinasPrev.Row, iGrid_NomeRedMaquinaPrev_Col) = objMaquinas.sNomeReduzido
    
    NomeRedMaquinaPrev.Text = objMaquinas.sNomeReduzido
       
    If GridMaquinasPrev.Row - GridMaquinasPrev.FixedRows = objGridMaquinasPrev.iLinhasExistentes Then
        objGridMaquinasPrev.iLinhasExistentes = objGridMaquinasPrev.iLinhasExistentes + 1
    End If
       
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoMaquinas_evSelecao:

    Select Case gErr

        Case 195180
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA", gErr, objMaquinas.sNomeReduzido, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195176)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_NomeRedMaquinaPrev(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CodigoItem do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim objMaquinas As New ClassMaquinas
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_NomeRedMaquinaPrev

    Set objGridInt.objControle = NomeRedMaquinaPrev

    'Se o campo foi preenchido
    If Len(Trim(NomeRedMaquinaPrev.Text)) > 0 Then
    
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", NomeRedMaquinaPrev, objMaquinas)
        If lErro <> SUCESSO Then gError 195181
            
        For iIndice = 1 To objGridMaquinasPrev.iLinhasExistentes
            If iIndice <> GridMaquinasPrev.Row Then
                'Tem a mesma máquina, erro
                If objMaquinas.sNomeReduzido = GridMaquinasPrev.TextMatrix(iIndice, iGrid_NomeRedMaquinaPrev_Col) Then
                    gError 195182
                End If
            End If
        Next
    
    
        GridMaquinasPrev.TextMatrix(GridMaquinasPrev.Row, iGrid_NomeRedMaquinaPrev_Col) = objMaquinas.sNomeReduzido
            
        'verifica se precisa preencher o grid com uma nova linha
        If GridMaquinasPrev.Row - GridMaquinasPrev.FixedRows = objGridMaquinasPrev.iLinhasExistentes Then
            objGridMaquinasPrev.iLinhasExistentes = objGridMaquinasPrev.iLinhasExistentes + 1
        End If
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195183

    Saida_Celula_NomeRedMaquinaPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_NomeRedMaquinaPrev:

    Saida_Celula_NomeRedMaquinaPrev = gErr

    Select Case gErr
        
        Case 195181, 195183
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 195182
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA", gErr, objMaquinas.sNomeReduzido, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195184)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HorasMaquinaPrev(objGridInt As AdmGrid) As Long
'faz a critica da celula de Horas do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim dHoras As Double
Dim iQtd As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As New ClassMaquinas
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_HorasMaquinaPrev

    Set objGridInt.objControle = HorasMaquinaPrev

    'verifica se a Horas está preenchida
    If Len(Trim(HorasMaquinaPrev.Text)) > 0 Then

        'verifica se a Horas é válida
        lErro = Valor_Positivo_Critica(HorasMaquinaPrev.Text)
        If lErro <> SUCESSO Then gError 195185
        
        HorasMaquinaPrev.Text = Formata_Estoque(StrParaDbl(HorasMaquinaPrev.Text))
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195186

    Saida_Celula_HorasMaquinaPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_HorasMaquinaPrev:

    Saida_Celula_HorasMaquinaPrev = gErr

    Select Case gErr

        Case 195185, 195186
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195187)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantMaquinaPrev(objGridInt As AdmGrid) As Long
'faz a critica da celula de QuantidadeMaquina do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim dHoras As Double
Dim iQtd As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As New ClassMaquinas
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_QuantMaquinaPrev

    Set objGridInt.objControle = QuantMaquinaPrev

    'verifica se a QuantidadeMaquina está preenchida
    If Len(Trim(QuantMaquinaPrev.Text)) > 0 Then

        'verifica se a QuantidadeMaquina é válida
        lErro = Valor_Positivo_Critica(QuantMaquinaPrev.Text)
        If lErro <> SUCESSO Then gError 195188
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195189

    Saida_Celula_QuantMaquinaPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantMaquinaPrev:

    Saida_Celula_QuantMaquinaPrev = gErr

    Select Case gErr

        Case 195188, 195189
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195190)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodigoMOPrev(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodTipoMO As String
Dim iLinha As Integer
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_CodigoMOPrev

    Set objGridInt.objControle = CodigoMOPrev
                    
    'Se o campo foi preenchido
    If Len(CodigoMOPrev.Text) > 0 Then

        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridMOPrev.Row Then
                                                    
                If GridMOPrev.TextMatrix(iLinha, iGrid_CodigoMOPrev_Col) = CodigoMOPrev.Text Then
                    sCodTipoMO = CodigoMOPrev.Text
                    CodigoMOPrev.Text = ""
                    gError 134130
                    
                End If
                    
            End If
                           
        Next
        
        Set objTiposDeMaodeObra = New ClassTiposDeMaodeObra
        
        objTiposDeMaodeObra.iCodigo = StrParaInt(CodigoMOPrev.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 135029
    
        If lErro = SUCESSO Then

            GridMOPrev.TextMatrix(GridMOPrev.Row, iGrid_NomeRedMOPrev_Col) = objTiposDeMaodeObra.sDescricao
            
            'verifica se precisa preencher o grid com uma nova linha
            If GridMOPrev.Row - GridMOPrev.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
        
        Else
        
            CodigoMOPrev.Text = ""
            gError 137935
        
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134132

    Saida_Celula_CodigoMOPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigoMOPrev:

    Saida_Celula_CodigoMOPrev = gErr

    Select Case gErr

        Case 134130
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, sCodTipoMO, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 137935
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134128, 134129, 134131, 134132
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162620)

    End Select

    Exit Function
    
End Function

Private Function Saida_Celula_HorasMOPrev(objGridInt As AdmGrid) As Long
'faz a critica da celula de Horas do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMO As ClassCTOperadores
Dim iMO As Integer
Dim dHoras As Double
Dim vbMsgBox As VbMsgBoxResult
Dim bAchou As Boolean

On Error GoTo Erro_Saida_Celula_HorasMOPrev

    Set objGridInt.objControle = HorasMOPrev

    'verifica se a Horas está preenchida
    If Len(Trim(HorasMOPrev.Text)) > 0 Then

        'verifica se a Horas é válida
        lErro = Valor_Positivo_Critica(HorasMOPrev.Text)
        If lErro <> SUCESSO Then gError 195193

        HorasMOPrev.Text = Formata_Estoque(StrParaDbl(HorasMOPrev.Text))

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195194

    Saida_Celula_HorasMOPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_HorasMOPrev:

    Saida_Celula_HorasMOPrev = gErr

    Select Case gErr

        Case 195193, 195194
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195195)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub GridMaquinasPrev_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaquinasPrev, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinasPrev, iAlterado)
    End If

End Sub


Private Sub GridMOPrev_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMOPrev, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMOPrev, iAlterado)
    End If

End Sub

Private Sub GridMOPrev_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMOPrev)

End Sub

Private Sub GridMOPrev_EnterCell()

    Call Grid_Entrada_Celula(objGridMOPrev, iAlterado)

End Sub

Private Sub GridMOPrev_LeaveCell()
    
    Call Saida_Celula(objGridMOPrev)

End Sub

Private Sub GridMOPrev_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMOPrev, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMOPrev, iAlterado)
    End If

End Sub

Private Sub GridMOPrev_RowColChange()

    Call Grid_RowColChange(objGridMOPrev)

End Sub

Private Sub GridMOPrev_Scroll()

    Call Grid_Scroll(objGridMOPrev)

End Sub

Private Sub GridMOPrev_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridMOPrev)
        
End Sub

Private Function Move_OperacaoMO_Memoria(objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemOSOperacoesMO As ClassItemOSOperacoesMO
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Move_OperacaoMO_Memoria

    'Limpa coleção para reiniciar
    Set objItemOSOperacoes.colMO = New Collection

        For iIndice2 = 1 To objGridMOPrev.iLinhasExistentes
        
            Set objItemOSOperacoesMO = New ClassItemOSOperacoesMO
        
            objItemOSOperacoesMO.lCodigoMO = StrParaInt(GridMOPrev.TextMatrix(iIndice2, iGrid_CodigoMOPrev_Col))
            objItemOSOperacoesMO.dHoras = StrParaDbl(GridMOPrev.TextMatrix(iIndice2, iGrid_HorasMOPrev_Col))
        
            objItemOSOperacoes.colMO.Add objItemOSOperacoesMO
                
            If objItemOSOperacoesMO.lCodigoMO = 0 Then gError 195118
            If objItemOSOperacoesMO.dHoras = 0 Then gError 195119
        
        Next

    Move_OperacaoMO_Memoria = SUCESSO

    Exit Function

Erro_Move_OperacaoMO_Memoria:

    Move_OperacaoMO_Memoria = gErr

    Select Case gErr

        Case 195118
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOMO_NAO_PREENCHIDO_GRID2", gErr, iIndice2)

        Case 195119
            Call Rotina_Erro(vbOKOnly, "ERRO_HORASMO_NAO_PREENCHIDO_GRID2", gErr, iIndice2)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195120)

    End Select

    Exit Function

End Function

Private Function Move_OperacaoMaquinas_Memoria(objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemOSOperacoesMaquinas As ClassItemOSOperacoesMaquinas
Dim objItemOPMO As ClassItemOPOperacoesMO
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Move_OperacaoMaquinas_Memoria

    'Limpa coleção para reiniciar
    Set objItemOSOperacoes.colMaquinas = New Collection

    'Ir preenchendo a colecao no objOperacoes com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridMaquinasPrev.iLinhasExistentes
    
        Set objItemOSOperacoesMaquinas = New ClassItemOSOperacoesMaquinas
        Set objMaquina = New ClassMaquinas
        
        objMaquina.sNomeReduzido = GridMaquinasPrev.TextMatrix(iIndice, iGrid_NomeRedMaquinaPrev_Col)
    
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 195113
        
        If lErro <> SUCESSO Then gError 195199
        
        objItemOSOperacoesMaquinas.dHoras = StrParaDbl(GridMaquinasPrev.TextMatrix(iIndice, iGrid_HorasMaquinaPrev_Col))
        objItemOSOperacoesMaquinas.iQuantidade = StrParaInt(GridMaquinasPrev.TextMatrix(iIndice, iGrid_QuantMaquinaPrev_Col))
        objItemOSOperacoesMaquinas.lNumIntDocMaq = objMaquina.lNumIntDoc
        
        If objMaquina.sNomeReduzido = "" Then gError 195114
        If objItemOSOperacoesMaquinas.dHoras = 0 Then gError 195115
        If objItemOSOperacoesMaquinas.iQuantidade = 0 Then gError 195116
        
        objItemOSOperacoes.colMaquinas.Add objItemOSOperacoesMaquinas
      
    Next

    Move_OperacaoMaquinas_Memoria = SUCESSO

    Exit Function

Erro_Move_OperacaoMaquinas_Memoria:

    Move_OperacaoMaquinas_Memoria = gErr

    Select Case gErr
    
        Case 195113
        
        Case 195114
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 195115
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAS_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 195116
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 195199
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195117)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMaquinasPrev(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim iIndice1 As Integer
Dim objItemOSOperacoesMaquinas As ClassItemOSOperacoesMaquinas
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Preenche_GridMaquinasPrev
    
    Call Grid_Limpa(objGridMaquinasPrev)
    
    'Exibe os dados da coleção na tela
    For Each objItemOSOperacoesMaquinas In objItemOSOperacoes.colMaquinas
    
        Set objMaquina = New ClassMaquinas
        
        objMaquina.lNumIntDoc = objItemOSOperacoesMaquinas.lNumIntDocMaq
    
        'Le a Máquina atraves do NumIntDoc
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 195196
        
        'Se nao encontrou => Erro
        If lErro = 106353 Then gError 195197
            
        iIndice1 = iIndice1 + 1
    
        GridMaquinasPrev.TextMatrix(iIndice1, iGrid_NomeRedMaquinaPrev_Col) = objMaquina.sNomeReduzido
        GridMaquinasPrev.TextMatrix(iIndice1, iGrid_HorasMaquinaPrev_Col) = Formata_Estoque(objItemOSOperacoesMaquinas.dHoras)
        GridMaquinasPrev.TextMatrix(iIndice1, iGrid_QuantMaquinaPrev_Col) = objItemOSOperacoesMaquinas.iQuantidade
    
    Next

    objGridMaquinasPrev.iLinhasExistentes = iIndice1
    
    Preenche_GridMaquinasPrev = SUCESSO
    
    Exit Function

Erro_Preenche_GridMaquinasPrev:

    Preenche_GridMaquinasPrev = gErr

    Select Case gErr

        Case 195196
        
        Case 195197
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195198)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMOPrev(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim iIndice2 As Integer
Dim objItemOSOperacoesMO As ClassItemOSOperacoesMO
Dim objMO As ClassTiposDeMaodeObra

On Error GoTo Erro_Preenche_GridMOPrev
    
    Call Grid_Limpa(objGridMOPrev)
        
    For Each objItemOSOperacoesMO In objItemOSOperacoes.colMO
    
        Set objMO = New ClassTiposDeMaodeObra
        
        objMO.iCodigo = objItemOSOperacoesMO.lCodigoMO
    
        'Lê os dados da mao de obra a ser excluido
        lErro = CF("TiposDeMaodeObra_Le", objMO)
        If lErro <> SUCESSO And lErro <> 193817 Then gError 195200
    
        'Verifica se mao de obra não está cadastrado
        If lErro <> SUCESSO Then gError 195201
    
        iIndice2 = iIndice2 + 1
    
        GridMOPrev.TextMatrix(iIndice2, iGrid_HorasMOPrev_Col) = Formata_Estoque(objItemOSOperacoesMO.dHoras)
        GridMOPrev.TextMatrix(iIndice2, iGrid_CodigoMOPrev_Col) = objMO.iCodigo
        GridMOPrev.TextMatrix(iIndice2, iGrid_NomeRedMOPrev_Col) = objMO.sDescricao
    
    Next

    objGridMOPrev.iLinhasExistentes = iIndice2
    
    Preenche_GridMOPrev = SUCESSO
    
    Exit Function

Erro_Preenche_GridMOPrev:

    Preenche_GridMOPrev = gErr

    Select Case gErr

        Case 195200
        
        Case 195201
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objMO.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195202)

    End Select

    Exit Function

End Function

Private Sub CodigoOS_Change()

    iAlterado = REGISTRO_ALTERADO
    iCodigoAlterado = REGISTRO_ALTERADO
    
    Call Limpa_NumIntDoc

End Sub

Private Sub Limpa_NumIntDoc()
    Dim iLinha As Integer
    If gobjOS.colItens.Count > 0 Then
        If UCase(CodigoOS.Text) <> UCase(gobjOS.colItens.Item(1).sCodigo) Then
            For iLinha = 1 To objGrid.iLinhasExistentes
                GridMovimentos.TextMatrix(iLinha, iGrid_NumIntDoc_Col) = ""
            Next
        End If
    End If
End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub ItemOS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_ItemOS_Validate

    'Verifica se ItemOS está preenchida
    If Len(Trim(ItemOS.ClipText)) <> 0 Then

        'Critica o ItemOP
        lErro = Inteiro_Critica(ItemOS.Text)
        If lErro <> SUCESSO Then gError 195203
        
        iItem = StrParaInt(ItemOS.Text)
        
        If giItemOS <> iItem Then
        
            'Se o valor estiver fora do range do grid... Erro
            If iItem < 1 Or iItem > objGrid.iLinhasExistentes Then gError 195204
            
            GridMovimentos.Row = iItem
            
            'Então, mostra a nova arvore
            lErro = Mostra_Arvore(GridMovimentos.TextMatrix(iItem, iGrid_Servico_Col), StrParaDbl(GridMovimentos.TextMatrix(iItem, iGrid_Quantidade_Col)))
            If lErro <> SUCESSO Then gError 195205
            
            giItemOS = iItem
            
        End If

    End If

    Exit Sub

Erro_ItemOS_Validate:

    Cancel = True

    Select Case gErr

        Case 195203, 195205
            
        Case 195204
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195206)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCcls_Click()
'chama tela de Lista de Ccl

Dim lErro As Long
Dim objCcls As New ClassCcl
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCcls_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 195207

    'Verifica se o Produto está preenchido
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col))) = 0 Then gError 195208

    Call Chama_Tela("CclLista", colSelecao, objCcls, objEventoCcl)
    
    Exit Sub
    
Erro_BotaoCcls_Click:

    Select Case gErr
    
        Case 195207
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195208
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195209)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'implementa gravação de uma nova ou atualizacao de uma OS

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Rotina de gravação da OP
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 195210

    'limpa a tela
    lErro = Limpa_Tela_OS
    If lErro <> SUCESSO Then gError 195211
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 195210, 195211

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195212)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objOS As New ClassOS
Dim sPedidoDeVenda As String
Dim vbMsg As VbMsgBoxResult
Dim sStatusItem As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim bAchou As Boolean
'Incluido por Jorge Specian
'------------------------------
Dim bEstaNoPMP As Boolean
Dim objPMPItem As ClassPMPItens
'------------------------------

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    bAchou = False
    
    If Len(Trim(CodigoOS.Text)) = 0 Then gError 195033

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 195034

    'Verifica se o codigo de solicitacao de servico foi preenchido
    If Len(Trim(CodigoSolSrv.ClipText)) = 0 Then gError 195510

    'Verifica se o Produto foi preenchido
    If Len(Trim(ProdutoSolServ.ClipText)) = 0 Then gError 195511

    If objGrid.iLinhasExistentes = 0 Then gError 195035
    
    'Loop de Validação dos dados do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a quantidade foi digitada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 195036

        'Verifica se a data de previsão de inicio da OS foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col))) = 0 Then gError 195037

        'Verifica se a data de previsão de fim da OP foi informada e se é menor que a data de início
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col))) > 0 Then
            If CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col)) < CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col)) Then gError 195038
        Else
            gError 195039
        End If

        'Verifica se o Status é Baixado para ítens novos
        sStatusItem = GridMovimentos.TextMatrix(iIndice, iGrid_StatusItem_Col)
        If Len(Trim(sStatusItem)) > 0 Then
            If UCase(sStatusItem) = UCase(STRING_BAIXADA) And gcolItemOS.Item(iIndice) = 0 Then gError 195040
        End If 'gobjOS.colItens
        
        'Verifica se é uma ordem de servico baixada
        If StatusOS.Caption = STRING_STATUS_BAIXADO Then
            
            'Verifica se o Status é Normal para o item
            If UCase(sStatusItem) <> UCase(STRING_BAIXADA) Then bAchou = True
                
        End If
        
    
    Next
    
    'se a OS está baixada e existe item com situacao='normal'
    If bAchou = True Then
    
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_REATIVACAO_OS", CodigoOS.Text)
        'se não for reativar a OS sai da gravação
        If vbMsg = vbNo Then gError 195041
    
    ElseIf bAchou = False And StatusOS.Caption = STRING_STATUS_BAIXADO Then
        gError 195042
    End If
    
    lErro = Move_Tela_Memoria(objOS)
    If lErro <> SUCESSO Then gError 195043

    lErro = CF("OS_Grava", objOS)
    If lErro <> SUCESSO Then gError 195044

    'Se a opcao de imprimir o Relatorio estiver marcada
    If ImprimeAoGravar.Value = MARCADO Then
        
        'Gera o(s) Relatorio(s)
        lErro = Executa_Relatorio(objOS)
        If lErro <> SUCESSO Then gError 195045
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 195033
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOOS_NAO_PREENCHIDO", gErr)

        Case 195034
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 195035
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEMOS_INFORMADO", gErr)

        Case 195036
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 195037
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_NAO_PREENCHIDA", gErr, iIndice)

        Case 195038
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_MENOR_OS", gErr, iIndice)

        Case 195039
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr, iIndice)

        Case 195040
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAR_ITEMNOVO_OS", gErr)
        
        Case 195041, 195043 To 195045
        
        Case 195042
            Call Rotina_Erro(vbOKOnly, "ERRO_OSBAIXADA_NAO_REATIVADA", gErr)
        
        Case 195510
            Call Rotina_Erro(vbOKOnly, "ERRO_CODSOLSRV_NAO_PREENCHIDO", gErr)
        
        Case 195511
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_OS_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195046)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'ativa a exclusão de uma OS

Dim lErro As Long
Dim objOS As New ClassOS
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se código está preenchido
    If Len(Trim(CodigoOS.Text)) = 0 Then gError 195313

    objOS.sCodigo = CodigoOS.Text
    objOS.iFilialEmpresa = giFilialEmpresa

    'Pede ao usuário que confire a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_OS", objOS.sCodigo, objOS.iFilialEmpresa)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'exclui a OS
    lErro = CF("OS_Exclui", objOS)
    If lErro <> SUCESSO And lErro <> 195312 Then gError 195314

    'se OS não existir -> erro
    If lErro <> SUCESSO Then gError 195315

    'Limpa a tela
    lErro = Limpa_Tela_OS
    If lErro <> SUCESSO Then gError 195316

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 195313
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOOS_NAO_PREENCHIDO", gErr)

        Case 195314, 195316

        Case 195315
            Call Rotina_Erro(vbOKOnly, "ERRO_OS_NAO_CADASTRADA", gErr, objOS.iFilialEmpresa, objOS.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195316)

    End Select

    Exit Sub

End Sub


Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 195317

    'limpa a tela
    lErro = Limpa_Tela_OS
    If lErro <> SUCESSO Then gError 195318

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 195317, 195318

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195319)

    End Select

    Exit Sub

End Sub

Private Sub BotaoServicos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoServicos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 195320

    'Verifica se o Produto está preenchido
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 195321
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza=?"

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico, sSelecaoSQL)
        
        
    Exit Sub

Erro_BotaoServicos_Click:

    Select Case gErr
    
        Case 195320
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195321
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195322)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objItemOS As New ClassItemOS

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    If GridMovimentos.Row <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195323

        'Se o produto não estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 195324

            If lErro = 28030 Then gError 195325

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            'mascara produto escolhido
            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 195326

            'verifica se este servico já foi usado na OS
            lErro = VerificaUso_Produto(objProduto)
            If lErro <> SUCESSO And lErro <> 41316 Then gError 195327

            If lErro = 41316 Then gError 195328

            Call Carrega_ComboVersoes(objProduto.sCodigo)
            
            Servico.PromptInclude = False
            Servico.Text = sProdutoMascarado
            Servico.PromptInclude = True
            
            If Not (Me.ActiveControl Is Servico) Then

                'preenche produto
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Servico_Col) = sProdutoMascarado
    
                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemOS)
                If lErro <> SUCESSO Then gError 195329
    
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 195323, 195324, 195327, 195329

        Case 195325
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 195326
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
            
        Case 195328
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, sProdutoMascarado, CodigoOS.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195330)

    End Select

    Exit Sub

End Sub

Private Sub CodigoOSLabel_Click()

Dim objOS As New ClassOS
Dim colSelecao As New Collection

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(CodigoOS.Text)) <> 0 Then objOS.sCodigo = CodigoOS.Text
    
    'lista as OP's
    Call Chama_Tela("OSLista", colSelecao, objOS, objEventoCodigo)

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objOS As New ClassOS

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(CodigoOS.Text)) = 0 Then gError 195331
    If Len(Trim(Data.ClipText)) = 0 Then gError 195332
    
    objOS.sCodigo = CodigoOS.Text
    objOS.dtDataEmissao = StrParaDate(Data.Text)
    objOS.iFilialEmpresa = giFilialEmpresa
    
    'pesquisa a op nao baixada, e preenche seus itens
    lErro = CF("OS_Le_ComItens", objOS)
    If lErro <> SUCESSO And lErro <> 195060 Then gError 195333
    
    'se nao achou => erro de inexistencia de os
    If lErro <> SUCESSO Then gError 195334
    
    'Executa o(s) Relatorio(s) de acordo com a selecao
    lErro = Executa_Relatorio(objOS)
    If lErro <> SUCESSO Then gError 195335
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 195331
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 195332
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
    
        Case 195333, 195335
        
        Case 195334
            Call Rotina_Erro(vbOKOnly, "ERRO_OS_NAO_CADASTRADA", gErr, objOS.iFilialEmpresa, objOS.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195336)

    End Select
    
End Sub

Private Function Executa_Relatorio(ByVal objOS As ClassOS) As Long
'Executa o(s) relatorio(s) de acordo com a selecao no frame de relatorios

Dim lErro As Long, lNumIntRel As Long
Dim objRelatorio1 As New AdmRelatorio, objRelatorio2 As New AdmRelatorio

On Error GoTo Erro_Executa_Relatorio

'    lErro = CF("ItensOSRel_Prepara", objOS, lNumIntRel)
'    If lErro <> SUCESSO Then gError 195337

    '#####################################
    'Alterado por Wagner
    'Imprime o Relatorio de OP
'    If OpcaoSimples.Value = True Then
'        lErro = objRelatorio1.ExecutarDireto("Ordens de Serviço", "OS = @TOS", 0, "OSINPAL", "TOS", objOS.sCodigo, "NNUMINTREL", CStr(lNumIntRel))
'    Else
'        lErro = objRelatorio1.ExecutarDireto("Ordens de Serviço Detalhado", "OS = @TOS", 0, "OSINPALD", "TOS", objOS.sCodigo, "NNUMINTREL", CStr(lNumIntRel))
'    End If
'    If lErro <> SUCESSO Then gError 195338
    
    lErro = objRelatorio1.ExecutarDireto("Ordem de Serviço", "Codigo = @TOSDE", 0, IIf(OpcaoDetalhado.Value, "OSDet", "OS"), "TOSDE", objOS.sCodigo, "TOSATE", objOS.sCodigo, "TPRODINIC", "", "TPRODFIM", "", "TSSDE", "", "TSSATE", "", "DINIC", DATA_NULA, "DFIM", DATA_NULA, "DENTINIC", DATA_NULA, "DENTFIM", DATA_NULA, "NSTATUS", 0)
    If lErro <> SUCESSO Then gError 195338
    '#####################################
    
    Executa_Relatorio = SUCESSO
    
    Exit Function
    
Erro_Executa_Relatorio:

    Executa_Relatorio = gErr
    
    Select Case gErr
    
        Case 195337, 195338
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195339)
    
    End Select

End Function

Private Sub BotaoImprimirOper_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimirOper_Click

    If Len(Trim(CodigoOS.Text)) <> 0 Then

        lErro = objRelatorio.ExecutarDireto("OS - Operações", "", 0, "", "TCODOS", CodigoOS.Text)
        If lErro <> SUCESSO Then gError 195340
    
    End If
    
    Exit Sub
    
Erro_BotaoImprimirOper_Click:
    
    Select Case gErr
    
        Case 195340
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195341)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Traz_Dados_Pedido_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim objSolicSRV As New ClassSolicSRV
Dim lErro As Long
Dim objItemSolicSRV As New ClassItensSolicSRV
Dim sProdutoEnxuto As String
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim objItemOSOperacoes As ClassItemOSOperacoes
Dim objItemOS As ClassItemOS
Dim objItemOSOperacoesPecas As ClassItemOSOperacoesPecas
Dim iIndice1 As Integer
Dim sPS As String
Dim iIndex As Integer

On Error GoTo Erro_Traz_Dados_Pedido_Tela

    sPS = PedSRV.Text
    iIndex = FilialPedSRV.ListIndex

    lErro = Limpa_Tela_OS
    If lErro <> SUCESSO Then gError 195733
    
    PedSRV.PromptInclude = False
    PedSRV.Text = sPS
    PedSRV.PromptInclude = True
    FilialPedSRV.ListIndex = iIndex

    If objPedidoVenda.lNumIntSolicSRV <> 0 Then
        
        objSolicSRV.lNumIntDoc = objPedidoVenda.lNumIntSolicSRV
        
        lErro = CF("SolicitacaoSRV_Le_NumIntDoc", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 186988 Then gError 195727

        If lErro = SUCESSO Then
            CodigoSolSrv.Text = objSolicSRV.lCodigo
        End If
        
        Set objItemSolicSRV = objSolicSRV.colItens(1)
       
        lErro = Mascara_RetornaProdutoEnxuto(objItemSolicSRV.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195728

        'Mascara o produto enxuto
        ProdutoSolServ.PromptInclude = False
        ProdutoSolServ.Text = sProdutoEnxuto
        ProdutoSolServ.PromptInclude = True
            
        Lote.Text = objItemSolicSRV.sLote
        If objItemSolicSRV.iFilialOP <> 0 Then FilialOP.Text = objItemSolicSRV.iFilialOP
            
    End If

    For iIndice = 1 To objPedidoVenda.colItensPedidoSRV.Count

        'Mascara produto
        lErro = Mascara_RetornaProdutoTela(objPedidoVenda.colItensPedidoSRV(iIndice).sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 195729

        Servico.PromptInclude = False
        Servico.Text = sProdutoMascarado
        Servico.PromptInclude = True

        GridMovimentos.TextMatrix(iIndice, iGrid_Servico_Col) = sProdutoMascarado

        objProduto.sCodigo = objPedidoVenda.colItensPedidoSRV(iIndice).sProduto

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195730

        GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objProduto.sDescricao

        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objPedidoVenda.colItensPedidoSRV(iIndice).sUnidadeMed
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objPedidoVenda.colItensPedidoSRV(iIndice).dQuantidade - objPedidoVenda.colItensPedidoSRV(iIndice).dQuantCancelada)

        objProduto.sCodigo = objPedidoVenda.colItensPedidoSRV(iIndice).sProduto
               
'        For Each objItemOS In gobjOS.colItens
'
'            If objItemOS.sServico = objPedidoVenda.colItensPedidoSRV(iIndice).sProduto And _
'                objItemOS.sVersao = GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col) Then
'
'                Set objItemOSOperacoes = objItemOS.colItemOSOperacoes(1)
'
'                For iIndice1 = 1 To objPedidoVenda.colItensPedido.Count
'
'                    Set objItemOSOperacoesPecas = New ClassItemOSOperacoesPecas
'
'                    objItemOSOperacoesPecas.sProduto = objPedidoVenda.colItensPedido(iIndice1).sProduto
'                    objItemOSOperacoesPecas.sUMProduto = objPedidoVenda.colItensPedido(iIndice1).sUnidadeMed
'                    objItemOSOperacoesPecas.dQuantidade = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice1).dQuantidade - objPedidoVenda.colItensPedido(iIndice1).dQuantCancelada)
'
'                    objItemOSOperacoes.colPecas.Add objItemOSOperacoesPecas
'
'                Next
'
'                Exit For
'
'            End If
'
'        Next
        

    Next
    
    objGrid.iLinhasExistentes = objPedidoVenda.colItensPedidoSRV.Count

    Set gobjOS = New ClassOS
    Set gcolItemOS = New Collection
    
    lErro = Move_Tela_Memoria(gobjOS)
    If lErro <> SUCESSO Then gError 195730
    
    For iIndice = 1 To gobjOS.colItens.Count
        gcolItemOS.Add 0
    Next

    If gobjOS.colItens.Count > 0 Then
        lErro = Trata_Arvore(GridMovimentos.TextMatrix(1, iGrid_Servico_Col), GridMovimentos.TextMatrix(1, iGrid_Versao_Col), GridMovimentos.TextMatrix(1, iGrid_UnidadeMed_Col), StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError 195731
    End If

    Traz_Dados_Pedido_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Dados_Pedido_Tela:

    Traz_Dados_Pedido_Tela = gErr
    
    Select Case gErr

        Case 195727, 195731, 195730, 195733

        Case 195728
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemSolicSRV.sProduto)

        Case 195729
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195732)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 195734

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOP.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOP.ItemData(FilialOP.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 195734

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195735)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialPedSRV() As Long
'Carrega FilialPedido com as Filiais Empresas

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Carrega_FilialPedSRV

    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA Then
            'coloca na combo
            FilialPedSRV.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialPedSRV.ItemData(FilialPedSRV.NewIndex) = objFiliais.iCodFilial
        End If
    Next

    'Seleciona a Filial na qual o usuário entrou no Sistema
    Call Seleciona_FilialEmpresa

    Carrega_FilialPedSRV = SUCESSO

    Exit Function

Erro_Carrega_FilialPedSRV:

    Carrega_FilialPedSRV = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195755)

    End Select

    Exit Function

End Function

Private Sub Seleciona_FilialEmpresa()

Dim iIndice As Integer

    'Se a filial da empresa não for EMPRESA_TODA
    If giFilialEmpresa <> EMPRESA_TODA Then
        'Seleciona na combo de filial do pedido a filial empresa
        For iIndice = 0 To FilialPedSRV.ListCount - 1
            If FilialPedSRV.ItemData(iIndice) = giFilialEmpresa Then
                FilialPedSRV.ListIndex = iIndice
                Exit For
            End If
        Next

    Else
        FilialPedSRV.ListIndex = 0
    End If

End Sub

Function AlteraQuantidade(ByVal objItemOS As ClassItemOS, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim dFatorQtdeOP As Double
Dim objItemOSOperacoes As ClassItemOSOperacoes
Dim objItemOSOpMO As ClassItemOSOperacoesMO, objItemOSOpMaq As ClassItemOSOperacoesMaquinas, objItemOSOpMP As ClassItemOSOperacoesPecas

On Error GoTo Erro_AlteraQuantidade

    dFatorQtdeOP = dQuantidade / objItemOS.dQuantidade
                            
    For Each objItemOSOperacoes In objItemOS.colItemOSOperacoes
                       
        For Each objItemOSOpMP In objItemOSOperacoes.colPecas
            'If objItemOSOpMP.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                objItemOSOpMP.dQuantidade = objItemOSOpMP.dQuantidade * dFatorQtdeOP
            'End If
        Next
        
        For Each objItemOSOpMaq In objItemOSOperacoes.colMaquinas
            objItemOSOpMaq.dHoras = objItemOSOpMaq.dHoras * dFatorQtdeOP
        Next
        
        For Each objItemOSOpMO In objItemOSOperacoes.colMO
            objItemOSOpMO.dHoras = objItemOSOpMO.dHoras * dFatorQtdeOP
        Next
                
    Next
    
    objItemOS.dQuantidade = dQuantidade
    
    AlteraQuantidade = SUCESSO
    
    Exit Function

Erro_AlteraQuantidade:

    AlteraQuantidade = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163897)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoIncluir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim sChave As String
Dim sTexto As String
Dim objNode As Node
Dim objNodePai As Node
Dim sChaveTvw As String
Dim iNivel As Integer
Dim objOperacoes As New ClassItemOSOperacoes
Dim objCompetencias As New ClassCompetencias
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoIncluir_Click

    If Len(Trim(CodigoCompetencia.ClipText)) = 0 Then gError 194960

    lErro = Move_Operacoes_Memoria(objOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 194961

    Set objNodePai = Roteiro.SelectedItem

    If objNodePai Is Nothing Then
        iNivel = 0
    Else
        If objNodePai.Selected = False Then gError 194962
        iNivel = objNodePai.Index + 1
    End If

    'prepara texto que identificará a nova Operação que está sendo incluida
    sTexto = objCompetencias.sNomeReduzido
    
    sCodProduto = Servico.Text

    'Critica o formato do MPProduto e se existe no BD
    lErro = CF("Produto_Critica", sCodProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 194963
            
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
    
    objOperacoes.sServico = gobjItemOS.sServico
    objOperacoes.sVersao = gobjItemOS.sVersao
    objOperacoes.sUMedida = gobjItemOS.sSiglaUM
    objOperacoes.dQuantidade = gobjItemOS.dQuantidade
   
    colComponentes.Add objOperacoes, sChave
    objNode.Tag = sChave

    Set gobjItemOS.colItemOSOperacoes = colComponentes

    Call Recalcula_Nivel_Sequencial

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 194976

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 194960
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)

        Case 194961, 194963, 194976
            'erro tratado na rotina chamada

        Case 194962
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194977)

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

    If objNode Is Nothing Then gError 194978
    If objNode.Selected = False Then gError 194979

    If objNode.Children > 0 Then

        'Envia aviso perguntando se realmente deseja excluir
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_TEM_FILHOS")

        If vbMsgRes = vbNo Then gError 194987

        'chama rotina que exclui filhos
        Call Remove_Filhos(objNode.Child)
    
    End If

    colComponentes.Remove (objNode.Tag)
    Roteiro.Nodes.Remove (objNode.Index)
    
    Set gobjItemOS.colItemOSOperacoes = colComponentes
    
    Call Recalcula_Nivel_Sequencial

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 194988

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoRemover_Click:

    Select Case gErr

        Case 194978, 194979
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 194987, 194988

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194989)

    End Select

    Exit Sub

End Sub

Sub Remove_Filhos(objNode As Node)
'remove objNode, seus irmaos e filhos de colComponentes

    colComponentes.Remove (objNode.Tag)

    If objNode.Children > 0 Then

        Call Remove_Filhos(objNode.Child)

    End If

    If objNode <> objNode.LastSibling Then Call Remove_Filhos(objNode.Next)

    Exit Sub

End Sub

Public Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Tipo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_Tipo(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Tipo

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPOOS, objComboBox)
    If lErro <> SUCESSO Then gError 186769

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    glTipoPadrao = Tipo.ListIndex

    Carrega_Tipo = SUCESSO

    Exit Function

Erro_Carrega_Tipo:

    Carrega_Tipo = gErr

    Select Case gErr

        Case 186769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186770)

    End Select

    Exit Function

End Function
