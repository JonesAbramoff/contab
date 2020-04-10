VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrdemProducao 
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   9600
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "'"
      Height          =   5145
      Index           =   1
      Left            =   165
      TabIndex        =   43
      Top             =   795
      Width           =   9285
      Begin VB.Frame Frame5 
         Caption         =   "Relatórios"
         Height          =   1230
         Left            =   150
         TabIndex        =   129
         Top             =   3810
         Width           =   8985
         Begin VB.CheckBox ImprimeAoGravar 
            Caption         =   "Imprimir ao Gravar uma O.P."
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
            Left            =   4575
            TabIndex        =   28
            Top             =   795
            Width           =   2715
         End
         Begin VB.OptionButton OpcaoRelatorio 
            Caption         =   "Imprimir Ambos."
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
            Left            =   630
            TabIndex        =   24
            Top             =   945
            Value           =   -1  'True
            Width           =   1665
         End
         Begin VB.OptionButton OpcaoRelatorio 
            Caption         =   "Imprimir Rótulos para Produção."
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
            Left            =   630
            TabIndex        =   23
            Top             =   585
            Width           =   3045
         End
         Begin VB.OptionButton OpcaoRelatorio 
            Caption         =   "Imprimir Ordem de Produção."
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
            Left            =   630
            TabIndex        =   22
            Top             =   225
            Width           =   2775
         End
         Begin VB.Frame Frame4 
            Caption         =   "Ordem de Produção"
            Height          =   660
            Left            =   4485
            TabIndex        =   130
            Top             =   135
            Width           =   4275
            Begin VB.OptionButton OpcaoAmbos 
               Caption         =   "Ambos"
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
               Left            =   2655
               TabIndex        =   27
               Top             =   315
               Width           =   1215
            End
            Begin VB.OptionButton OpcaoDetalhado 
               Caption         =   "Espelho"
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
               TabIndex        =   26
               Top             =   315
               Width           =   1215
            End
            Begin VB.OptionButton OpcaoSimples 
               Caption         =   "Simples"
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
               TabIndex        =   25
               Top             =   300
               Value           =   -1  'True
               Width           =   1065
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem de Produção"
         Height          =   2220
         Left            =   150
         TabIndex        =   120
         Top             =   90
         Width           =   8985
         Begin VB.CheckBox GeraOPs 
            Caption         =   "Gera uma O.P. para cada subproduto"
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
            Height          =   270
            Left            =   540
            TabIndex        =   9
            Top             =   1335
            Width           =   3720
         End
         Begin VB.TextBox CodOPGerada 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2820
            MaxLength       =   9
            TabIndex        =   5
            Top             =   990
            Width           =   1305
         End
         Begin VB.CheckBox GeraReqCompra 
            Caption         =   "Gera Req. de Compras para matérias-primas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5370
            TabIndex        =   7
            Top             =   675
            Width           =   2910
         End
         Begin VB.CheckBox GeraOP 
            Caption         =   "Gera O.P. para subprodutos"
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
            Left            =   195
            TabIndex        =   4
            Top             =   645
            Width           =   3270
         End
         Begin VB.TextBox Codigo 
            Height          =   285
            Left            =   1365
            MaxLength       =   9
            TabIndex        =   0
            Top             =   285
            Width           =   1350
         End
         Begin VB.TextBox PrestadorServico 
            Height          =   285
            Left            =   5970
            MaxLength       =   6
            TabIndex        =   10
            Top             =   1320
            Width           =   2895
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2715
            Picture         =   "OrdemProducaoCro.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   285
            Width           =   300
         End
         Begin VB.CommandButton botaoProxNum2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4125
            Picture         =   "OrdemProducaoCro.ctx":00EA
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Numeração Automática"
            Top             =   990
            Width           =   300
         End
         Begin VB.CommandButton BotaoOPRC 
            Caption         =   "..."
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
            Height          =   395
            Left            =   8355
            TabIndex        =   8
            Top             =   720
            Width           =   495
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
            Left            =   3915
            TabIndex        =   12
            Top             =   1725
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   5970
            TabIndex        =   13
            Top             =   1725
            Width           =   2895
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   5370
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   270
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   4305
            TabIndex        =   2
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
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1380
            TabIndex        =   11
            Top             =   1740
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   6885
            TabIndex        =   128
            Top             =   315
            Width           =   615
         End
         Begin VB.Label StatusOP 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7545
            TabIndex        =   127
            Top             =   255
            Width           =   1305
         End
         Begin VB.Label LabelCodOPGerada 
            AutoSize        =   -1  'True
            Caption         =   "Código O.P. a ser gerada:"
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
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   555
            TabIndex        =   126
            Top             =   1035
            Width           =   2220
         End
         Begin VB.Label CodigoOPLabel 
            AutoSize        =   -1  'True
            Caption         =   "Código O.P.:"
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
            Left            =   225
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   125
            Top             =   330
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
            Left            =   3765
            TabIndex        =   124
            Top             =   330
            Width           =   480
         End
         Begin VB.Label LabelPrestador 
            AutoSize        =   -1  'True
            Caption         =   "Prest. de Serv:"
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
            Left            =   4635
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   123
            Top             =   1365
            Width           =   1290
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
            Index           =   62
            Left            =   5310
            TabIndex        =   122
            Top             =   1785
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
            Left            =   645
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   121
            Top             =   1785
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Padrões "
         Height          =   1440
         Index           =   0
         Left            =   150
         TabIndex        =   113
         Top             =   2340
         Width           =   8985
         Begin VB.ComboBox DestinacaoPadrao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrdemProducaoCro.ctx":01D4
            Left            =   2955
            List            =   "OrdemProducaoCro.ctx":01D6
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1005
            Width           =   1710
         End
         Begin MSMask.MaskEdBox CclPadrao 
            Height          =   315
            Left            =   2955
            TabIndex        =   14
            Top             =   225
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxPadrao 
            Height          =   315
            Left            =   6990
            TabIndex        =   15
            Top             =   225
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownInicio 
            Height          =   300
            Left            =   4050
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   630
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicioPadrao 
            Height          =   300
            Left            =   2955
            TabIndex        =   16
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownFim 
            Height          =   300
            Left            =   8085
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   630
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFimPadrao 
            Height          =   300
            Left            =   6990
            TabIndex        =   18
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrioridadePadrao 
            Height          =   315
            Left            =   6990
            TabIndex        =   21
            Top             =   1005
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label CclPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo/Lucro Padrão:"
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
            Left            =   255
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   119
            Top             =   285
            Width           =   2670
         End
         Begin VB.Label AlmoxPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado Padrão:"
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
            Left            =   5130
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   118
            Top             =   285
            Width           =   1815
         End
         Begin VB.Label DataPrevIniLbl 
            AutoSize        =   -1  'True
            Caption         =   "Data de Previsão de Início:"
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
            Left            =   555
            TabIndex        =   117
            Top             =   690
            Width           =   2370
         End
         Begin VB.Label DataPrevFimLbl 
            AutoSize        =   -1  'True
            Caption         =   "Data de Previsão de Fim:"
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
            Left            =   4785
            TabIndex        =   116
            Top             =   690
            Width           =   2160
         End
         Begin VB.Label DestPadraoLbl 
            AutoSize        =   -1  'True
            Caption         =   "Destinação Padrão:"
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
            Left            =   1230
            TabIndex        =   115
            Top             =   1065
            Width           =   1695
         End
         Begin VB.Label PrioridadePadraoLbl 
            AutoSize        =   -1  'True
            Caption         =   " Prioridade Padrão:"
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
            Left            =   5295
            TabIndex        =   114
            Top             =   1065
            Width           =   1650
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   3
      Left            =   165
      TabIndex        =   271
      Top             =   795
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame Frame6 
         Caption         =   "Dados da Ficha de Processo"
         Height          =   2685
         Left            =   75
         TabIndex        =   286
         Top             =   1950
         Width           =   2985
         Begin VB.Frame Frame12 
            Caption         =   "Mistura"
            Height          =   1245
            Left            =   210
            TabIndex        =   287
            Top             =   315
            Width           =   2715
            Begin VB.ComboBox TempoMistura 
               Height          =   315
               ItemData        =   "OrdemProducaoCro.ctx":01D8
               Left            =   1005
               List            =   "OrdemProducaoCro.ctx":01DA
               TabIndex        =   73
               Top             =   750
               Width           =   1665
            End
            Begin VB.ComboBox Mistura 
               Height          =   315
               ItemData        =   "OrdemProducaoCro.ctx":01DC
               Left            =   1005
               List            =   "OrdemProducaoCro.ctx":01DE
               TabIndex        =   72
               Top             =   300
               Width           =   1665
            End
            Begin VB.Label TempoMisturaLabel 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   345
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   289
               Top             =   810
               Width           =   645
            End
            Begin VB.Label MisturaLabel 
               AutoSize        =   -1  'True
               Caption         =   "Mistura:"
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
               TabIndex        =   288
               Top             =   345
               Width           =   690
            End
         End
         Begin VB.ComboBox CargaDrays 
            Height          =   315
            ItemData        =   "OrdemProducaoCro.ctx":01E0
            Left            =   1215
            List            =   "OrdemProducaoCro.ctx":01E2
            TabIndex        =   74
            Top             =   1755
            Width           =   1635
         End
         Begin VB.ComboBox Embalagem 
            Height          =   315
            ItemData        =   "OrdemProducaoCro.ctx":01E4
            Left            =   1215
            List            =   "OrdemProducaoCro.ctx":01E6
            TabIndex        =   75
            Top             =   2205
            Width           =   1635
         End
         Begin VB.Label EmbalagemLabel 
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
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   291
            Top             =   2235
            Width           =   1035
         End
         Begin VB.Label CargaDraysLabel 
            AutoSize        =   -1  'True
            Caption         =   "Carga Drays:"
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
            TabIndex        =   290
            Top             =   1800
            Width           =   1110
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados da Carga"
         Height          =   1890
         Left            =   75
         TabIndex        =   278
         Top             =   0
         Width           =   9150
         Begin VB.TextBox Descricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   3765
            MaxLength       =   50
            TabIndex        =   280
            Top             =   930
            Width           =   2025
         End
         Begin MSMask.MaskEdBox QtdTotal 
            Height          =   285
            Left            =   4080
            TabIndex        =   279
            Top             =   570
            Width           =   870
            _ExtentX        =   1535
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
         Begin MSMask.MaskEdBox QtdCarga 
            Height          =   285
            Left            =   2685
            TabIndex        =   281
            Top             =   705
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox KgCarga 
            Height          =   285
            Left            =   1980
            TabIndex        =   282
            Top             =   1305
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox Item 
            Height          =   255
            Left            =   2040
            TabIndex        =   283
            Top             =   465
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox KgProdBase 
            Height          =   285
            Left            =   1650
            TabIndex        =   284
            Top             =   930
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox ProdutoBase 
            Height          =   255
            Left            =   585
            TabIndex        =   285
            Top             =   525
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCarga 
            Height          =   1425
            Left            =   60
            TabIndex        =   71
            Top             =   240
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   2514
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Componentes da Carga"
         Height          =   2685
         Left            =   3105
         TabIndex        =   272
         Top             =   1950
         Width           =   6120
         Begin VB.TextBox DescInsumo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            MaxLength       =   50
            TabIndex        =   276
            Top             =   555
            Width           =   1605
         End
         Begin MSMask.MaskEdBox Lote 
            Height          =   255
            Left            =   3780
            TabIndex        =   273
            Top             =   1320
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   255
            Left            =   3495
            TabIndex        =   274
            Top             =   855
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox KgCargaInsumo 
            Height          =   285
            Left            =   1545
            TabIndex        =   275
            Top             =   1425
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox CodInsumo 
            Height          =   255
            Left            =   840
            TabIndex        =   277
            Top             =   1035
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridInsumo 
            Height          =   1800
            Left            =   45
            TabIndex        =   76
            Top             =   255
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3175
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton botaoProdutoBase 
         Caption         =   "Produto"
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
         Left            =   105
         TabIndex        =   77
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton BotaoLimpaCarga 
         Caption         =   "Limpa Grids"
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
         Left            =   5290
         TabIndex        =   79
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton BotaoCarga 
         Caption         =   "Traz Dados da Carga"
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
         Left            =   2450
         TabIndex        =   78
         Top             =   4800
         Width           =   2070
      End
      Begin VB.CommandButton BotaoFornecedor 
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
         Height          =   315
         Left            =   6201
         TabIndex        =   80
         Top             =   4800
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CommandButton BotaoLote 
         Caption         =   "Lote"
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
         Left            =   7770
         TabIndex        =   81
         Top             =   4800
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   4
      Left            =   225
      TabIndex        =   220
      Top             =   810
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Frame Frame9 
         Caption         =   "Processo"
         Height          =   1140
         Left            =   90
         TabIndex        =   253
         Top             =   75
         Width           =   3750
         Begin MSMask.MaskEdBox Agua 
            Height          =   315
            Left            =   1320
            TabIndex        =   82
            Top             =   255
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Oleo 
            Height          =   315
            Left            =   1320
            TabIndex        =   83
            Top             =   675
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Água:"
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
            Index           =   0
            Left            =   435
            TabIndex        =   257
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Óleo:"
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
            Index           =   1
            Left            =   435
            TabIndex        =   256
            Top             =   735
            Width           =   810
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "L"
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
            Index           =   21
            Left            =   2370
            TabIndex        =   255
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "L"
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
            Index           =   22
            Left            =   2385
            TabIndex        =   254
            Top             =   765
            Width           =   165
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Temperatura Camara"
         Height          =   1695
         Left            =   90
         TabIndex        =   243
         Top             =   2730
         Width           =   3780
         Begin MSMask.MaskEdBox TCZonaDe 
            Height          =   315
            Index           =   1
            Left            =   1335
            TabIndex        =   90
            Top             =   735
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TCZonaAte 
            Height          =   315
            Index           =   1
            Left            =   2385
            TabIndex        =   91
            Top             =   750
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TCZonaDe 
            Height          =   315
            Index           =   2
            Left            =   1335
            TabIndex        =   92
            Top             =   1200
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TCZonaAte 
            Height          =   315
            Index           =   2
            Left            =   2385
            TabIndex        =   93
            Top             =   1215
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TCZonaDe 
            Height          =   315
            Index           =   0
            Left            =   1335
            TabIndex        =   88
            Top             =   300
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TCZonaAte 
            Height          =   315
            Index           =   0
            Left            =   2385
            TabIndex        =   89
            Top             =   315
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 2:"
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
            Index           =   12
            Left            =   390
            TabIndex        =   252
            Top             =   795
            Width           =   900
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 3:"
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
            Index           =   13
            Left            =   375
            TabIndex        =   251
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   17
            Left            =   2205
            TabIndex        =   250
            Top             =   795
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   18
            Left            =   3195
            TabIndex        =   249
            Top             =   795
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   19
            Left            =   2205
            TabIndex        =   248
            Top             =   1245
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   20
            Left            =   3210
            TabIndex        =   247
            Top             =   1245
            Width           =   240
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 1:"
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
            Index           =   11
            Left            =   405
            TabIndex        =   246
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   14
            Left            =   2220
            TabIndex        =   245
            Top             =   360
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   16
            Left            =   3195
            TabIndex        =   244
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Temperatura Extrusora"
         Height          =   3120
         Left            =   3990
         TabIndex        =   221
         Top             =   75
         Width           =   4965
         Begin MSMask.MaskEdBox TEZonaDe 
            Height          =   315
            Index           =   0
            Left            =   1185
            TabIndex        =   95
            Top             =   285
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaDe 
            Height          =   315
            Index           =   1
            Left            =   1185
            TabIndex        =   97
            Top             =   690
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaAte 
            Height          =   315
            Index           =   1
            Left            =   2235
            TabIndex        =   98
            Top             =   690
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaDe 
            Height          =   315
            Index           =   2
            Left            =   1185
            TabIndex        =   99
            Top             =   1080
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaAte 
            Height          =   315
            Index           =   2
            Left            =   2220
            TabIndex        =   100
            Top             =   1080
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaDe 
            Height          =   315
            Index           =   3
            Left            =   1185
            TabIndex        =   101
            Top             =   1485
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaAte 
            Height          =   315
            Index           =   3
            Left            =   2235
            TabIndex        =   102
            Top             =   1485
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaDe 
            Height          =   315
            Index           =   4
            Left            =   1185
            TabIndex        =   103
            Top             =   1890
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaAte 
            Height          =   315
            Index           =   4
            Left            =   2235
            TabIndex        =   104
            Top             =   1890
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaDe 
            Height          =   315
            Index           =   5
            Left            =   1185
            TabIndex        =   105
            Top             =   2280
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaAte 
            Height          =   315
            Index           =   5
            Left            =   2235
            TabIndex        =   106
            Top             =   2280
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaDe 
            Height          =   315
            Index           =   6
            Left            =   1185
            TabIndex        =   107
            Top             =   2685
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaAte 
            Height          =   315
            Index           =   6
            Left            =   2235
            TabIndex        =   108
            Top             =   2685
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TEZonaAte 
            Height          =   315
            Index           =   0
            Left            =   2235
            TabIndex        =   96
            Top             =   285
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 2:"
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
            Index           =   6
            Left            =   345
            TabIndex        =   242
            Top             =   735
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 3:"
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
            Index           =   7
            Left            =   345
            TabIndex        =   241
            Top             =   1125
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 4:"
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
            Index           =   8
            Left            =   345
            TabIndex        =   240
            Top             =   1515
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 5:"
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
            Index           =   15
            Left            =   315
            TabIndex        =   239
            Top             =   1935
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 6:"
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
            Index           =   9
            Left            =   345
            TabIndex        =   238
            Top             =   2325
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 7*:"
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
            Index           =   10
            Left            =   345
            TabIndex        =   237
            Top             =   2730
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   30
            Left            =   2070
            TabIndex        =   236
            Top             =   720
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   31
            Left            =   3105
            TabIndex        =   235
            Top             =   735
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   32
            Left            =   2070
            TabIndex        =   234
            Top             =   1140
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   33
            Left            =   3105
            TabIndex        =   233
            Top             =   1140
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   34
            Left            =   2070
            TabIndex        =   232
            Top             =   1530
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   35
            Left            =   3105
            TabIndex        =   231
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   36
            Left            =   2070
            TabIndex        =   230
            Top             =   1950
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   37
            Left            =   3105
            TabIndex        =   229
            Top             =   1950
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   38
            Left            =   2070
            TabIndex        =   228
            Top             =   2340
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   39
            Left            =   3105
            TabIndex        =   227
            Top             =   2340
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   40
            Left            =   2070
            TabIndex        =   226
            Top             =   2745
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   41
            Left            =   3105
            TabIndex        =   225
            Top             =   2745
            Width           =   240
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona 1:"
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
            Index           =   5
            Left            =   330
            TabIndex        =   224
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "a"
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
            Index           =   28
            Left            =   2070
            TabIndex        =   223
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
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
            Index           =   29
            Left            =   3105
            TabIndex        =   222
            Top             =   315
            Width           =   240
         End
      End
      Begin VB.TextBox Amostras 
         Height          =   330
         Left            =   5175
         MaxLength       =   255
         TabIndex        =   111
         Top             =   3690
         Width           =   3780
      End
      Begin VB.TextBox ObservacaoCro 
         Height          =   855
         Left            =   5175
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   112
         Top             =   4110
         Width           =   3795
      End
      Begin MSMask.MaskEdBox Amperagem 
         Height          =   315
         Left            =   1410
         TabIndex        =   84
         Top             =   1425
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Format          =   "#,##0"
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VRDe 
         Height          =   315
         Left            =   1410
         TabIndex        =   85
         Top             =   1860
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VRAte 
         Height          =   315
         Left            =   2460
         TabIndex        =   86
         Top             =   1860
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Gate 
         Height          =   315
         Left            =   1410
         TabIndex        =   87
         Top             =   2310
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "#,##0"
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Tela 
         Height          =   315
         Left            =   1410
         TabIndex        =   94
         Top             =   4560
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VVDe 
         Height          =   315
         Left            =   5175
         TabIndex        =   109
         Top             =   3300
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VVAte 
         Height          =   315
         Left            =   6240
         TabIndex        =   110
         Top             =   3300
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Amperagem:"
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
         Index           =   2
         Left            =   255
         TabIndex        =   270
         Top             =   1515
         Width           =   1080
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Vel. Rotores:"
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
         Index           =   3
         Left            =   165
         TabIndex        =   269
         Top             =   1905
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Abertura Gate:"
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
         Index           =   4
         Left            =   90
         TabIndex        =   268
         Top             =   2325
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "A"
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
         Index           =   23
         Left            =   2505
         TabIndex        =   267
         Top             =   1500
         Width           =   165
      End
      Begin VB.Label Label4 
         Caption         =   "a"
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
         Index           =   24
         Left            =   2295
         TabIndex        =   266
         Top             =   1905
         Width           =   165
      End
      Begin VB.Label Label4 
         Caption         =   "rpm"
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
         Index           =   25
         Left            =   3360
         TabIndex        =   265
         Top             =   1920
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
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
         Index           =   26
         Left            =   2490
         TabIndex        =   264
         Top             =   2355
         Width           =   165
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Telas:"
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
         Index           =   27
         Left            =   690
         TabIndex        =   263
         Top             =   4620
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Vel. Variador:"
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
         Index           =   42
         Left            =   3945
         TabIndex        =   262
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "a"
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
         Index           =   43
         Left            =   6075
         TabIndex        =   261
         Top             =   3360
         Width           =   165
      End
      Begin VB.Label Label4 
         Caption         =   "rpm"
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
         Index           =   44
         Left            =   7125
         TabIndex        =   260
         Top             =   3375
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "R. Amostras:"
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
         Index           =   45
         Left            =   3945
         TabIndex        =   259
         Top             =   3750
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Index           =   46
         Left            =   3930
         TabIndex        =   258
         Top             =   4140
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   5
      Left            =   135
      TabIndex        =   131
      Top             =   795
      Visible         =   0   'False
      Width           =   9360
      Begin VB.Frame Frame2 
         Caption         =   "Mão de Obra"
         Height          =   2880
         Index           =   5
         Left            =   75
         TabIndex        =   155
         Top             =   2265
         Visible         =   0   'False
         Width           =   9195
         Begin VB.CommandButton BotaoMO 
            Caption         =   "Mão de Obra"
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
            Left            =   60
            TabIndex        =   70
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   2460
            Width           =   1305
         End
         Begin VB.TextBox TipoMaoDeObraReal 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   195
            TabIndex        =   158
            Top             =   1410
            Width           =   990
         End
         Begin VB.TextBox DescricaoTipoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1185
            TabIndex        =   157
            Top             =   1425
            Width           =   1515
         End
         Begin VB.ComboBox NomeMaquinaRealMO 
            Height          =   315
            Left            =   2745
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   1440
            Width           =   1890
         End
         Begin MSMask.MaskEdBox HorasRealMO 
            Height          =   315
            Left            =   6000
            TabIndex        =   159
            Top             =   1410
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataRealMO 
            Height          =   315
            Left            =   7350
            TabIndex        =   160
            Top             =   1425
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeMaquinaRealMO 
            Height          =   315
            Left            =   4650
            TabIndex        =   161
            Top             =   1425
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridMOReal 
            Height          =   1785
            Left            =   75
            TabIndex        =   69
            Top             =   255
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   3149
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Máquinas"
         Height          =   2880
         Index           =   4
         Left            =   75
         TabIndex        =   162
         Top             =   2265
         Visible         =   0   'False
         Width           =   9195
         Begin VB.CommandButton BotaoMaquinasReal 
            Caption         =   "Maquinas"
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
            Left            =   60
            TabIndex        =   67
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   2460
            Width           =   1305
         End
         Begin VB.TextBox NomeMaquinaReal 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   180
            TabIndex        =   163
            Top             =   1425
            Width           =   2895
         End
         Begin VB.CommandButton BotaoTrazerMaquinasPrevistas 
            Caption         =   "Trazer Previsto"
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
            Left            =   6840
            TabIndex        =   68
            ToolTipText     =   "Traz as máquinas que tiveram o uso previsto pelo MRP."
            Top             =   2460
            Width           =   2220
         End
         Begin MSMask.MaskEdBox HorasReal 
            Height          =   315
            Left            =   4905
            TabIndex        =   164
            Top             =   1470
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataReal 
            Height          =   315
            Left            =   6720
            TabIndex        =   165
            Top             =   1455
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeMaquinaReal 
            Height          =   315
            Left            =   3120
            TabIndex        =   166
            Top             =   1455
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaquinasReal 
            Height          =   2145
            Left            =   90
            TabIndex        =   66
            Top             =   225
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   3784
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Index           =   3
         Left            =   75
         TabIndex        =   167
         Top             =   2250
         Visible         =   0   'False
         Width           =   9195
         Begin VB.TextBox TaxaProducao 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   7530
            MaxLength       =   15
            TabIndex        =   178
            Top             =   1425
            Width           =   1185
         End
         Begin VB.TextBox NomeMaquina 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2565
            TabIndex        =   177
            Top             =   1425
            Width           =   1950
         End
         Begin VB.TextBox OPCodigoMRP 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   750
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   176
            Top             =   600
            Width           =   1170
         End
         Begin VB.CheckBox MRP 
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
            Height          =   225
            Left            =   75
            TabIndex        =   175
            Top             =   195
            Width           =   225
         End
         Begin VB.Frame FrameDatas 
            Caption         =   "Datas"
            Height          =   1155
            Left            =   75
            TabIndex        =   170
            Top             =   1050
            Width           =   1995
            Begin VB.Label LabelDataFinal 
               Caption         =   "Final:"
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
               TabIndex        =   174
               Top             =   720
               Width           =   525
            End
            Begin VB.Label DataFinal 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   675
               TabIndex        =   173
               Top             =   690
               Width           =   1155
            End
            Begin VB.Label DataInicio 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   675
               TabIndex        =   172
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label LabelDataInicio 
               Caption         =   "Início:"
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
               Left            =   105
               TabIndex        =   171
               Top             =   270
               Width           =   525
            End
         End
         Begin VB.CommandButton BotaoAbrirOP 
            Caption         =   "Ordem de Produção"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   60
            TabIndex        =   61
            ToolTipText     =   "Abre a tela de Ordem de Produção para a Sub O.P."
            Top             =   2340
            Width           =   1590
         End
         Begin VB.CommandButton BotaoMaquinaPMP 
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
            Height          =   480
            Left            =   1965
            TabIndex        =   62
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   2340
            Width           =   1305
         End
         Begin VB.CommandButton BotaoCompetencia 
            Caption         =   "Competências"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   5955
            TabIndex        =   64
            ToolTipText     =   "Abre a Tela de Competências"
            Top             =   2340
            Width           =   1500
         End
         Begin VB.CommandButton BotaoCT 
            Caption         =   "Centro de Trabalho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3660
            TabIndex        =   63
            ToolTipText     =   "Abre a tela de Centro de Trabalho"
            Top             =   2340
            Width           =   1950
         End
         Begin VB.CommandButton BotaoGrafico 
            Caption         =   "Cronograma Gráfico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   7845
            TabIndex        =   65
            ToolTipText     =   "Abre a tela do Cronograma Gráfico das Etapas de Produção do Item da OP"
            Top             =   2340
            Width           =   1245
         End
         Begin MSMask.MaskEdBox HorasPO 
            Height          =   315
            Left            =   5475
            TabIndex        =   168
            Top             =   1410
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataPO 
            Height          =   315
            Left            =   6315
            TabIndex        =   169
            Top             =   1425
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeMaquina 
            Height          =   315
            Left            =   4515
            TabIndex        =   179
            Top             =   1425
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaquinas 
            Height          =   1785
            Left            =   2160
            TabIndex        =   60
            Top             =   435
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   3149
            _Version        =   393216
         End
         Begin VB.Label LabelMaquinas 
            Caption         =   "Uso de Máquinas Previsto:"
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
            Left            =   2145
            TabIndex        =   181
            Top             =   210
            Width           =   2535
         End
         Begin VB.Label LabelOP 
            Caption         =   "O.P.:"
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
            Left            =   255
            TabIndex        =   180
            Top             =   630
            Width           =   420
         End
         Begin VB.Label LabelMRP 
            Caption         =   "P.M.P. Gerado"
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
            Left            =   360
            TabIndex        =   59
            Top             =   210
            Width           =   1305
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Insumos por Operações"
         Height          =   2895
         Index           =   2
         Left            =   75
         TabIndex        =   182
         Top             =   2250
         Visible         =   0   'False
         Width           =   9195
         Begin VB.ComboBox VersaoKitComp 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OrdemProducaoCro.ctx":01E8
            Left            =   5475
            List            =   "OrdemProducaoCro.ctx":01EA
            Style           =   2  'Dropdown List
            TabIndex        =   187
            Top             =   1035
            Width           =   795
         End
         Begin VB.ComboBox ComposicaoInsumos 
            Height          =   315
            ItemData        =   "OrdemProducaoCro.ctx":01EC
            Left            =   6975
            List            =   "OrdemProducaoCro.ctx":01F6
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   1020
            Width           =   990
         End
         Begin VB.ComboBox UMProduto 
            Height          =   315
            Left            =   4830
            TabIndex        =   185
            Top             =   1035
            Width           =   660
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1890
            TabIndex        =   184
            Top             =   1035
            Width           =   1770
         End
         Begin VB.TextBox OrigemProduto 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3195
            TabIndex        =   183
            Top             =   690
            Width           =   375
         End
         Begin MSMask.MaskEdBox ProdutoInsumos 
            Height          =   315
            Left            =   435
            TabIndex        =   188
            Top             =   1035
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PerdaInsumos 
            Height          =   315
            Left            =   6285
            TabIndex        =   189
            Top             =   1035
            Width           =   690
            _ExtentX        =   1217
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
            Left            =   7965
            TabIndex        =   190
            Top             =   1020
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
         Begin MSMask.MaskEdBox QuantidadeProduto 
            Height          =   315
            Left            =   3660
            TabIndex        =   191
            Top             =   1035
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridOperacaoInsumos 
            Height          =   2430
            Left            =   90
            TabIndex        =   58
            Top             =   255
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   4286
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operação"
         Height          =   2925
         Index           =   1
         Left            =   75
         TabIndex        =   132
         Top             =   2220
         Width           =   9195
         Begin VB.TextBox Observacao 
            Height          =   315
            Left            =   1365
            MaxLength       =   255
            TabIndex        =   51
            Top             =   1680
            Width           =   6285
         End
         Begin VB.Frame FrameConsideraAlgoritmo 
            Caption         =   "Algoritmo"
            Height          =   480
            Left            =   5175
            TabIndex        =   133
            Top             =   2355
            Width           =   2475
            Begin VB.OptionButton OptionMRP 
               Caption         =   "MRP"
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
               Left            =   375
               TabIndex        =   54
               Tag             =   "Não controla a alocação de máquinas."
               Top             =   210
               Width           =   870
            End
            Begin VB.OptionButton OptionMRPII 
               Caption         =   "MRPII"
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
               Left            =   1485
               TabIndex        =   55
               Tag             =   "Controla a alocação de máquinas."
               Top             =   210
               Width           =   870
            End
         End
         Begin VB.CommandButton BotaoRoteiros 
            Caption         =   "Roteiro de Fabricação"
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
            Left            =   7710
            TabIndex        =   56
            ToolTipText     =   "Abre a Tela de Roteiro de Fabricação"
            Top             =   1650
            Width           =   1305
         End
         Begin VB.CommandButton BotaoKit 
            Caption         =   "Kit"
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
            Left            =   7710
            TabIndex        =   57
            ToolTipText     =   "Abre a Tela de Kit"
            Top             =   2310
            Width           =   1305
         End
         Begin MSMask.MaskEdBox CodigoCTPadrao 
            Height          =   315
            Left            =   1380
            TabIndex        =   50
            Top             =   1320
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
            TabIndex        =   49
            Top             =   945
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumMaxMaqPorOper 
            Height          =   315
            Left            =   3945
            TabIndex        =   52
            Top             =   2055
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
            Left            =   7200
            TabIndex        =   53
            Top             =   2055
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label DescricaoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   154
            Top             =   1320
            Width           =   5145
         End
         Begin VB.Label DescricaoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   153
            Top             =   945
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
            TabIndex        =   152
            Top             =   2550
            Width           =   540
         End
         Begin VB.Label Sequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3030
            TabIndex        =   151
            Top             =   2520
            Width           =   420
         End
         Begin VB.Label Nivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   150
            Top             =   2520
            Width           =   420
         End
         Begin VB.Label LabelSeq 
            AutoSize        =   -1  'True
            Caption         =   "Seqüencial:"
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
            Left            =   1935
            TabIndex        =   149
            Top             =   2565
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
            TabIndex        =   148
            Top             =   1740
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
            TabIndex        =   147
            Top             =   1350
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
            TabIndex        =   146
            Top             =   975
            Width           =   1155
         End
         Begin VB.Label LabelCodigoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3480
            TabIndex        =   145
            Top             =   2520
            Width           =   420
         End
         Begin VB.Label LabelCodigoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3930
            TabIndex        =   144
            Top             =   2520
            Width           =   420
         End
         Begin VB.Label VersaoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   143
            Top             =   570
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
            TabIndex        =   142
            Top             =   630
            Width           =   690
         End
         Begin VB.Label ProdutoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   141
            Top             =   195
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
            TabIndex        =   140
            Top             =   225
            Width           =   810
         End
         Begin VB.Label QtdeLabel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6015
            TabIndex        =   139
            Top             =   570
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
            TabIndex        =   138
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label UMLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8235
            TabIndex        =   137
            Top             =   570
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
            TabIndex        =   136
            Top             =   630
            Width           =   480
         End
         Begin VB.Label LabelNumMaxMaqPorOper 
            AutoSize        =   -1  'True
            Caption         =   "Número Máximo de Máquinas:"
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
            Left            =   1380
            TabIndex        =   135
            Top             =   2085
            Width           =   2550
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
            Left            =   5175
            TabIndex        =   134
            Top             =   2085
            Width           =   2010
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
         Height          =   495
         Left            =   8040
         Picture         =   "OrdemProducaoCro.ctx":020A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Altera a Operação da Árvore do Roteiro"
         Top             =   750
         Width           =   1275
      End
      Begin VB.Frame FrameRoteiro 
         Caption         =   "Roteiro de Fabricação:"
         Height          =   1890
         Left            =   30
         TabIndex        =   193
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
      Begin VB.Frame FrameItemOP 
         Caption         =   "Item da OP"
         Height          =   660
         Left            =   8040
         TabIndex        =   192
         Top             =   -15
         Width           =   1275
         Begin MSComCtl2.UpDown UpDownItemOP 
            Height          =   315
            Left            =   780
            TabIndex        =   46
            Top             =   255
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ItemOP 
            Height          =   315
            Left            =   270
            TabIndex        =   45
            Top             =   255
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
         Height          =   480
         Left            =   8055
         TabIndex        =   48
         ToolTipText     =   "Abre o Relatório de Operações"
         Top             =   1380
         Width           =   1245
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   3300
         Left            =   30
         TabIndex        =   194
         Top             =   1905
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   5821
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhe"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Insumos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsão"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Apontamento de Máquinas"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Apontamento de Mão de Obra"
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
      Left            =   135
      TabIndex        =   195
      Top             =   840
      Visible         =   0   'False
      Width           =   9360
      Begin MSMask.MaskEdBox Maquina2 
         Height          =   255
         Left            =   6465
         TabIndex        =   292
         Top             =   3480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoMaquinas 
         Caption         =   "Maquinas"
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
         Left            =   8025
         TabIndex        =   42
         Top             =   4725
         Width           =   1215
      End
      Begin VB.ComboBox Versao 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "OrdemProducaoCro.ctx":1B30
         Left            =   6600
         List            =   "OrdemProducaoCro.ctx":1B32
         Style           =   2  'Dropdown List
         TabIndex        =   205
         Top             =   2340
         Width           =   1875
      End
      Begin VB.ComboBox Destinacao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OrdemProducaoCro.ctx":1B34
         Left            =   6825
         List            =   "OrdemProducaoCro.ctx":1B36
         Style           =   2  'Dropdown List
         TabIndex        =   204
         Top             =   1125
         Width           =   1830
      End
      Begin VB.CheckBox Benef 
         Height          =   210
         Left            =   4590
         TabIndex        =   203
         Top             =   2730
         Width           =   870
      End
      Begin VB.TextBox FilialCliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   240
         Left            =   4170
         TabIndex        =   202
         Text            =   "Filial do Cliente"
         Top             =   2400
         Width           =   1110
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3225
         TabIndex        =   201
         Text            =   "Cliente"
         Top             =   2100
         Width           =   1260
      End
      Begin VB.CommandButton BotaoPedidoDeVenda 
         Caption         =   "Pedido de Venda"
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
         Left            =   6189
         TabIndex        =   41
         Top             =   4725
         Width           =   1680
      End
      Begin VB.TextBox UnidadeMed 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   200
         Top             =   435
         Width           =   600
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   5850
         MaxLength       =   50
         TabIndex        =   199
         Top             =   1950
         Width           =   2600
      End
      Begin VB.ComboBox Situacao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OrdemProducaoCro.ctx":1B38
         Left            =   4920
         List            =   "OrdemProducaoCro.ctx":1B3A
         Style           =   2  'Dropdown List
         TabIndex        =   198
         Top             =   1140
         Width           =   1830
      End
      Begin VB.ComboBox ComboFilialPedido 
         Height          =   315
         ItemData        =   "OrdemProducaoCro.ctx":1B3C
         Left            =   2175
         List            =   "OrdemProducaoCro.ctx":1B3E
         Style           =   2  'Dropdown List
         TabIndex        =   197
         Top             =   1590
         Width           =   1875
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
         Left            =   4353
         TabIndex        =   40
         Top             =   4725
         Width           =   1680
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
         Left            =   105
         TabIndex        =   37
         Top             =   4725
         Width           =   1680
      End
      Begin VB.CommandButton BotaoEstoque 
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
         Height          =   360
         Left            =   1941
         TabIndex        =   38
         Top             =   4725
         Width           =   1125
      End
      Begin VB.CommandButton BotaoGrade 
         Caption         =   "Grade"
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
         Left            =   3222
         TabIndex        =   39
         Top             =   4725
         Width           =   975
      End
      Begin VB.CommandButton BotaoImprimirPrevia 
         Caption         =   "Prévia de req. compras e OP subprod"
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
         TabIndex        =   35
         Top             =   4260
         Width           =   3480
      End
      Begin VB.CommandButton BotaoImprimirRotulos 
         Caption         =   "Imprimir Rótulos de Produção"
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
         Left            =   3750
         TabIndex        =   36
         Top             =   4290
         Width           =   2880
      End
      Begin VB.CheckBox ProduzLogo 
         Height          =   210
         Left            =   4755
         TabIndex        =   196
         Top             =   3405
         Width           =   1125
      End
      Begin MSMask.MaskEdBox Maquina 
         Height          =   255
         Left            =   6870
         TabIndex        =   206
         Top             =   2700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PedidoDeVendaId 
         Height          =   255
         Left            =   4920
         TabIndex        =   207
         Top             =   1590
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Prioridade 
         Height          =   255
         Left            =   6645
         TabIndex        =   208
         Top             =   1590
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
         Left            =   2040
         TabIndex        =   209
         Top             =   480
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
      Begin MSMask.MaskEdBox Produto 
         Height          =   255
         Left            =   330
         TabIndex        =   210
         Top             =   1650
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   270
         Left            =   3570
         TabIndex        =   211
         Top             =   615
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   270
         Left            =   5310
         TabIndex        =   212
         Top             =   720
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
         Left            =   3510
         TabIndex        =   213
         Top             =   1170
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
         Left            =   7440
         TabIndex        =   214
         Top             =   705
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
         Height          =   3720
         Left            =   90
         TabIndex        =   34
         Top             =   420
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   6562
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label QuantDisponivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8355
         TabIndex        =   217
         Top             =   4290
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Qtde Disponível:"
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
         Left            =   6810
         TabIndex        =   216
         Top             =   4350
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Material a ser Produzido"
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
         Left            =   105
         TabIndex        =   215
         Top             =   120
         Width           =   2100
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6900
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   218
      TabStop         =   0   'False
      Top             =   135
      Width           =   2595
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "OrdemProducaoCro.ctx":1B40
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprimir OP"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   555
         Picture         =   "OrdemProducaoCro.ctx":1C42
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1065
         Picture         =   "OrdemProducaoCro.ctx":1D9C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1560
         Picture         =   "OrdemProducaoCro.ctx":1F26
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2040
         Picture         =   "OrdemProducaoCro.ctx":2458
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5625
      Left            =   105
      TabIndex        =   219
      Top             =   435
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9922
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
            Caption         =   "Ficha de Processo 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ficha 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "OrdemProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTOrdemProducao
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    
    Set objCT = New CTOrdemProducao
    Set objCT.objUserControl = Me

    Set objCT.gobjInfoUsu = New CTOrdemProdVGCro
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTOrdemProdCro

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

Private Sub AlmoxPadraoLabel_Click()
     Call objCT.AlmoxPadraoLabel_Click
End Sub

Private Sub BotaoAbrirOP_Click()
     Call objCT.BotaoAbrirOP_Click
End Sub

Private Sub BotaoCompetencia_Click()
     Call objCT.BotaoCompetencia_Click
End Sub

Private Sub BotaoCT_Click()
     Call objCT.BotaoCT_Click
End Sub

Private Sub BotaoGrade_Click()
     Call objCT.BotaoGrade_Click
End Sub

Private Sub BotaoGrafico_Click()
     Call objCT.BotaoGrafico_Click
End Sub

Private Sub BotaoMaquinaPMP_Click()
     Call objCT.BotaoMaquinaPMP_Click
End Sub

Private Sub BotaoMaquinas_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoMaquinas_Click(objCT)
End Sub

Private Sub BotaoImprimir_Click()
     Call objCT.BotaoImprimir_Click
End Sub

Private Sub BotaoPedidoDeVenda_Click()
     Call objCT.BotaoPedidoDeVenda_Click
End Sub

Private Sub Benef_Click()
     Call objCT.Benef_Click
End Sub

Private Sub Benef_GotFocus()
     Call objCT.Benef_GotFocus
End Sub

Private Sub Benef_KeyPress(KeyAscii As Integer)
     Call objCT.Benef_KeyPress(KeyAscii)
End Sub

Private Sub Benef_Validate(Cancel As Boolean)
     Call objCT.Benef_Validate(Cancel)
End Sub

Private Sub Ccl_Change()
     Call objCT.Ccl_Change
End Sub

Private Sub Ccl_GotFocus()
     Call objCT.Ccl_GotFocus
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)
     Call objCT.Ccl_KeyPress(KeyAscii)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub CclPadraoLabel_Click()
     Call objCT.CclPadraoLabel_Click
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub CodOPGerada_Change()
     Call objCT.CodOPGerada_Change
End Sub

Private Sub ComboFilialPedido_GotFocus()
     Call objCT.ComboFilialPedido_GotFocus
End Sub

Private Sub ComboFilialPedido_KeyPress(KeyAscii As Integer)
     Call objCT.ComboFilialPedido_KeyPress(KeyAscii)
End Sub

Private Sub Data_Change()
     Call objCT.Data_Change
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub DataFimPadrao_GotFocus()
     Call objCT.DataFimPadrao_GotFocus
End Sub

Private Sub DataInicioPadrao_GotFocus()
     Call objCT.DataInicioPadrao_GotFocus
End Sub

Private Sub DataPrevisaoFim_Change()
     Call objCT.DataPrevisaoFim_Change
End Sub

Private Sub DataPrevisaoFim_GotFocus()
     Call objCT.DataPrevisaoFim_GotFocus
End Sub

Private Sub DataPrevisaoFim_KeyPress(KeyAscii As Integer)
     Call objCT.DataPrevisaoFim_KeyPress(KeyAscii)
End Sub

Private Sub DataPrevisaoFim_Validate(Cancel As Boolean)
     Call objCT.DataPrevisaoFim_Validate(Cancel)
End Sub

Private Sub DataPrevisaoInicio_Change()
     Call objCT.DataPrevisaoInicio_Change
End Sub

Private Sub DataPrevisaoInicio_GotFocus()
     Call objCT.DataPrevisaoInicio_GotFocus
End Sub

Private Sub DataPrevisaoInicio_KeyPress(KeyAscii As Integer)
     Call objCT.DataPrevisaoInicio_KeyPress(KeyAscii)
End Sub

Private Sub DataPrevisaoInicio_Validate(Cancel As Boolean)
     Call objCT.DataPrevisaoInicio_Validate(Cancel)
End Sub

Private Sub Destinacao_Click()
     Call objCT.Destinacao_Click
End Sub

Private Sub Destinacao_GotFocus()
     Call objCT.Destinacao_GotFocus
End Sub

Private Sub Destinacao_KeyPress(KeyAscii As Integer)
     Call objCT.Destinacao_KeyPress(KeyAscii)
End Sub

Private Sub Destinacao_Validate(Cancel As Boolean)
     Call objCT.Destinacao_Validate(Cancel)
End Sub

Private Sub GeraOP_Click()
     Call objCT.GeraOP_Click
End Sub

Private Sub GeraReqCompra_Click()
     Call objCT.GeraReqCompra_Click
End Sub

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMovimentos_KeyDown(KeyCode, Shift)
End Sub

Private Sub ItemOP_Validate(Cancel As Boolean)
     Call objCT.ItemOP_Validate(Cancel)
End Sub

Private Sub LabelPrestador_Click()
     Call objCT.LabelPrestador_Click
End Sub

Private Sub Maquina_Change()
     Call objCT.Maquina_Change
End Sub

Private Sub BotaoCcls_Click()
     Call objCT.BotaoCcls_Click
End Sub

Private Sub BotaoEstoque_Click()
     Call objCT.BotaoEstoque_Click
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

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub CodigoOPLabel_Click()
     Call objCT.CodigoOPLabel_Click
End Sub

Private Sub ComboFilialPedido_Change()
     Call objCT.ComboFilialPedido_Change
End Sub

Private Sub ComboFilialPedido_Validate(Cancel As Boolean)
     Call objCT.ComboFilialPedido_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub GridMovimentos_Click()
     Call objCT.GridMovimentos_Click
End Sub

Private Sub GridMovimentos_GotFocus()
     Call objCT.GridMovimentos_GotFocus
End Sub

Private Sub GridMovimentos_EnterCell()
     Call objCT.GridMovimentos_EnterCell
End Sub

Private Sub GridMovimentos_LeaveCell()
     Call objCT.GridMovimentos_LeaveCell
End Sub

Private Sub GridMovimentos_KeyPress(KeyAscii As Integer)
     Call objCT.GridMovimentos_KeyPress(KeyAscii)
End Sub

Private Sub GridMovimentos_RowColChange()
     Call objCT.GridMovimentos_RowColChange
End Sub

Private Sub GridMovimentos_Scroll()
     Call objCT.GridMovimentos_Scroll
End Sub

Function Trata_Parametros(Optional objOrdemDeProducao As ClassOrdemDeProducao) As Long
     Trata_Parametros = objCT.Trata_Parametros(objOrdemDeProducao)
End Function

Private Sub PedidoDeVendaId_Change()
     Call objCT.PedidoDeVendaId_Change
End Sub

Private Sub PedidoDeVendaId_GotFocus()
     Call objCT.PedidoDeVendaId_GotFocus
End Sub

Private Sub PedidoDeVendaId_KeyPress(KeyAscii As Integer)
     Call objCT.PedidoDeVendaId_KeyPress(KeyAscii)
End Sub

Private Sub PedidoDeVendaId_Validate(Cancel As Boolean)
     Call objCT.PedidoDeVendaId_Validate(Cancel)
End Sub

Private Sub PrestadorServico_Change()
     Call objCT.PrestadorServico_Change
End Sub

Private Sub PrestadorServico_Validate(Cancel As Boolean)
     Call objCT.PrestadorServico_Validate(Cancel)
End Sub

Private Sub Prioridade_Change()
     Call objCT.Prioridade_Change
End Sub

Private Sub Prioridade_GotFocus()
     Call objCT.Prioridade_GotFocus
End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)
     Call objCT.Prioridade_KeyPress(KeyAscii)
End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)
     Call objCT.Prioridade_Validate(Cancel)
End Sub

Private Sub PrioridadePadrao_GotFocus()
     Call objCT.PrioridadePadrao_GotFocus
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

Private Sub Situacao_Click()
     Call objCT.Situacao_Click
End Sub

Private Sub Situacao_GotFocus()
     Call objCT.Situacao_GotFocus
End Sub

Private Sub Situacao_KeyPress(KeyAscii As Integer)
     Call objCT.Situacao_KeyPress(KeyAscii)
End Sub

Private Sub Situacao_Validate(Cancel As Boolean)
     Call objCT.Situacao_Validate(Cancel)
End Sub

Private Sub UpDownData_Change()
     Call objCT.UpDownData_Change
End Sub

Private Sub UpDownData_DownClick()
     Call objCT.UpDownData_DownClick
End Sub

Private Sub UpDownData_UpClick()
     Call objCT.UpDownData_UpClick
End Sub

Private Sub UpDownInicio_Change()
     Call objCT.UpDownInicio_Change
End Sub

Private Sub UpDownInicio_DownClick()
     Call objCT.UpDownInicio_DownClick
End Sub

Private Sub UpDownInicio_UpClick()
     Call objCT.UpDownInicio_UpClick
End Sub

Private Sub UpDownFim_Change()
     Call objCT.UpDownFim_Change
End Sub

Private Sub UpDownFim_DownClick()
     Call objCT.UpDownFim_DownClick
End Sub

Private Sub UpDownFim_UpClick()
     Call objCT.UpDownFim_UpClick
End Sub

Private Sub UpDownItemOP_DownClick()
     Call objCT.UpDownItemOP_DownClick
End Sub

Private Sub UpDownItemOP_UpClick()
     Call objCT.UpDownItemOP_UpClick
End Sub

Private Sub PrioridadePadraoLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PrioridadePadraoLbl, Source, X, Y)
End Sub
Private Sub PrioridadePadraoLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PrioridadePadraoLbl, Button, Shift, X, Y)
End Sub
Private Sub DestPadraoLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DestPadraoLbl, Source, X, Y)
End Sub
Private Sub DestPadraoLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DestPadraoLbl, Button, Shift, X, Y)
End Sub
Private Sub DataPrevFimLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPrevFimLbl, Source, X, Y)
End Sub
Private Sub DataPrevFimLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPrevFimLbl, Button, Shift, X, Y)
End Sub
Private Sub DataPrevIniLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPrevIniLbl, Source, X, Y)
End Sub
Private Sub DataPrevIniLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPrevIniLbl, Button, Shift, X, Y)
End Sub
Private Sub AlmoxPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxPadraoLabel, Source, X, Y)
End Sub
Private Sub AlmoxPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxPadraoLabel, Button, Shift, X, Y)
End Sub
Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub
Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub
Private Sub QuantDisponivel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivel, Source, X, Y)
End Sub
Private Sub QuantDisponivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivel, Button, Shift, X, Y)
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub CodigoOPLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoOPLabel, Source, X, Y)
End Sub
Private Sub CodigoOPLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoOPLabel, Button, Shift, X, Y)
End Sub
Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub AlmoxPadrao_Validate(Cancel As Boolean)
     Call objCT.AlmoxPadrao_Validate(Cancel)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
     Call objCT.Data_Validate(Cancel)
End Sub

Private Sub DataFimPadrao_Validate(Cancel As Boolean)
     Call objCT.DataFimPadrao_Validate(Cancel)
End Sub

Private Sub DataInicioPadrao_Validate(Cancel As Boolean)
     Call objCT.DataInicioPadrao_Validate(Cancel)
End Sub

Private Sub CclPadrao_Validate(Cancel As Boolean)
     Call objCT.CclPadrao_Validate(Cancel)
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub GridMovimentos_Validate(Cancel As Boolean)
     Call objCT.GridMovimentos_Validate(Cancel)
End Sub

Private Sub LabelCodOPGerada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodOPGerada, Source, X, Y)
End Sub
Private Sub LabelCodOPGerada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodOPGerada, Button, Shift, X, Y)
End Sub
Private Sub Maquina_GotFocus()
     Call objCT.Maquina_GotFocus
End Sub

Private Sub Maquina_KeyPress(KeyAscii As Integer)
     Call objCT.Maquina_KeyPress(KeyAscii)
End Sub

Private Sub Maquina_Validate(Cancel As Boolean)
     Call objCT.Maquina_Validate(Cancel)
End Sub

Private Sub Versao_Change()
     Call objCT.Versao_Change
End Sub

Private Sub Versao_GotFocus()
     Call objCT.Versao_GotFocus
End Sub

Private Sub Versao_KeyPress(KeyAscii As Integer)
     Call objCT.Versao_KeyPress(KeyAscii)
End Sub

Private Sub Versao_Validate(Cancel As Boolean)
     Call objCT.Versao_Validate(Cancel)
End Sub

Private Sub BotaoImprimirPrevia_Click()
     Call objCT.BotaoImprimirPrevia_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoAlterar_Click()
     Call objCT.BotaoAlterar_Click
End Sub

Private Sub CodigoCompetencia_Change()
     Call objCT.CodigoCompetencia_Change
End Sub

Private Sub CodigoCompetencia_GotFocus()
     Call objCT.CodigoCompetencia_GotFocus
End Sub

Private Sub CodigoCompetencia_Validate(Cancel As Boolean)
     Call objCT.CodigoCompetencia_Validate(Cancel)
End Sub

Private Sub CodigoCTPadrao_Change()
     Call objCT.CodigoCTPadrao_Change
End Sub

Private Sub CodigoCTPadrao_GotFocus()
     Call objCT.CodigoCTPadrao_GotFocus
End Sub

Private Sub CodigoCTPadrao_Validate(Cancel As Boolean)
     Call objCT.CodigoCTPadrao_Validate(Cancel)
End Sub

Private Sub CompetenciaLabel_Click()
     Call objCT.CompetenciaLabel_Click
End Sub

Private Sub CTLabel_Click()
     Call objCT.CTLabel_Click
End Sub

Private Sub MRP_Click()
     Call objCT.MRP_Click
End Sub

Private Sub Roteiro_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.Roteiro_NodeClick(Node)
End Sub

Private Sub TabStrip2_Click()
     Call objCT.TabStrip2_Click
End Sub

Private Sub GridOperacaoInsumos_Click()
     Call objCT.GridOperacaoInsumos_Click
End Sub

Private Sub GridOperacaoInsumos_GotFocus()
     Call objCT.GridOperacaoInsumos_GotFocus
End Sub

Private Sub GridOperacaoInsumos_EnterCell()
     Call objCT.GridOperacaoInsumos_EnterCell
End Sub

Private Sub GridOperacaoInsumos_LeaveCell()
     Call objCT.GridOperacaoInsumos_LeaveCell
End Sub

Private Sub GridOperacaoInsumos_KeyPress(KeyAscii As Integer)
     Call objCT.GridOperacaoInsumos_KeyPress(KeyAscii)
End Sub

Private Sub GridOperacaoInsumos_Scroll()
     Call objCT.GridOperacaoInsumos_Scroll
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

Private Sub QuantidadeProduto_Change()
     Call objCT.QuantidadeProduto_Change
End Sub

Private Sub QuantidadeProduto_GotFocus()
     Call objCT.QuantidadeProduto_GotFocus
End Sub

Private Sub QuantidadeProduto_KeyPress(KeyAscii As Integer)
     Call objCT.QuantidadeProduto_KeyPress(KeyAscii)
End Sub

Private Sub QuantidadeProduto_Validate(Cancel As Boolean)
     Call objCT.QuantidadeProduto_Validate(Cancel)
End Sub

Private Sub UMProduto_Change()
     Call objCT.UMProduto_Change
End Sub

Private Sub UMProduto_GotFocus()
     Call objCT.UMProduto_GotFocus
End Sub

Private Sub UMProduto_KeyPress(KeyAscii As Integer)
     Call objCT.UMProduto_KeyPress(KeyAscii)
End Sub

Private Sub UMProduto_Validate(Cancel As Boolean)
     Call objCT.UMProduto_Validate(Cancel)
End Sub

Private Sub VersaoKitComp_Change()
     Call objCT.VersaoKitComp_Change
End Sub

Private Sub VersaoKitComp_GotFocus()
     Call objCT.VersaoKitComp_GotFocus
End Sub

Private Sub VersaoKitComp_KeyPress(KeyAscii As Integer)
     Call objCT.VersaoKitComp_KeyPress(KeyAscii)
End Sub

Private Sub VersaoKitComp_Validate(Cancel As Boolean)
     Call objCT.VersaoKitComp_Validate(Cancel)
End Sub

Private Sub PerdaInsumos_Change()
     Call objCT.PerdaInsumos_Change
End Sub

Private Sub PerdaInsumos_GotFocus()
     Call objCT.PerdaInsumos_GotFocus
End Sub

Private Sub PerdaInsumos_KeyPress(KeyAscii As Integer)
     Call objCT.PerdaInsumos_KeyPress(KeyAscii)
End Sub

Private Sub PerdaInsumos_Validate(Cancel As Boolean)
     Call objCT.PerdaInsumos_Validate(Cancel)
End Sub

Private Sub ComposicaoInsumos_Change()
     Call objCT.ComposicaoInsumos_Change
End Sub

Private Sub ComposicaoInsumos_GotFocus()
     Call objCT.ComposicaoInsumos_GotFocus
End Sub

Private Sub ComposicaoInsumos_KeyPress(KeyAscii As Integer)
     Call objCT.ComposicaoInsumos_KeyPress(KeyAscii)
End Sub

Private Sub ComposicaoInsumos_Validate(Cancel As Boolean)
     Call objCT.ComposicaoInsumos_Validate(Cancel)
End Sub

Private Sub CustoStandardInsumos_Change()
     Call objCT.CustoStandardInsumos_Change
End Sub

Private Sub CustoStandardInsumos_GotFocus()
     Call objCT.CustoStandardInsumos_GotFocus
End Sub

Private Sub CustoStandardInsumos_KeyPress(KeyAscii As Integer)
     Call objCT.CustoStandardInsumos_KeyPress(KeyAscii)
End Sub

Private Sub CustoStandardInsumos_Validate(Cancel As Boolean)
     Call objCT.CustoStandardInsumos_Validate(Cancel)
End Sub

Private Sub ProdutoInsumos_Change()
     Call objCT.ProdutoInsumos_Change
End Sub

Private Sub ProdutoInsumos_GotFocus()
     Call objCT.ProdutoInsumos_GotFocus
End Sub

Private Sub ProdutoInsumos_KeyPress(KeyAscii As Integer)
     Call objCT.ProdutoInsumos_KeyPress(KeyAscii)
End Sub

Private Sub ProdutoInsumos_Validate(Cancel As Boolean)
     Call objCT.ProdutoInsumos_Validate(Cancel)
End Sub

Private Sub GridMaquinas_Click()
     Call objCT.GridMaquinas_Click
End Sub

Private Sub GridMaquinas_GotFocus()
     Call objCT.GridMaquinas_GotFocus
End Sub

Private Sub GridMaquinas_EnterCell()
     Call objCT.GridMaquinas_EnterCell
End Sub

Private Sub GridMaquinas_LeaveCell()
     Call objCT.GridMaquinas_LeaveCell
End Sub

Private Sub GridMaquinas_KeyPress(KeyAscii As Integer)
     Call objCT.GridMaquinas_KeyPress(KeyAscii)
End Sub

Private Sub GridMaquinas_Scroll()
     Call objCT.GridMaquinas_Scroll
End Sub

Private Sub NomeMaquina_GotFocus()
     Call objCT.NomeMaquina_GotFocus
End Sub

Private Sub NomeMaquina_KeyPress(KeyAscii As Integer)
     Call objCT.NomeMaquina_KeyPress(KeyAscii)
End Sub

Private Sub NomeMaquina_Validate(Cancel As Boolean)
     Call objCT.NomeMaquina_Validate(Cancel)
End Sub

Private Sub QuantidadeMaquina_GotFocus()
     Call objCT.QuantidadeMaquina_GotFocus
End Sub

Private Sub QuantidadeMaquina_KeyPress(KeyAscii As Integer)
     Call objCT.QuantidadeMaquina_KeyPress(KeyAscii)
End Sub

Private Sub QuantidadeMaquina_Validate(Cancel As Boolean)
     Call objCT.QuantidadeMaquina_Validate(Cancel)
End Sub

Private Sub TaxaProducao_GotFocus()
     Call objCT.TaxaProducao_GotFocus
End Sub

Private Sub TaxaProducao_KeyPress(KeyAscii As Integer)
     Call objCT.TaxaProducao_KeyPress(KeyAscii)
End Sub

Private Sub TaxaProducao_Validate(Cancel As Boolean)
     Call objCT.TaxaProducao_Validate(Cancel)
End Sub

Private Sub HorasPO_GotFocus()
     Call objCT.HorasPO_GotFocus
End Sub

Private Sub HorasPO_KeyPress(KeyAscii As Integer)
     Call objCT.HorasPO_KeyPress(KeyAscii)
End Sub

Private Sub HorasPO_Validate(Cancel As Boolean)
     Call objCT.HorasPO_Validate(Cancel)
End Sub

Private Sub DataPO_GotFocus()
     Call objCT.DataPO_GotFocus
End Sub

Private Sub DataPO_KeyPress(KeyAscii As Integer)
     Call objCT.DataPO_KeyPress(KeyAscii)
End Sub

Private Sub DataPO_Validate(Cancel As Boolean)
     Call objCT.DataPO_Validate(Cancel)
End Sub

Private Sub Roteiro_Collapse(ByVal Node As MSComctlLib.Node)
     Call objCT.Roteiro_Collapse(Node)
End Sub

Private Sub BotaoImprimirOper_Click()
     Call objCT.BotaoImprimirOper_Click
End Sub

Private Sub NomeMaquinaReal_change()
     Call objCT.NomeMaquinaReal_change
End Sub

Private Sub NomeMaquinaReal_GotFocus()
     Call objCT.NomeMaquinaReal_GotFocus
End Sub

Private Sub NomeMaquinaReal_KeyPress(KeyAscii As Integer)
     Call objCT.NomeMaquinaReal_KeyPress(KeyAscii)
End Sub

Private Sub NomeMaquinaReal_Validate(Cancel As Boolean)
     Call objCT.NomeMaquinaReal_Validate(Cancel)
End Sub

Private Sub QuantidadeMaquinaReal_change()
     Call objCT.QuantidadeMaquinaReal_change
End Sub

Private Sub QuantidadeMaquinaReal_GotFocus()
     Call objCT.QuantidadeMaquinaReal_GotFocus
End Sub

Private Sub QuantidadeMaquinaReal_KeyPress(KeyAscii As Integer)
     Call objCT.QuantidadeMaquinaReal_KeyPress(KeyAscii)
End Sub

Private Sub QuantidadeMaquinaReal_Validate(Cancel As Boolean)
     Call objCT.QuantidadeMaquinaReal_Validate(Cancel)
End Sub

Private Sub HorasReal_change()
     Call objCT.HorasReal_change
End Sub

Private Sub HorasReal_GotFocus()
     Call objCT.HorasReal_GotFocus
End Sub

Private Sub HorasReal_KeyPress(KeyAscii As Integer)
     Call objCT.HorasReal_KeyPress(KeyAscii)
End Sub

Private Sub HorasReal_Validate(Cancel As Boolean)
     Call objCT.HorasReal_Validate(Cancel)
End Sub

Private Sub DataReal_change()
     Call objCT.DataReal_change
End Sub

Private Sub DataReal_GotFocus()
     Call objCT.DataReal_GotFocus
End Sub

Private Sub DataReal_KeyPress(KeyAscii As Integer)
     Call objCT.DataReal_KeyPress(KeyAscii)
End Sub

Private Sub DataReal_Validate(Cancel As Boolean)
     Call objCT.DataReal_Validate(Cancel)
End Sub

Private Sub NomeMaquinaRealMO_change()
     Call objCT.NomeMaquinaRealMO_change
End Sub

Private Sub NomeMaquinaRealMO_GotFocus()
     Call objCT.NomeMaquinaRealMO_GotFocus
End Sub

Private Sub NomeMaquinaRealMO_KeyPress(KeyAscii As Integer)
     Call objCT.NomeMaquinaRealMO_KeyPress(KeyAscii)
End Sub

Private Sub NomeMaquinaRealMO_Validate(Cancel As Boolean)
     Call objCT.NomeMaquinaRealMO_Validate(Cancel)
End Sub

Private Sub QuantidadeMaquinaRealMO_change()
     Call objCT.QuantidadeMaquinaRealMO_change
End Sub

Private Sub QuantidadeMaquinaRealMO_GotFocus()
     Call objCT.QuantidadeMaquinaRealMO_GotFocus
End Sub

Private Sub QuantidadeMaquinaRealMO_KeyPress(KeyAscii As Integer)
     Call objCT.QuantidadeMaquinaRealMO_KeyPress(KeyAscii)
End Sub

Private Sub QuantidadeMaquinaRealMO_Validate(Cancel As Boolean)
     Call objCT.QuantidadeMaquinaRealMO_Validate(Cancel)
End Sub

Private Sub HorasRealMO_change()
     Call objCT.HorasRealMO_change
End Sub

Private Sub HorasRealMO_GotFocus()
     Call objCT.HorasRealMO_GotFocus
End Sub

Private Sub HorasRealMO_KeyPress(KeyAscii As Integer)
     Call objCT.HorasRealMO_KeyPress(KeyAscii)
End Sub

Private Sub HorasRealMO_Validate(Cancel As Boolean)
     Call objCT.HorasRealMO_Validate(Cancel)
End Sub

Private Sub DataRealMO_change()
     Call objCT.DataRealMO_change
End Sub

Private Sub DataRealMO_GotFocus()
     Call objCT.DataRealMO_GotFocus
End Sub

Private Sub DataRealMO_KeyPress(KeyAscii As Integer)
     Call objCT.DataRealMO_KeyPress(KeyAscii)
End Sub

Private Sub DataRealMO_Validate(Cancel As Boolean)
     Call objCT.DataRealMO_Validate(Cancel)
End Sub

Private Sub DescricaoTipoMO_Change()
     Call objCT.DescricaoTipoMO_Change
End Sub

Private Sub DescricaoTipoMO_GotFocus()
     Call objCT.DescricaoTipoMO_GotFocus
End Sub

Private Sub DescricaoTipoMO_KeyPress(KeyAscii As Integer)
     Call objCT.DescricaoTipoMO_KeyPress(KeyAscii)
End Sub

Private Sub DescricaoTipoMO_Validate(Cancel As Boolean)
     Call objCT.DescricaoTipoMO_Validate(Cancel)
End Sub

Private Sub BotaoMO_Click()
     Call objCT.BotaoMO_Click
End Sub

Private Sub BotaoMaquinasReal_Click()
     Call objCT.BotaoMaquinasReal_Click
End Sub

Private Sub BotaoTrazerMaquinasPrevistas_Click()
     Call objCT.BotaoTrazerMaquinasPrevistas_Click
End Sub

Private Sub GridMaquinasReal_Click()
     Call objCT.GridMaquinasReal_Click
End Sub

Private Sub GridMaquinasReal_GotFocus()
     Call objCT.GridMaquinasReal_GotFocus
End Sub

Private Sub GridMaquinasReal_EnterCell()
     Call objCT.GridMaquinasReal_EnterCell
End Sub

Private Sub GridMaquinasReal_LeaveCell()
     Call objCT.GridMaquinasReal_LeaveCell
End Sub

Private Sub GridMaquinasReal_KeyPress(KeyAscii As Integer)
     Call objCT.GridMaquinasReal_KeyPress(KeyAscii)
End Sub

Private Sub GridMaquinasReal_RowColChange()
     Call objCT.GridMaquinasReal_RowColChange
End Sub

Private Sub GridMaquinasReal_Scroll()
     Call objCT.GridMaquinasReal_Scroll
End Sub

Private Sub GridMaquinasReal_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMaquinasReal_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridMaquinasReal_LostFocus()
     Call objCT.GridMaquinasReal_LostFocus
End Sub

Private Sub GridMOReal_Click()
     Call objCT.GridMOReal_Click
End Sub

Private Sub GridMOReal_GotFocus()
     Call objCT.GridMOReal_GotFocus
End Sub

Private Sub GridMOReal_EnterCell()
     Call objCT.GridMOReal_EnterCell
End Sub

Private Sub GridMOReal_LeaveCell()
     Call objCT.GridMOReal_LeaveCell
End Sub

Private Sub GridMOReal_KeyPress(KeyAscii As Integer)
     Call objCT.GridMOReal_KeyPress(KeyAscii)
End Sub

Private Sub GridMOReal_RowColChange()
     Call objCT.GridMOReal_RowColChange
End Sub

Private Sub GridMOReal_Scroll()
     Call objCT.GridMOReal_Scroll
End Sub

Private Sub GridMOReal_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMOReal_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridMOReal_LostFocus()
     Call objCT.GridMOReal_LostFocus
End Sub

Private Sub BotaoImprimirRotulos_Click()
     Call objCT.BotaoImprimirRotulos_Click
End Sub

Private Sub NumMaxMaqPorOper_Validate(Cancel As Boolean)
     Call objCT.NumMaxMaqPorOper_Validate(Cancel)
End Sub

Private Sub NumMaxMaqPorOper_GotFocus()
     Call objCT.NumMaxMaqPorOper_GotFocus
End Sub

Private Sub NumMaxMaqPorOper_Change()
     Call objCT.NumMaxMaqPorOper_Change
End Sub

Private Sub BotaoRoteiros_Click()
     Call objCT.BotaoRoteiros_Click
End Sub

Private Sub BotaoKit_Click()
     Call objCT.BotaoKit_Click
End Sub

Private Sub Repeticao_Validate(Cancel As Boolean)
     Call objCT.Repeticao_Validate(Cancel)
End Sub

Private Sub Repeticao_GotFocus()
     Call objCT.Repeticao_GotFocus
End Sub

Private Sub Repeticao_Change()
     Call objCT.Repeticao_Change
End Sub

Private Sub BotaoOPRC_Click()
     Call objCT.BotaoOPRC_Click
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
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

'#####CROMATON #######
Private Sub Embalagem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Embalagem, Source, X, Y)
End Sub

Private Sub EmbalagemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(EmbalagemLabel, Source, X, Y)
End Sub

Private Sub EmbalagemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(EmbalagemLabel, Button, Shift, X, Y)
End Sub

Private Sub Mistura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Mistura, Source, X, Y)
End Sub

Private Sub MisturaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MisturaLabel, Source, X, Y)
End Sub

Private Sub MisturaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MisturaLabel, Button, Shift, X, Y)
End Sub

Private Sub TempoMistura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TempoMistura, Source, X, Y)
End Sub

Private Sub TempoMisturaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TempoMisturaLabel, Source, X, Y)
End Sub

Private Sub TempoMisturaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TempoMisturaLabel, Button, Shift, X, Y)
End Sub

Private Sub CargaDrays_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CargaDrays, Source, X, Y)
End Sub

Private Sub CargaDraysLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CargaDraysLabel, Source, X, Y)
End Sub

Private Sub CargaDraysLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CargaDraysLabel, Button, Shift, X, Y)
End Sub

'#########################################

Private Sub Mistura_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Mistura_Change(objCT)
End Sub

Private Sub Mistura_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Mistura_Validate(objCT, Cancel)
End Sub

Private Sub TempoMistura_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TempoMistura_Validate(objCT, Cancel)
End Sub

Private Sub TempoMistura_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.TempoMistura_Change(objCT)
End Sub

Private Sub CargaDrays_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.CargaDrays_Validate(objCT, Cancel)
End Sub

Private Sub CargaDrays_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.CargaDrays_Change(objCT)
End Sub

Private Sub Embalagem_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Embalagem_Validate(objCT, Cancel)
End Sub

Private Sub Embalagem_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Embalagem_Change(objCT)
End Sub

Public Sub GridCarga_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_GotFocus(objCT)
End Sub

Public Sub GridCarga_EnterCell()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_EnterCell(objCT)
End Sub

Public Sub GridCarga_LeaveCell()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_LeaveCell(objCT)
End Sub

Private Sub GridCarga_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_Click(objCT)
End Sub

Public Sub GridCarga_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_KeyDown(objCT, KeyCode, Shift)
End Sub

Public Sub GridCarga_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_KeyPress(objCT, KeyAscii)
End Sub

Public Sub GridCarga_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_Validate(objCT, Cancel)
End Sub

Public Sub GridCarga_RowColChange()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_RowColChange(objCT)
End Sub

Public Sub GridCarga_Scroll()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridCarga_Scroll(objCT)
End Sub

Public Sub GridInsumo_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_GotFocus(objCT)
End Sub

Public Sub GridInsumo_EnterCell()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_EnterCell(objCT)
End Sub

Public Sub GridInsumo_LeaveCell()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_LeaveCell(objCT)
End Sub

Private Sub GridInsumo_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_Click(objCT)
End Sub

Public Sub GridInsumo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_KeyDown(objCT, KeyCode, Shift)
End Sub

Public Sub GridInsumo_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_KeyPress(objCT, KeyAscii)
End Sub

Public Sub GridInsumo_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_Validate(objCT, Cancel)
End Sub

Public Sub GridInsumo_RowColChange()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_RowColChange(objCT)
End Sub

Public Sub GridInsumo_Scroll()
    Call objCT.gobjInfoUsu.gobjTelaUsu.GridInsumo_Scroll(objCT)
End Sub

Private Sub ProdutoBase_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.ProdutoBase_Change(objCT)
End Sub

Private Sub ProdutoBase_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.ProdutoBase_GotFocus(objCT)
End Sub

Private Sub ProdutoBase_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.ProdutoBase_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ProdutoBase_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.ProdutoBase_Validate(objCT, Cancel)
End Sub

Private Sub KgProdBase_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgProdBase_Change(objCT)
End Sub

Private Sub KgProdBase_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgProdBase_GotFocus(objCT)
End Sub

Private Sub KgProdBase_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgProdBase_KeyPress(objCT, KeyAscii)
End Sub

Private Sub KgProdBase_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgProdBase_Validate(objCT, Cancel)
End Sub

Private Sub KgCarga_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCarga_Change(objCT)
End Sub

Private Sub KgCarga_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCarga_GotFocus(objCT)
End Sub

Private Sub KgCarga_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCarga_KeyPress(objCT, KeyAscii)
End Sub

Private Sub KgCarga_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCarga_Validate(objCT, Cancel)
End Sub

Private Sub QtdCarga_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.QtdCarga_Change(objCT)
End Sub

Private Sub QtdCarga_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.QtdCarga_GotFocus(objCT)
End Sub

Private Sub QtdCarga_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.QtdCarga_KeyPress(objCT, KeyAscii)
End Sub

Private Sub QtdCarga_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.QtdCarga_Validate(objCT, Cancel)
End Sub

Private Sub Item_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Item_Change(objCT)
End Sub

Private Sub Item_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Item_GotFocus(objCT)
End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Item_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Item_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Item_Validate(objCT, Cancel)
End Sub

Private Sub CodInsumo_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.CodInsumo_Change(objCT)
End Sub

Private Sub CodInsumo_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.CodInsumo_GotFocus(objCT)
End Sub

Private Sub CodInsumo_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.CodInsumo_KeyPress(objCT, KeyAscii)
End Sub

Private Sub CodInsumo_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.CodInsumo_Validate(objCT, Cancel)
End Sub

Private Sub DescInsumo_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.DescInsumo_Change(objCT)
End Sub

Private Sub DescInsumo_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.DescInsumo_GotFocus(objCT)
End Sub

Private Sub DescInsumo_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.DescInsumo_KeyPress(objCT, KeyAscii)
End Sub

Private Sub DescInsumo_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.DescInsumo_Validate(objCT, Cancel)
End Sub

Private Sub KgCargaInsumo_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCargaInsumo_Change(objCT)
End Sub

Private Sub KgCargaInsumo_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCargaInsumo_GotFocus(objCT)
End Sub

Private Sub KgCargaInsumo_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCargaInsumo_KeyPress(objCT, KeyAscii)
End Sub

Private Sub KgCargaInsumo_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.KgCargaInsumo_Validate(objCT, Cancel)
End Sub

Private Sub Fornecedor_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Fornecedor_Change(objCT)
End Sub

Private Sub Fornecedor_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Fornecedor_GotFocus(objCT)
End Sub

Private Sub Fornecedor_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Fornecedor_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Fornecedor_Validate(objCT, Cancel)
End Sub

Private Sub Lote_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Lote_Change(objCT)
End Sub

Private Sub Lote_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Lote_GotFocus(objCT)
End Sub

Private Sub Lote_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Lote_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Lote_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Lote_Validate(objCT, Cancel)
End Sub

Private Sub CargaDrays_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.CargaDrays_Click(objCT)
End Sub

Private Sub Maquina2_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Maquina2_Change(objCT)
End Sub

Private Sub Embalagem_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Embalagem_Click(objCT)
End Sub

Private Sub Mistura_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Mistura_Click(objCT)
End Sub

Private Sub TempoMistura_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.TempoMistura_Click(objCT)
End Sub

Private Sub Maquina2_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Maquina2_GotFocus(objCT)
End Sub

Private Sub Maquina2_KeyPress(KeyAscii As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Maquina2_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Maquina2_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Maquina2_Validate(objCT, Cancel)
End Sub

Private Sub BotaoProdutoBase_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoProdutoBase_Click(objCT)
End Sub

Private Sub Agua_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Agua_Change(objCT)
End Sub

Private Sub Oleo_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Oleo_Change(objCT)
End Sub

Private Sub Amperagem_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Amperagem_Change(objCT)
End Sub

Private Sub VRDe_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.VRDe_Change(objCT)
End Sub

Private Sub VRAte_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.VRAte_Change(objCT)
End Sub

Private Sub Gate_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Gate_Change(objCT)
End Sub

Private Sub TCZonaDe_Change(Index As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TCZonaDe_Change(objCT, Index)
End Sub

Private Sub TCZonaAte_Change(Index As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TCZonaAte_Change(objCT, Index)
End Sub

Private Sub Tela_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Tela_Change(objCT)
End Sub

Private Sub TEZonaDe_Change(Index As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TEZonaDe_Change(objCT, Index)
End Sub

Private Sub TEZonaAte_Change(Index As Integer)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TEZonaAte_Change(objCT, Index)
End Sub

Private Sub VVDe_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.VVDe_Change(objCT)
End Sub

Private Sub VVAte_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.VVAte_Change(objCT)
End Sub

Private Sub Observacao_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Observacao_Change(objCT)
End Sub

Private Sub Amostras_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Amostras_Change(objCT)
End Sub

Private Sub Agua_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Agua_Validate(objCT, Cancel)
End Sub

Private Sub Oleo_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Oleo_Validate(objCT, Cancel)
End Sub

Private Sub Amperagem_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Amperagem_Validate(objCT, Cancel)
End Sub

Private Sub VRDe_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.VRDe_Validate(objCT, Cancel)
End Sub

Private Sub VRAte_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.VRAte_Validate(objCT, Cancel)
End Sub

Private Sub Gate_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.Gate_Validate(objCT, Cancel)
End Sub

Private Sub TCZonaDe_Validate(iIndice As Integer, Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TCZonaDe_Validate(objCT, iIndice, Cancel)
End Sub

Private Sub TCZonaAte_Validate(iIndice As Integer, Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TCZonaAte_Validate(objCT, iIndice, Cancel)
End Sub

Private Sub TEZonaDe_Validate(iIndice As Integer, Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TEZonaDe_Validate(objCT, iIndice, Cancel)
End Sub

Private Sub TEZonaAte_Validate(iIndice As Integer, Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TEZonaAte_Validate(objCT, iIndice, Cancel)
End Sub

Private Sub VVDe_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.VVDe_Validate(objCT, Cancel)
End Sub

Private Sub VVAte_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.VVAte_Validate(objCT, Cancel)
End Sub

'#########################################

'################################################################
'Inserido por Wagner 22/09/05
Private Sub BotaoFornecedor_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoFornecedor_Click(objCT)
End Sub

'##########################################################################

'################################################################
'Inserido por Wagner 10/02/2006
Private Sub BotaoLote_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoLote_Click(objCT)
End Sub

Private Sub BotaoCarga_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoCarga_Click(objCT)
End Sub

Private Sub BotaoLimpaCarga_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoLimpaCarga_Click(objCT)
End Sub

'#####################################
'Inserido por Wagner 03/08/2006
Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub

Sub LabelProjeto_Click()
    Call objCT.LabelProjeto_Click
End Sub

Sub Projeto_GotFocus()
    Call objCT.Projeto_GotFocus
End Sub

Sub Projeto_Change()
    Call objCT.Projeto_Change
End Sub

Sub Projeto_Validate(Cancel As Boolean)
    Call objCT.Projeto_Validate(Cancel)
End Sub

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long
    Call objCT.Obter_ItensPRJCR(colItensPRJ)
End Function

Sub Etapa_Change()
    Call objCT.Etapa_Change
End Sub

Sub Etapa_Click()
    Call objCT.Etapa_Click
End Sub

Sub Etapa_Validate(Cancel As Boolean)
    Call objCT.Etapa_Validate(Cancel)
End Sub
'#####################################

