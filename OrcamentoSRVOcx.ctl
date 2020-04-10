VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrcamentoSRVOcx 
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4950
      Index           =   1
      Left            =   180
      TabIndex        =   38
      Top             =   855
      Width           =   9165
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
         Left            =   105
         TabIndex        =   166
         Top             =   4590
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   4230
      End
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
         Left            =   4530
         TabIndex        =   165
         Top             =   4605
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1020
         Index           =   0
         Left            =   90
         TabIndex        =   53
         Top             =   150
         Width           =   8865
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   5355
            TabIndex        =   124
            Top             =   1020
            Visible         =   0   'False
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
            Left            =   3780
            TabIndex        =   123
            Top             =   1050
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton BotaoVerSolicitacaoSRV 
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
            Left            =   7185
            TabIndex        =   7
            Top             =   615
            Width           =   1005
         End
         Begin VB.CommandButton BotaoSolicitacaoSRV 
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
            Left            =   5460
            TabIndex        =   6
            Top             =   615
            Width           =   1350
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2190
            Picture         =   "OrcamentoSRVOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   225
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   8250
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   315
            Left            =   7200
            TabIndex        =   3
            Top             =   225
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1395
            TabIndex        =   0
            Top             =   225
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoBase 
            Height          =   315
            Left            =   4515
            TabIndex        =   2
            Top             =   210
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodSolicSRV 
            Height          =   315
            Left            =   4515
            TabIndex        =   5
            Top             =   630
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1830
            TabIndex        =   125
            Top             =   1050
            Visible         =   0   'False
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
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
            Height          =   315
            Left            =   1410
            TabIndex        =   129
            Top             =   615
            Width           =   525
         End
         Begin VB.Label LblNatOpInternaEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza OP.:"
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
            Left            =   165
            TabIndex        =   128
            Top             =   675
            Width           =   1215
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
            TabIndex        =   127
            Top             =   1080
            Visible         =   0   'False
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
            Left            =   1095
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   126
            Top             =   1095
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label LabelSolicitaSRV 
            AutoSize        =   -1  'True
            Caption         =   "Solicitação de Serviço:"
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
            Left            =   2430
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   660
            Width           =   1995
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
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   56
            Top             =   285
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
            Left            =   6390
            TabIndex        =   55
            Top             =   270
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
            Left            =   3195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   255
            Width           =   1200
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Preços"
         Height          =   690
         Index           =   2
         Left            =   105
         TabIndex        =   49
         Top             =   2235
         Width           =   8865
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   1365
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   690
         Index           =   6
         Left            =   105
         TabIndex        =   46
         Top             =   1365
         Width           =   8865
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4515
            TabIndex        =   9
            Top             =   225
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   315
            Left            =   1365
            TabIndex        =   8
            Top             =   225
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
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
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   270
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
            Left            =   3915
            TabIndex        =   47
            Top             =   270
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Outros"
         Height          =   1485
         Left            =   105
         TabIndex        =   39
         Top             =   3075
         Width           =   8865
         Begin VB.ComboBox Status 
            Height          =   315
            Left            =   1365
            TabIndex        =   15
            Top             =   645
            Width           =   4065
         End
         Begin VB.ComboBox MotivoPerda 
            Height          =   315
            Left            =   1365
            TabIndex        =   16
            Top             =   1065
            Width           =   4050
         End
         Begin MSMask.MaskEdBox PrazoValidade 
            Height          =   300
            Left            =   5160
            TabIndex        =   14
            Top             =   240
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
            Left            =   1365
            TabIndex        =   13
            Top             =   240
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
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
            Left            =   3480
            TabIndex        =   45
            Top             =   270
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
            TabIndex        =   44
            Top             =   300
            Width           =   885
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
            Left            =   6045
            TabIndex        =   43
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Usuario 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6825
            TabIndex        =   42
            Top             =   240
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
            Left            =   120
            TabIndex        =   41
            Top             =   1125
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
            Left            =   720
            TabIndex        =   40
            Top             =   690
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4950
      Index           =   2
      Left            =   45
      TabIndex        =   26
      Top             =   870
      Visible         =   0   'False
      Width           =   9405
      Begin VB.Frame FrameItens 
         BorderStyle     =   0  'None
         Caption         =   "Peças"
         Height          =   2520
         Index           =   2
         Left            =   105
         TabIndex        =   95
         Top             =   360
         Visible         =   0   'False
         Width           =   9045
         Begin MSMask.MaskEdBox PrecoTotalB 
            Height          =   225
            Left            =   6420
            TabIndex        =   159
            Top             =   1395
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
         Begin MSMask.MaskEdBox VersaoKit 
            Height          =   225
            Left            =   4140
            TabIndex        =   121
            Top             =   1815
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
            Left            =   5430
            TabIndex        =   122
            Top             =   1875
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
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5745
            MaxLength       =   50
            TabIndex        =   103
            Top             =   1065
            Width           =   1305
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1545
            MaxLength       =   250
            TabIndex        =   99
            Top             =   270
            Width           =   2490
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoSRVOcx.ctx":00EA
            Left            =   5340
            List            =   "OrcamentoSRVOcx.ctx":00EC
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   210
            Width           =   720
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   300
            TabIndex        =   100
            Top             =   285
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   4005
            TabIndex        =   101
            Top             =   240
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   6285
            TabIndex        =   102
            Top             =   255
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
            Left            =   4305
            TabIndex        =   104
            Top             =   1050
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   4065
            TabIndex        =   105
            Top             =   570
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
            Left            =   5190
            TabIndex        =   106
            Top             =   555
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2310
            Left            =   90
            TabIndex        =   97
            Top             =   120
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   4075
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
         Caption         =   "Totais"
         Height          =   1290
         Index           =   1
         Left            =   75
         TabIndex        =   132
         Top             =   3150
         Width           =   9285
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   90
            TabIndex        =   27
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
            TabIndex        =   133
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   30
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
            TabIndex        =   31
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
         Begin VB.Label ValorProdutos2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   156
            Top             =   405
            Width           =   1140
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
            TabIndex        =   155
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
            TabIndex        =   154
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
            Index           =   1
            Left            =   2790
            TabIndex        =   153
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
            Index           =   3
            Left            =   1470
            TabIndex        =   152
            Top             =   705
            Width           =   1125
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
            Index           =   5
            Left            =   105
            TabIndex        =   151
            Top             =   705
            Width           =   1125
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
            TabIndex        =   150
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label ISSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   149
            Top             =   405
            Width           =   1140
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
            Index           =   16
            Left            =   5430
            TabIndex        =   148
            Top             =   705
            Width           =   1125
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
            Index           =   6
            Left            =   4125
            TabIndex        =   147
            Top             =   705
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
            TabIndex        =   146
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
            Index           =   24
            Left            =   8100
            TabIndex        =   145
            Top             =   210
            Width           =   1065
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
            TabIndex        =   144
            Top             =   210
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
            TabIndex        =   143
            Top             =   210
            Width           =   1170
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
            TabIndex        =   142
            Top             =   195
            Width           =   1065
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
            TabIndex        =   141
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   140
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4065
            TabIndex        =   139
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2745
            TabIndex        =   138
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1410
            TabIndex        =   137
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   90
            TabIndex        =   136
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   135
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   134
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   157
            Top             =   405
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin VB.Frame FrameItens 
         BorderStyle     =   0  'None
         Caption         =   "Serviços"
         Height          =   2520
         Index           =   1
         Left            =   105
         TabIndex        =   36
         Top             =   360
         Width           =   9045
         Begin MSMask.MaskEdBox PrecoTotalSRVB 
            Height          =   225
            Left            =   6870
            TabIndex        =   158
            Top             =   1935
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
         Begin VB.ComboBox MotivoPerdaItemSRV 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4050
            TabIndex        =   120
            Top             =   1635
            Width           =   1905
         End
         Begin VB.ComboBox UnidadeMedSRV 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoSRVOcx.ctx":00EE
            Left            =   5310
            List            =   "OrcamentoSRVOcx.ctx":00F0
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   1320
            Width           =   720
         End
         Begin VB.TextBox DescServicoSRV 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6675
            MaxLength       =   250
            TabIndex        =   88
            Top             =   330
            Width           =   2490
         End
         Begin VB.TextBox ObservacaoSRV 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5280
            MaxLength       =   50
            TabIndex        =   84
            Top             =   315
            Width           =   1305
         End
         Begin VB.ComboBox StatusItemSRV 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5115
            TabIndex        =   83
            Top             =   585
            Width           =   1920
         End
         Begin MSMask.MaskEdBox VersaoKitSRV 
            Height          =   225
            Left            =   4095
            TabIndex        =   85
            Top             =   975
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
         Begin MSMask.MaskEdBox VersaoKitBaseSRV 
            Height          =   225
            Left            =   5370
            TabIndex        =   86
            Top             =   1020
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
         Begin MSMask.MaskEdBox ServicoSRV 
            Height          =   225
            Left            =   6015
            TabIndex        =   87
            Top             =   1725
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEntregaSRV 
            Height          =   225
            Left            =   4110
            TabIndex        =   89
            Top             =   300
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
         Begin MSMask.MaskEdBox DescontoSRV 
            Height          =   225
            Left            =   4050
            TabIndex        =   90
            Top             =   1365
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
         Begin MSMask.MaskEdBox PercentDescSRV 
            Height          =   225
            Left            =   4095
            TabIndex        =   91
            Top             =   660
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
         Begin MSMask.MaskEdBox PrecoUnitarioSRV 
            Height          =   225
            Left            =   7125
            TabIndex        =   92
            Top             =   1005
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
         Begin MSMask.MaskEdBox QuantidadeSRV 
            Height          =   225
            Left            =   6150
            TabIndex        =   93
            Top             =   1410
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
         Begin MSMask.MaskEdBox PrecoTotalSRV 
            Height          =   225
            Left            =   7155
            TabIndex        =   94
            Top             =   660
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
         Begin MSFlexGridLib.MSFlexGrid GridItensSRV 
            Height          =   2310
            Left            =   90
            TabIndex        =   37
            Top             =   120
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   4075
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
      Begin VB.CommandButton BotaoProdSolicSRV 
         Caption         =   "Serviços x Solicitaçao"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3780
         TabIndex        =   34
         Top             =   4485
         Visible         =   0   'False
         Width           =   1605
      End
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
         Height          =   420
         Left            =   75
         TabIndex        =   32
         Top             =   4485
         Width           =   1605
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   2985
         Left            =   75
         TabIndex        =   107
         Top             =   45
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   5265
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Serviços"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Peças"
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
         Height          =   420
         Left            =   1950
         TabIndex        =   33
         Top             =   4485
         Width           =   1605
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
         Height          =   420
         Left            =   5625
         TabIndex        =   35
         Top             =   4485
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4950
      Index           =   3
      Left            =   30
      TabIndex        =   57
      Top             =   870
      Visible         =   0   'False
      Width           =   9420
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   4515
         Left            =   60
         TabIndex        =   59
         Top             =   360
         Width           =   9330
         Begin VB.ComboBox TipoDesconto3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3105
            TabIndex        =   64
            Top             =   1845
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            TabIndex        =   63
            Top             =   1530
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   62
            Top             =   1215
            Width           =   1965
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3060
            Picture         =   "OrcamentoSRVOcx.ctx":00F2
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   615
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3060
            Picture         =   "OrcamentoSRVOcx.ctx":014C
            Style           =   1  'Graphical
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   765
            Width           =   240
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7470
            TabIndex        =   65
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
            TabIndex        =   66
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
            TabIndex        =   67
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
            TabIndex        =   68
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
            TabIndex        =   69
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
            TabIndex        =   70
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
            TabIndex        =   71
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
            TabIndex        =   72
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
            TabIndex        =   73
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
            TabIndex        =   74
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
            TabIndex        =   75
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
            Left            =   1950
            TabIndex        =   77
            Top             =   615
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
            Left            =   30
            TabIndex        =   78
            Top             =   1020
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4842
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
            Left            =   4425
            TabIndex        =   76
            Top             =   210
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
            Index           =   4
            Left            =   3495
            TabIndex        =   164
            Top             =   240
            Width           =   885
         End
         Begin VB.Label ValorOriginalTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1965
            TabIndex        =   163
            Top             =   210
            Width           =   1185
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
            Index           =   2
            Left            =   660
            TabIndex        =   162
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label ValorTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7980
            TabIndex        =   161
            Top             =   225
            Width           =   1185
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
            Index           =   7
            Left            =   6450
            TabIndex        =   160
            Top             =   255
            Width           =   1455
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
            Left            =   120
            TabIndex        =   79
            Top             =   660
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
         Left            =   60
         TabIndex        =   58
         Top             =   105
         Value           =   1  'Checked
         Width           =   3360
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
      Height          =   495
      Left            =   4545
      TabIndex        =   19
      Top             =   30
      Width           =   1605
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
      Left            =   2190
      TabIndex        =   18
      Top             =   135
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4680
      Index           =   5
      Left            =   135
      TabIndex        =   130
      Top             =   1080
      Visible         =   0   'False
      Width           =   9195
      Begin TelasSRV.TabTributacaoFat TabTrib 
         Height          =   4560
         Left            =   105
         TabIndex        =   131
         Top             =   45
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   8043
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Private Sub Observacao_Change()Private Sub Observacao_Change()"
      Height          =   4680
      Index           =   4
      Left            =   120
      TabIndex        =   108
      Top             =   1065
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   3630
         Left            =   30
         TabIndex        =   110
         Top             =   75
         Width           =   9180
         Begin VB.ComboBox BLGTipoBloqueio 
            Height          =   315
            ItemData        =   "OrcamentoSRVOcx.ctx":01A6
            Left            =   225
            List            =   "OrcamentoSRVOcx.ctx":01A8
            TabIndex        =   114
            Top             =   585
            Width           =   1440
         End
         Begin VB.TextBox BLGObservacao 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2715
            MaxLength       =   250
            TabIndex        =   113
            Top             =   1110
            Width           =   2715
         End
         Begin VB.TextBox BLGSeqBloqueio 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   7185
            MaxLength       =   250
            TabIndex        =   112
            Top             =   1140
            Width           =   585
         End
         Begin MSMask.MaskEdBox BLGResponsavelLib 
            Height          =   270
            Left            =   7050
            TabIndex        =   115
            Top             =   585
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BLGDataLiberacao 
            Height          =   270
            Left            =   5850
            TabIndex        =   116
            Top             =   615
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSMask.MaskEdBox BLGUsuario 
            Height          =   270
            Left            =   3000
            TabIndex        =   117
            Top             =   615
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BLGResponsavel 
            Height          =   270
            Left            =   4470
            TabIndex        =   118
            Top             =   615
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BLGDataBloqueio 
            Height          =   270
            Left            =   1875
            TabIndex        =   119
            Top             =   615
            Width           =   1065
            _ExtentX        =   1879
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
         Begin MSFlexGridLib.MSFlexGrid BLGGridBloqueio 
            Height          =   2715
            Left            =   75
            TabIndex        =   111
            Top             =   165
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   4789
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BLGBotaoLibBloq 
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
         Left            =   7530
         Picture         =   "OrcamentoSRVOcx.ctx":01AA
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   3825
         Width           =   1650
      End
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
      Left            =   195
      TabIndex        =   17
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6210
      ScaleHeight     =   450
      ScaleWidth      =   3150
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   45
      Width           =   3210
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   579
         Picture         =   "OrcamentoSRVOcx.ctx":27A4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   1098
         Picture         =   "OrcamentoSRVOcx.ctx":28A6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1617
         Picture         =   "OrcamentoSRVOcx.ctx":2A00
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   2136
         Picture         =   "OrcamentoSRVOcx.ctx":2B8A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2655
         Picture         =   "OrcamentoSRVOcx.ctx":30BC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoEmail 
         Height          =   345
         Left            =   60
         Picture         =   "OrcamentoSRVOcx.ctx":323A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Enviar email"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5325
      Left            =   0
      TabIndex        =   81
      Top             =   540
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9393
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
            Caption         =   "Bloqueio"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
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
Attribute VB_Name = "OrcamentoSRVOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gcolProdSolicSRV As New Collection

Dim m_objUserControl As Object

Public gobjAnotacao As ClassAnotacoes 'Wagner
Public gobjInfoUsu As Object 'Inserido por Wagner

Public gobjTelaProjetoInfo As ClassTelaPRJInfo
Public gobjInfoAdic As ClassInfoAdic
Dim dValorDescontoItensAnt As Double
Dim dPercDescontoItensAnt As Double
Dim iDescontoAlterado As Integer

Dim gobjTelaTrataBloqGen As ClassTelaTrataBloqGen

'Grid ItensSRV
Public iGrid_ItemServicoSRV_Col As Integer
Public iGrid_ServicoSRV_Col As Integer
Public iGrid_DescServicoSRV_Col As Integer
Public iGrid_UnidadeMedSRV_Col As Integer
Public iGrid_QuantidadeSRV_Col As Integer
Public iGrid_PrecoUnitarioSRV_Col As Integer
Public iGrid_PercDescSRV_Col As Integer
Public iGrid_DescontoSRV_Col As Integer
Public iGrid_PrecoTotalSRV_Col As Integer
Public iGrid_PrecoTotalSRVB_Col As Integer
Public iGrid_DataEntregaSRV_Col As Integer
Public iGrid_StatusItemSRV_Col As Integer
Public iGrid_MotivoPerdaItemSRV_Col As Integer
Public iGrid_ObservacaoSRV_Col As Integer
Public iGrid_VersaoKitSRV_Col As Integer
Public iGrid_VersaoKitBaseSRV_Col As Integer

'Grid Itens
Public iGrid_ItemProduto_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_Observacao_Col As Integer
Public iGrid_VersaoKit_Col As Integer
Public iGrid_VersaoKitBase_Col As Integer

'#######################################
'INSERIDO POR WAGNER
Const VALOR_CODIGO = 1
Const VALOR_CODIGOBASE = 0

Private WithEvents objEventoNumeroBase As AdmEvento
Attribute objEventoNumeroBase.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1
Private WithEvents objEventoNumeroBaseSRV As AdmEvento
Attribute objEventoNumeroBaseSRV.VB_VarHelpID = -1
Private WithEvents objEventoVersaoSRV As AdmEvento
Attribute objEventoVersaoSRV.VB_VarHelpID = -1
'#######################################

'Grid Parcelas
Public iGrid_Vencimento_Col  As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_Desc1Codigo_Col As Integer
Public iGrid_Desc1Ate_Col As Integer
Public iGrid_Desc1Valor_Col As Integer
Public iGrid_Desc1Perc_Col As Integer
Public iGrid_Desc2Codigo_Col As Integer
Public iGrid_Desc2Ate_Col As Integer
Public iGrid_Desc2Valor_Col As Integer
Public iGrid_Desc2Perc_Col As Integer
Public iGrid_Desc3Codigo_Col As Integer
Public iGrid_Desc3Ate_Col As Integer
Public iGrid_Desc3Valor_Col As Integer
Public iGrid_Desc3Perc_Col As Integer

Dim giTrazendoTribTela As Integer
Dim giFrameAtual As Integer
'Dim giFrameAtualTributacao As Integer
Dim gsCodigoAnterior  As String
Dim giFrameAtualItens As Integer

Public iAlterado As Integer
Public iVendedorAlterado As Integer

Dim giClienteAlterado As Integer
Dim giFilialAlterada As Integer
Dim gdDesconto As Double
Dim giValorFreteAlterado As Integer
Dim giValorSeguroAlterado As Integer
Dim giValorDescontoAlterado As Integer
Dim giValorDespesasAlterado  As Integer
Dim giDataReferenciaAlterada As Integer
Dim giNaturezaOpAlterada As Integer

Dim giValorDescontoManual As Integer

Dim giPercAcresFinAlterado As Integer

Public gobjOrcamentoVenda As New ClassOrcamentoVenda 'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
    'todos os dados relevantes p/tributacao dentro de gobjOrcamentoVenda estarao sincronizados com a tela antes da atualizacao da tributacao

Public objGridItens As AdmGrid
Dim objGridParcelas As AdmGrid
Public objGridItensSRV As AdmGrid

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
'Private WithEvents objEventoNaturezaOp As AdmEvento
'Private WithEvents objEventoTiposDeTributacao As AdmEvento
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCodSolicSRV As AdmEvento
Attribute objEventoCodSolicSRV.VB_VarHelpID = -1

Dim giLinhaAnterior As Integer
'Dim giRecalculandoTributacao As Integer
'Dim gcolTiposTribICMS As New Collection
'Dim gcolTiposTribIPI As New Collection

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private giExercicio As Integer, giPeriodo As Integer
Private gcolAlmoxFilial As New Collection
Private gobjGeracaoNFiscal As ClassGeracaoNFiscal

'Constantes públicas dos tabs
Private Const TAB_Principal = 1
Private Const TAB_Itens = 2
Private Const TAB_Cobranca = 3
Private Const TAB_Tributacao = 4

'Property Variables:
Dim m_Caption As String
Public gbCarregandoTela As Boolean
Public iFrameAtual As Integer
Dim bTrazendoDoc As Boolean
Dim giPosCargaOk As Integer

Public gobjTribTab As New ClassTribTab
'Dim iValorIRRFAlterado As Integer
'
''Incluidos por Leo em 30/04/02 para tratamento da tributação
'Dim giISSAliquotaAlterada As Integer
'Dim giISSValorAlterado As Integer
'Dim giValorIRRFAlterado As Integer
'Dim giTipoTributacaoAlterado As Integer
'Dim giAliqIRAlterada As Integer
'Dim iPISRetidoAlterado As Integer
'Dim iCOFINSRetidoAlterado As Integer
'Dim iCSLLRetidoAlterado As Integer
'
'Dim giTrazendoTribItemTela As Integer 'por Leo em 02/05/02
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

Event Unload()

Function Trata_Parametros(Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim objOrcamentoVendaAux As New ClassOrcamentoVenda

On Error GoTo Erro_Trata_Parametros

    lErro = CargaPosFormLoad(True)
    If lErro <> SUCESSO Then gError 186129

    If Not (objOrcamentoVenda Is Nothing) Then

        'Se foi passado o código do orçamento
        If objOrcamentoVenda.lCodigo > 0 Then

            objOrcamentoVendaAux.lCodigo = objOrcamentoVenda.lCodigo
            objOrcamentoVendaAux.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa

            'Coloca o Orcamento na tela
            lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVendaAux, VALOR_CODIGO)
            If lErro <> SUCESSO And lErro <> 84363 Then gError 186130

            If lErro <> SUCESSO Then  'Não encontrou no BD o código de Pedido

                'Limpa a tela e coloca o código na Tela
                Call Limpa_OrcamentoVenda
                Codigo.Text = CStr(objOrcamentoVenda.lCodigo)

            End If

        'Se foi passado o código do cliente
        ElseIf objOrcamentoVenda.lCliente > 0 Then

            'Joga o código do cliente na tela
            Cliente.Text = objOrcamentoVenda.lCliente
            Call Cliente_Validate(bSGECancelDummy)

            'Se foi passada uma filial de cliente
            If objOrcamentoVenda.iFilial > 0 Then

                'Joga a filial do cliente na tela
                Filial.Text = objOrcamentoVenda.iFilial
                Call Filial_Validate(bSGECancelDummy)
            End If

            'Cria um número automático para o orçamento
            Call BotaoProxNum_Click

        End If

    End If

    iAlterado = 0
    iVendedorAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 186129, 186130

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186131)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objContainer As Object
Dim objUserControl As Object

On Error GoTo Erro_Form_Load

    'precodesc
'    Set objContainer = Frame2(3)
'    Set objUserControl = Me
'
'    'funcao usada para alguns clientes customizarem o  Form_Load
'    lErro = CF("Orcamento_Form_Load", objUserControl, objContainer)
'    If lErro <> SUCESSO Then gError 126500

    '#############################
    'INSERIDO POR WAGNER
    Set objEventoNumeroBase = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Usuario.Caption = gsUsuario

    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        BotaoVersaoKitBase.Visible = True
    Else
        BotaoVersaoKitBase.Visible = False
    End If
    '#############################

    giPosCargaOk = 0

    giFrameAtual = 1
    'giFrameAtualTributacao = 1
    giFrameAtualItens = 1

    If gobjFAT.iExibeOPImpPreco = MARCADO Then
        ImprimirOVComPreco.Visible = True
    End If
    
    'Preenche Data Referencia e Data de Emissão coma Data Atual
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    giDataReferenciaAlterada = 0

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    lErro = Carrega_TipoDesconto
    If lErro <> SUCESSO Then gError 186148

    '##########################
    'Inserido por Wagner 09/12/05
    Call Carrega_Status(Status)
    Call Carrega_Status(StatusItemSRV)
    Call Carrega_Motivo(MotivoPerda)
    Call Carrega_Motivo(MotivoPerdaItemSRV)

'    lErro = CF2(Me, "OV_Form_Load")
'    If lErro <> SUCESSO Then gError 141406
    '##########################

    iAlterado = 0

    iVendedorAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 186148

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186149)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Carrega_TipoDesconto() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    Set colCodigoDescricao = gobjCRFAT.colTiposDesconto

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item nas List's das Combos de Tipos Desconto
        TipoDesconto1.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto1.ItemData(TipoDesconto1.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto2.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto2.ItemData(TipoDesconto2.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto3.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto3.ItemData(TipoDesconto3.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186150)

    End Select

    Exit Function

End Function

Private Function Carrega_CondicaoPagamento() As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 186151

   For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na List da Combo CondicaoPagamento
        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 186151

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186152)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjOrcamentoVenda = Nothing

    Set gobjTelaProjetoInfo = Nothing
    Set gobjTelaTrataBloqGen = Nothing
    
    Set gobjInfoAdic = Nothing
    
    'trecho incluido por Leo em 20/03/02 - Filtrar
    Set objEventoCliente = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoServico = Nothing
    Set objEventoVendedor = Nothing
'    Set objEventoNaturezaOp = Nothing
'    Set objEventoTiposDeTributacao = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCodSolicSRV = Nothing
    
    'giGridBloqueioLinhasBloqueadas = 0

    '###############################################
    'INSERIDO POR WAGNER
    Set objEventoNumeroBase = Nothing
    Set objEventoVersao = Nothing
    '###############################################

'    Set objEventoTiposDeTributacao = Nothing
'    Set gcolTiposTribICMS = Nothing
'    Set gcolTiposTribIPI = Nothing
'
'    'Encerra tributacao
'    Call TributacaoOV_Terminar

    Set gobjAnotacao = Nothing 'Inserido por Wagner
    Set gobjInfoUsu = Nothing 'Inserido por Wagner
    
    Set gobjTribTab = Nothing

    Call ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Orçamento de Serviço"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OrcamentoSRV"

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
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
End Property

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 186153

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 186153

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186154)

    End Select

    Exit Function

End Function

'Function Carrega_FilialEmpresa() As Long
'
'Dim colCodigoDescricao As New AdmColCodigoNome
'Dim objCodDescricao As AdmCodigoNome
'Dim lErro As Long
'
'On Error GoTo Erro_Carrega_FilialEmpresa
'
'    'Lê o código e a descrição de todas as Filiais
'    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAISEMPRESA_NOME, colCodigoDescricao)
'    If lErro <> SUCESSO Then gError 84093
'
'    For Each objCodDescricao In colCodigoDescricao
'
'        'Adiciona a filial na Lista de Filiais
'        Filial.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        Filial.ItemData(Filial.NewIndex) = objCodDescricao.iCodigo
'
'    Next
'
'    Carrega_FilialEmpresa = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_FilialEmpresa:
'
'    Carrega_FilialEmpresa = gErr
'
'    Select Case gErr
'
'        Case 84093
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_FILIALEMPRESA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157723)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub BotaoDataReferenciaDown_Click()

Dim lErro As Long
Dim bCancel As Boolean
Dim sData As String

On Error GoTo Erro_BotaoDataReferenciaDown_Click

    sData = DataReferencia.Text

    'diminui a data em um dia
    lErro = Data_Up_Down_Click(DataReferencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 186155

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaDown_Click:

    Select Case gErr

        Case 186155

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186156)

    End Select

    Exit Sub

End Sub

Public Sub BotaoDataReferenciaUp_Click()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoDataReferenciaUp_Click

    sData = DataReferencia.Text

    'aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataReferencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 186157

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaUp_Click:

    Select Case gErr

        Case 186157

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186158)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código foi informado
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 186159

    objOrcamentoVenda.lCodigo = StrParaLong(Codigo.Text)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    Set objOrcamentoVenda.objTela = Me

    'Lê o Orcamento
    lErro = CF("OrcamentoSRV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186534 Then gError 186160

    If lErro <> SUCESSO Then gError 186161

    'Se o orçamento estiver vinculado a um Pedido de Venda, não poderá ser excluido -> Erro.
    If objOrcamentoVenda.lNumIntPedVenda <> 0 Then

        gError 186162

    'Se o orçamento estiver vinculado a uma Nota Fiscal, não poderá ser excluido -> Erro.
    ElseIf objOrcamentoVenda.lNumIntNFiscal <> 0 Then

        gError 186163

    End If
    'leo *********

    'Pede a confirmação da exclusão do Orcamento de Venda
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ORCAMENTO_VENDA", objOrcamentoVenda.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Faz a exclusão do Orcamento de Venda
    lErro = CF("OrcamentoSRV_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186164

    'Limpa a Tela de Orcamento de Venda
    Call Limpa_OrcamentoVenda

    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 186159
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 186160, 186164

        Case 186161
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)

        Case 186162
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO_VINCULADO_PEDIDO", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case 186163
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO_VINCULADO_NFISCAL", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186165)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 186166

    'Limpa a Tela
    Call Limpa_OrcamentoVenda

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 186166

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186167)

    End Select

    Exit Sub

End Sub


Private Sub BotaoProdSolicSRV_Click()
    
Dim lErro As Long
Dim iIndice As Integer
Dim colProdSolicSRVOrigem As New Collection
Dim objProdSolicSRV As ClassProdSolicSRV
Dim sServico As String
Dim iPreenchido As Integer
Dim objSolicSRV As New ClassSolicSRV
    
On Error GoTo Erro_BotaoProdSolicSRV_Click
    
    'Verifica se Solicitacao de Servico existe
    If Len(Trim(CodSolicSRV.Text)) > 0 Then
    
        objSolicSRV.iFilialEmpresa = giFilialEmpresa
        objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)
        
        lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 183258 Then gError 186969

        If lErro <> SUCESSO Then gError 186970
        
    End If
    
    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col), sServico, iPreenchido)
        If lErro <> SUCESSO Then gError 186971

        Set objProdSolicSRV = New ClassProdSolicSRV

        objProdSolicSRV.sServicoOrcSRV = sServico
        objProdSolicSRV.dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))
        
        colProdSolicSRVOrigem.Add objProdSolicSRV

    Next
    
    Call Chama_Tela("ProdSolic", colProdSolicSRVOrigem, objSolicSRV.lCodigo, gcolProdSolicSRV, Me)
    
    
    Exit Sub
    
Erro_BotaoProdSolicSRV_Click:

    Select Case gErr

        Case 186968, 186969, 186871

        Case 186970
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186170)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSolicitacaoSRV_Click()

Dim lErro As Long
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_BotaoSolicitacaoSRV_Click

    'Verfifica se o Pedido e a FilialPedido foram informadas
    If Len(Trim(CodSolicSRV.Text)) = 0 Then gError 186998

    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)

    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 186999
    
    If lErro <> SUCESSO Then gError 188000

    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_SolicSRV_Tela(objSolicSRV)
    If lErro <> SUCESSO Then gError 188001
    
    Exit Sub

Erro_BotaoSolicitacaoSRV_Click:

    Select Case gErr

        Case 186998
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_PREENCHIDA", gErr)

        Case 186999, 188001

        Case 188000
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188002)

    End Select

    Exit Sub

End Sub

Public Sub BotaoVersaoKitBase_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVersaoKitBase_Click

    If FrameItens(1).Visible = True Then

        If Not (GridItensSRV.Row > 0 And GridItensSRV.Row <= objGridItensSRV.iLinhasExistentes) Then Exit Sub
    
        lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 186753
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objKit.sProdutoRaiz = sProdutoFormatado
    
            colSelecao.Add sProdutoFormatado
    
            Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersaoSRV)
    
        Else
             gError 186752
    
        End If

    Else

        If Not (GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes) Then Exit Sub
    
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 186168
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objKit.sProdutoRaiz = sProdutoFormatado
    
            colSelecao.Add sProdutoFormatado
    
            Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
        Else
             gError 186169
    
        End If

    End If

    Exit Sub

Erro_BotaoVersaoKitBase_Click:

    Select Case gErr

        Case 186168, 186753

        Case 186169, 186752
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186170)

    End Select

    Exit Sub


End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 186171

    'Limpa a Tela
    Call Limpa_OrcamentoVenda
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 186171

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186172)

    End Select

    Exit Sub


End Sub

'mario
Private Sub Limpa_OrcamentoVenda()

Dim lErro As Long

On Error GoTo Erro_OrcamentoVenda

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_OrcamentoVenda2

    'Call BotaoGravarTrib

    Set gobjAnotacao = Nothing 'Inserido por Wagner

    iAlterado = 0

    iVendedorAlterado = 0

    Exit Sub

Erro_OrcamentoVenda:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186173)

    End Select

    Exit Sub

End Sub

'mario
Private Sub Limpa_OrcamentoVenda2()
'Limpa os campos da tela sem fechar o sistema de setas

Dim iIndice As Integer

    Call Limpa_Tela(Me)

    Codigo.Enabled = True
    Codigo.Text = ""
    Filial.Clear

'    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
'    Set gobjTelaProjetoInfo.objUserControl = objUserControl
'    Set gobjTelaProjetoInfo.objTela = Me

    Set gobjInfoAdic = Nothing

    CobrancaAutomatica.Value = vbChecked
    ValorTotal.Caption = ""
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
    ValorOriginalTit.Caption = ""
    ValorTit.Caption = ""
    CondicaoPagamento.Text = ""
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
    'Trecho Incluido por Leo em 22/03/02
    Vendedor.PromptInclude = False
    Vendedor.Text = ""
    Vendedor.PromptInclude = True
    TabelaPreco.Text = ""

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

    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridItensSRV)
    Call Grid_Limpa(objGridParcelas)

'    'Resseta tributação
'    Call TributacaoOrcSRV_Reset

    ImprimirOVComPreco.Value = vbChecked
    ImprimirOVComCodProd.Value = vbUnchecked
    
    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
    iAlterado = 0
    giValorDescontoAlterado = 0
    giClienteAlterado = 0
    giFilialAlterada = 0
    giDataReferenciaAlterada = 0

    giValorDescontoManual = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0


    '##########################
    'Inserido por Wagner 09/12/05
    Status.ListIndex = -1
    MotivoPerda.ListIndex = -1
    '##########################

    Set gcolProdSolicSRV = New Collection
    'giGridBloqueioLinhasBloqueadas = 0
    
    Call gobjTelaTrataBloqGen.Limpa

    Call gobjTribTab.Limpa_Tela

    Exit Sub

End Sub

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoServicos_Click

    If Me.ActiveControl Is ServicoSRV Then

        sProduto1 = ServicoSRV.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridItensSRV.Row = 0 Then gError 186174

        sProduto1 = GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 186175

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza=?"

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico, sSelecaoSQL)

    Exit Sub

Erro_BotaoServicos_Click:

    Select Case gErr

        Case 186174
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 186175

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186176)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lOrcamentoSRV As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_CODIGO_ORCAMENTOSRV", "OrcamentoSRV", "Codigo", lOrcamentoSRV)
    If lErro <> SUCESSO Then gError 186176

    Codigo.Text = lOrcamentoSRV

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 186176

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186177)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVerSolicitacaoSRV_Click()

Dim lErro As Long
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_BotaoVerSolicitacaoSRV_Click

    If Len(Trim(CodSolicSRV.Text)) = 0 Then gError 186994
        objSolicSRV.iFilialEmpresa = giFilialEmpresa
    objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)

    
    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 186995
    
    If lErro <> SUCESSO Then gError 186996
    
    'Chama  a tela de Solicitacao de Servico a solicitacao da tela
    Call Chama_Tela("SolicitacaoSRV", objSolicSRV)
    
    Exit Sub

Erro_BotaoVerSolicitacaoSRV_Click:

    Select Case gErr
    
        Case 186994
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_PREENCHIDA", gErr)
            
        Case 186995
        
        Case 186996
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186997)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    giClienteAlterado = 1

    Call Cliente_Preenche

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long, sNatOp As String, iTipoTrib As Integer
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Cliente_Validate

    If giClienteAlterado = 1 Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then

            'Busca o Cliente no BD
            lErro = TP_Cliente_Le_Orcamento(Cliente, objcliente, iCodFilial)
            If lErro <> SUCESSO And lErro <> 94449 And lErro <> 94450 And lErro <> 94451 And lErro <> 94452 Then gError 186178

            If lErro = SUCESSO Then

                'gobjOrcamentoVenda.lCliente = objcliente.lCodigo

                lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
                If lErro <> SUCESSO Then gError 186179

                'Preenche ComboBox de Filiais
                Call CF("Filial_Preenche", Filial, colCodigoNome)

                If Not gbCarregandoTela Then

                    If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then

                        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

                        'Seleciona filial na Combo Filial
                        Call CF("Filial_Seleciona", Filial, iCodFilial)

                    End If

                End If

                'Se o Tipo estiver preenchido
                If objcliente.iTipo > 0 Then
                    objTipoCliente.iCodigo = objcliente.iTipo
                    'Lê o Tipo de Cliente
                    lErro = CF("TipoCliente_Le", objTipoCliente)
                    If lErro <> SUCESSO And lErro <> 19062 Then gError 186180
                End If


                giValorDescontoManual = 0

                'Guarda o valor do desconto do cliente
                If objcliente.dDesconto > 0 Then

                    gdDesconto = objcliente.dDesconto

                ElseIf objTipoCliente.dDesconto > 0 Then

                    gdDesconto = objTipoCliente.dDesconto

                Else

                    gdDesconto = 0

                End If

                If Not gbCarregandoTela Then

                    Call DescontoGlobal_Recalcula

                    'ATualiza o total com o novo desconto
                    lErro = ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 186181

                    'Incluído por Luiz Nogueira em 26/01/04
                    'Coloca na tela a tabela do cliente
                    If objcliente.iVendedor > 0 Then

                        Vendedor.Text = objcliente.iVendedor
                        Call Vendedor_Validate(bSGECancelDummy)

                    ElseIf objTipoCliente.iVendedor > 0 Then

                        Vendedor.Text = objTipoCliente.iVendedor
                        Call Vendedor_Validate(bSGECancelDummy)

                    End If

                    'Coloca na tela a tabela do cliente
                    If objcliente.iTabelaPreco > 0 Then

                        TabelaPreco.Text = objcliente.iTabelaPreco
                        Call TabelaPreco_Validate(bSGECancelDummy)

                    ElseIf objTipoCliente.iTabelaPreco > 0 Then

                        TabelaPreco.Text = objTipoCliente.iTabelaPreco
                        Call TabelaPreco_Validate(bSGECancelDummy)

                    End If

                    'Se cobrança automática estiver selecionada preenche a CondPagto e dispara o Validate
                    If CobrancaAutomatica.Value = MARCADO Then

                        If objcliente.iCondicaoPagto > 0 Then

                            CondicaoPagamento.Text = objcliente.iCondicaoPagto
                            Call CondicaoPagamento_Validate(bSGECancelDummy)

                        ElseIf objTipoCliente.iCondicaoPagto > 0 Then

                            CondicaoPagamento.Text = objTipoCliente.iCondicaoPagto
                            Call CondicaoPagamento_Validate(bSGECancelDummy)

                        End If

                    End If

                End If

                giClienteAlterado = 0

            Else

                gobjOrcamentoVenda.lCliente = 0
                giValorDescontoManual = 0
                gdDesconto = 0

                If Not gbCarregandoTela Then

                    Call DescontoGlobal_Recalcula

                    'ATualiza o total com o novo desconto
                    lErro = ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 186182

                    TabelaPreco.ListIndex = -1

                    objcliente.lCodigo = 0

                    Filial.Clear

                End If

            End If

            'Leo em 17/04/02 até aqui

        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear

        End If

        giClienteAlterado = 0


        If Not gbCarregandoTela Then
    
    ''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then Call gobjTribTab.FilialCliente_Alterada(objcliente.lCodigo, iCodFilial) '####
    '*** fim tributacao
        End If
    
    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 186178 To 186182

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186183)

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
'            gobjOrcamentoVenda.dValorDesconto = dValorDesconto
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

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        dValorTotal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))

        dValor = dValor + dValorTotal

    Next

    dValorProdutos = dValor

End Sub

Public Sub CobrancaAutomatica_Click()

  iAlterado = REGISTRO_ALTERADO

  Call Cobranca_Automatica

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

    '################################
    'INSERIDO POR WAGNER
    Call Atualiza_VersaoKit_Grid
    '################################

    Call Atualiza_VersaoKitSRV_Grid

End Sub

Public Sub Codigo_GotFocus()

    gsCodigoAnterior = Codigo.Text
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub CodigoBase_Validate(Cancel As Boolean)

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim lErro As Long

On Error GoTo Erro_CodigoBase_Validate

    If StrParaLong(CodigoBase.Text) <> 0 Then

        objOrcamentoVenda.lCodigo = StrParaLong(CodigoBase.Text)
        objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

        'Lê os dados do Orcamento de Venda
        lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101232 Then gError 186184

        If lErro = 101232 Then gError 186185

        lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGOBASE)
        If lErro <> SUCESSO Then gError 186186

    End If

    Exit Sub

Erro_CodigoBase_Validate:

    Cancel = True

    Select Case gErr

        Case 131184, 131186

        Case 131185
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186187)

      End Select

    Exit Sub

End Sub

Private Sub CodSolicSRV_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodSolicSRV_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodSolicSRV, iAlterado)
End Sub

Private Sub CodSolicSRV_Validate(Cancel As Boolean)

Dim objSolicSRV As New ClassSolicSRV
Dim lErro As Long

On Error GoTo Erro_CodSolicSRV_Validate

    If Len(Trim(CodSolicSRV.Text)) > 0 Then

        objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)
        objSolicSRV.iFilialEmpresa = giFilialEmpresa

        'Lê no BD os dados da solicitacao em questao
        lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 183258 Then gError 193465
    
        'Se não encontrou a solicitacao => erro
        If lErro <> SUCESSO Then gError 193466

    End If
    
    Exit Sub
    
Erro_CodSolicSRV_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 193465
    
        Case 193466
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193467)
    
    End Select
    
    Exit Sub

End Sub

'Public Sub ComboICMSTipo_Click()
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
'Public Sub ComboIPITipo_Click()
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

Public Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CondicaoPagamento_Click()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_CondicaoPagamento_Click

    'Verifica se alguma Condição foi selecionada
    If CondicaoPagamento.ListIndex = -1 Then Exit Sub

    'Passa o código da Condição para objCondicaoPagto
    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 186188

    If lErro = 19205 Then gError 186189

    'Altera PercAcrescFin
    If Len(Trim(PercAcrescFin.ClipText)) > 0 Then

        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
        If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
    Else
        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
        Call PercAcrescFin_Validate(bSGECancelDummy)
    End If

    'Testa se ValorTotal está preenchido
    If Len(Trim(ValorTotal)) > 0 Then
        'Testa se DataReferencia está preenchida e ValorTotal é positivo
        If Len(Trim(DataReferencia.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then

            'Preenche o GridParcelas
            lErro = Cobranca_Automatica()
            If lErro <> SUCESSO Then gError 186190

        End If
    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 186188, 186190

        Case 186189
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186191)

      End Select

    Exit Sub

End Sub

Private Function Cobranca_Automatica() As Long
'recalcula o tab de cobranca

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Cobranca_Automatica

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 Then

        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)

        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 186192

        If lErro <> SUCESSO Then gError 186193

        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 186194

    End If

    Cobranca_Automatica = SUCESSO

    Exit Function

Erro_Cobranca_Automatica:

    Cobranca_Automatica = gErr

    Select Case gErr

        Case 186192, 186194

        Case 186193
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186195)

    End Select

    Exit Function

End Function

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim iIndice As Integer
Dim dPercAcrescFin As Double
Dim iTamanho As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna As Integer

Dim colDescontoPadrao As New Collection

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)

    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 186196
    
    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        objCondicaoPagto.dValorTotal = dValorPagar

        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then gError 186196

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    'Se Data Referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        objCondicaoPagto.dtDataRef = CDate(DataReferencia.Text)

        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 186197

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

        Next

    End If

    ' Se dValorPagar>0 coloca desconto padrao (quantos houver, se houver) em todas as parcelas.
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError 186198
    Next

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 186196 To 186198

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186199)

    End Select

    Exit Function

End Function

Function Preenche_DescontoPadrao(iLinha As Integer) As Long

Dim lErro As Long
Dim colDescontoPadrao As New ColDesconto
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna  As Integer
Dim dtDataVencimento As Date
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Preenche_DescontoPadrao

    'Se a data de referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col))
        lErro = CF("Parcela_GeraDescontoPadrao", colDescontoPadrao, dtDataVencimento)
        If lErro <> SUCESSO Then gError 186200

        If colDescontoPadrao.Count > 0 Then

            'Para cada um dos desontos padrão
            For iIndice1 = 1 To colDescontoPadrao.Count

                'Seleciona a coluna correspondente ao Desconto
                If iIndice1 = 1 Then iColuna = iGrid_Desc1Codigo_Col
                If iIndice1 = 2 Then iColuna = iGrid_Desc2Codigo_Col
                If iIndice1 = 3 Then iColuna = iGrid_Desc3Codigo_Col

                'Seleciona o tipo de desconto
                For iIndice2 = 0 To TipoDesconto1.ListCount - 1
                    If colDescontoPadrao.Item(iIndice1).iCodigo = TipoDesconto1.ItemData(iIndice2) Then
                        GridParcelas.TextMatrix(iLinha, iColuna) = TipoDesconto1.List(iIndice2)
                        GridParcelas.TextMatrix(iLinha, iColuna + 1) = Format(colDescontoPadrao.Item(iIndice1).dtData, "dd/mm/yyyy")
                        GridParcelas.TextMatrix(iLinha, iColuna + 3) = Format(colDescontoPadrao.Item(iIndice1).dValor, "Percent")

                        '*** Inicio colocacao Valor Desconto na tela
                        dPercentual = colDescontoPadrao.Item(iIndice1).dValor
                        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))

                        'Coloca Valor do Desconto na tela
                        If dValorParcela > 0 Then
                            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
                            GridParcelas.TextMatrix(iLinha, iColuna + 2) = sValorDesconto
                        End If
                        '*** Fim colocacao Valor Desconto na tela

                    End If
                Next
            Next

        End If

    End If

    Preenche_DescontoPadrao = SUCESSO

    Exit Function

Erro_Preenche_DescontoPadrao:

    Preenche_DescontoPadrao = gErr

    Select Case gErr

        Case 186200

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186201)

    End Select

    Exit Function

End Function

Public Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_Condicaopagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 186201

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 186202

        If lErro = 19205 Then gError 186203

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 186204

        'Coloca na Tela
        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida

        'Altera PercAcrescFin
        If Len(Trim(PercAcrescFin.ClipText)) > 0 Then
            dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
            If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
                If vbMsgRes = vbYes Then
                    PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                   Call PercAcrescFin_Validate(bSGECancelDummy)
                End If
            End If
        Else
            PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
            Call PercAcrescFin_Validate(bSGECancelDummy)
        End If

        'Se ValorTotal e DataReferencia estiverem preenchidos, preenche GridParcelas
        If Len(Trim(ValorTotal)) > 0 Then
            If Len(Trim(DataReferencia.ClipText)) > 0 And CLng(ValorTotal.Caption) > 0 Then

                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 186205

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 186206

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case gErr

        Case 186201, 186202, 186205

        Case 186206
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

       Case 186203
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 186204
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186207)

    End Select

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
    End If

    'Chama a Tela CondicoesPagamentoCRLista
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 186208

    If gobjOrcamentoVenda.dtDataEmissao <> StrParaDate(DataEmissao.Text) Then

        gobjOrcamentoVenda.dtDataEmissao = StrParaDate(DataEmissao.Text)

        Call ValorTotal_Calcula

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 186208

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186209)

    End Select

    Exit Sub

End Sub

Public Sub DataReferencia_Change()

     iAlterado = REGISTRO_ALTERADO
     giDataReferenciaAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataReferencia_GotFocus()

    Dim iDataAux As Integer

    iDataAux = giDataReferenciaAlterada
    Call MaskEdBox_TrataGotFocus(DataReferencia, iAlterado)
    giDataReferenciaAlterada = iDataAux

End Sub

Public Sub DataReferencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataEmissao As Date
Dim dtDataReferencia As Date
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_DataReferencia_Validate

    If giDataReferenciaAlterada <> REGISTRO_ALTERADO Then Exit Sub

    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataReferencia.Text)
        If lErro <> SUCESSO Then gError 186210

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            If dtDataEmissao > dtDataReferencia Then gError 186211

        End If


    End If

    giDataReferenciaAlterada = 0

    'Preenche o GridParcelas
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 186212

    Exit Sub

Erro_DataReferencia_Validate:

    Cancel = True

    Select Case gErr

        Case 186210, 186212

        Case 186211
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186213)

    End Select

    Exit Sub

End Sub

Public Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataVencimento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub DataVencimento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto1Ate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto1Ate_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto1Ate_KeyPress(KeyAscii As Integer)

   Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto1Ate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto1Percentual_Change()

   iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto1Percentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto1Percentual_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto1Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto1Valor_GotFocus()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto1Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto1Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto2Ate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto2Ate_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto2Ate_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto2Ate_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto2Percentual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto2Percentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto2Percentual_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto2Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto2Valor_GotFocus()

Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto2Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto2Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto3Ate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto3Ate_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto3Ate_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto3Ate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto3Percentual_Change()

    iAlterado = REGISTRO_ALTERADO

 End Sub

Private Sub Desconto3Percentual_GotFocus()

     Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto3Percentual_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto3Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto3Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Desconto3Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Desconto3Valor_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescServicoSRV_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DescServicoSRV_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensSRV)
End Sub

Private Sub DescServicoSRV_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)
End Sub

Private Sub DescServicoSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescServicoSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    giFilialAlterada = 1

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    'Se nenhuma filial foi selecionada, sai.
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento para a filial do cliente selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 186214

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 186214

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186215)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objcliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Or giFilialAlterada = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 186216

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 186217

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 186218

        If lErro = 17660 Then

            'Lê o Cliente
            objcliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objcliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 186219

            'Se encontrou o Cliente
            If lErro = SUCESSO Then

                objFilialCliente.lCodCliente = objcliente.lCodigo

                gError 186220

            End If

        End If

        If iCodigo <> 0 Then

            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

            lErro = Trata_FilialCliente
            If lErro <> SUCESSO Then gError 186221

        Else

            objcliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0

        End If

    'Não encontrou a STRING
    ElseIf lErro = 6731 Then

        'trecho incluido por Leo em 17/04/02
        objcliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 186222

        If lErro = SUCESSO Then gError 186223

    End If

    giFilialAlterada = 0

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 186216, 186218, 186219, 186221, 186222

        Case 186217
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 186220
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 186223
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186224)

    End Select

    Exit Sub

End Sub

Function Trata_FilialCliente() As Long

Dim objFilialCliente As New ClassFilialCliente
Dim objcliente As New ClassCliente
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
    objcliente.sNomeReduzido = Trim(Cliente.Text)

    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 186225

    If lErro = 17660 Then gError 186226

    gobjOrcamentoVenda.iFilial = objFilialCliente.iCodFilial

    Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

    'Calula o valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 186227

    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 186225, 186227

        Case 186226
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186228)

    End Select

    Exit Function

End Function

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorTotal As Double
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult 'Inserido por Wagner
Dim sPeca As String
Dim iPreenchido As Integer

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iItemAtual, iGrid_Produto_Col), sPeca, iPreenchido)
    If lErro <> SUCESSO Then gError 188043

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        '************ grade ************
        'Retira a "#" caso o item excluído tenha sido um de grade
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        '*******************************

        'Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)

        Call gobjTribTab.Exclusao_Item_Grid(iItemAtual)

        Call ProdutoSRV_Remover_Item(sPeca, gcolProdSolicSRV)

'        'Calcula a soma dos valores de produtos
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'            End If
'        Next
'
'        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'            End If
'        Next
'
'        '#################################
'        'Inserido por Wagner
'        If objGridItens.iLinhasExistentes <> 0 Then
'            Call PrecoTotal_Calcula(objGridItens.iLinhasExistentes)
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
'                gobjOrcamentoVenda.dValorDesconto = 0
'
'            End If
'        End If
'        '#################################
'
'        'Coloca valor total dos produtos na tela
'        ValorProdutos.Caption = Format(dValorTotal, "Standard")

        Call SubTotal_Calcula

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186229

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 186229, 188043

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186230)

    End Select

    Exit Sub

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If


End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub


Public Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_EnterCell()

     Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Public Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

End Sub

Public Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Public Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Public Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Public Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)

End Sub
'
'Public Sub ICMSAliquotaItem_Change()
'
'    giICMSAliquotaItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ICMSAliquotaItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 183231
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
'        Case 183231
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183232)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSBase_Change()
'
'    ICMSBase1.Caption = ICMSBase.Caption
'
'End Sub
'
'Public Sub ICMSBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSBaseItemAlterado = 1
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ICMSBaseItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186233
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
'        Case 186233
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186234)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSPercRedBaseItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186235
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
'        Case 186235
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186236)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSSubstAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstAliquotaItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186237
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
'        Case 186237
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186238)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSSubstBase_Change()
'
'    ICMSSubstBase1.Caption = ICMSSubstBase.Caption
'
'End Sub
'
'Public Sub ICMSSubstBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstBaseItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186239
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
'        Case 186239
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186240)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSSubstValor_Change()
'
'    ICMSSubstValor1.Caption = ICMSSubstValor.Caption
'
'End Sub
'
'Public Sub ICMSSubstValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstValorItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSSubstValorItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186241
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
'        Case 186241
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186242)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSValor_Change()
'
'    ICMSValor1.Caption = ICMSValor.Caption
'
'End Sub
'
'Public Sub ICMSValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSValorItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSValorItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186243
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
'        Case 186243
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186244)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIAliquotaItemAlterado = 1
'
'End Sub
'
'Public Sub IPIAliquotaItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186245
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
'        Case 186245
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186246)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIBaseItemAlterado = 1
'
'End Sub
'
'Public Sub IPIBaseItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186247
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
'        Case 186247
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186248)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIPercRedBaseItemAlterado = 1
'
'End Sub
'
'Public Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186249
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
'        Case 186249
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186250)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIValor_Change()
'
'    IPIValor1.Caption = IPIValor.Caption
'
'End Sub
'
'Public Sub IPIValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIValorItemAlterado = 1
'
'End Sub
'
'Public Sub IPIValorItem_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO Then gError 186251
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
'        Case 186251
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186252)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IRAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giAliqIRAlterada = 1
'
'End Sub
'
'Public Sub IRAliquota_Validate(Cancel As Boolean)
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
'        If lErro <> SUCESSO Then gError 186253
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
'        Case 186253
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186254)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ISSAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giISSAliquotaAlterada = 1
'
'End Sub
'
'Public Sub ISSAliquota_Validate(Cancel As Boolean)
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
'        If lErro <> SUCESSO Then gError 186255
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
'        Case 186255
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186256)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ISSValor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giISSValorAlterado = 1
'
'End Sub
'
'Public Sub ISSValor_Validate(Cancel As Boolean)
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
'        If lErro <> SUCESSO Then gError 186257
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
'        Case 186257
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186258)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)


End Sub

Private Sub LabelSolicitaSRV_Click()

Dim objSolicSRV As New ClassSolicSRV
Dim colSelecao As New Collection

    objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)
    
    Call Chama_Tela("SolicitacaoSRVLista", colSelecao, objSolicSRV, objEventoCodSolicSRV)
    
End Sub

Private Sub objEventoCodSolicSRV_evSelecao(obj1 As Object)

Dim objSolicSRV As ClassSolicSRV
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodSolicSRV_evSelecao

    Set objSolicSRV = obj1
    
    CodSolicSRV.Text = objSolicSRV.lCodigo
    Call BotaoSolicitacaoSRV_Click
    
    Me.Show

    Exit Sub

Erro_objEventoCodSolicSRV_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186993)
    
    End Select

End Sub
'
'Public Sub LblTipoTrib_Click()
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
'Public Sub LblTipoTribItem_Click()
'
'    Call LblTipoTrib_Click
'
'End Sub
'
'Public Sub NaturezaItemLabel_Click()
'
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim colSelecao As New Collection
'Dim dtDataRef As Date, sSelecao As String
'
'    If Len(Trim(NaturezaOpItem.Text)) > 0 Then objNaturezaOp.sCodigo = NaturezaOpItem.Text
'
'    If Len(Trim(DataEmissao.ClipText)) > 0 Then
'        dtDataRef = MaskedParaDate(DataEmissao)
'    Else
'        dtDataRef = DATA_NULA
'    End If
'
'    sSelecao = "Codigo >= " & NATUREZA_SAIDA_COD_INICIAL & " AND Codigo <= " & NATUREZA_SAIDA_COD_FINAL & " AND {fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4")
'
'    Call Chama_Tela("NaturezaOperacaoLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, sSelecao)
'
'End Sub
'
'Public Sub NaturezaLabel_Click()
'
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim colSelecao As New Collection
'Dim dtDataRef As Date
'
'    'Se NaturezaOP estiver preenchida coloca no Obj
'    objNaturezaOp.sCodigo = gobjTribTab.sNatOpInterna
'
'    If Len(Trim(DataEmissao.ClipText)) > 0 Then
'        dtDataRef = MaskedParaDate(DataEmissao)
'    Else
'        dtDataRef = DATA_NULA
'    End If
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
'Public Sub NaturezaOp_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giNaturezaOpAlterada = 1
'
'End Sub
'
'Public Sub NaturezaOp_GotFocus()
'
'Dim iNaturezaAux As Integer
'
'    iNaturezaAux = giNaturezaOpAlterada
'    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
'    giNaturezaOpAlterada = iNaturezaAux
'
'End Sub
'
'Public Sub NaturezaOp_Validate(Cancel As Boolean)
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
'    objNaturezaOp.sCodigo = Trim(gobjTribTab.sNatOpInterna)
'
'    If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 186259
'
'    'Lê a NaturezaOp
'    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'    If lErro <> SUCESSO And lErro <> 17958 Then gError 186260
'
'    'Se não existir --> Erro
'    If lErro = 17958 Then gError 186261
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
'        Case 186259
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
'
'        Case 186260
'
'        Case 186261
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", gobjTribTab.sNatOpInterna)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
'            Else
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186262)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub NaturezaOpItem_Change()
'
'    giNatOpItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub

Public Sub NumeroLabel_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As Collection

On Error GoTo Erro_NumeroLabel_Click

    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186263

    Call Chama_Tela("OrcamentoSRV1Lista", colSelecao, objOrcamentoVenda, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 186263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186264)

    End Select

    Exit Sub

End Sub


Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

'Private Sub objEventoNaturezaOp_evSelecao(obj1 As Object)
'
'Dim objNaturezaOp As New ClassNaturezaOp
'
'    Set objNaturezaOp = obj1
'
'    If giFrameAtual = 1 Then
'
'        'Preenche a natureza de Opereração do frame principal
'        gobjTribTab.sNatOpInterna = objNaturezaOp.sCodigo
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

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objOrcamentoVenda = obj1

    lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGO)
    If lErro <> SUCESSO And lErro <> 84363 Then gError 186265

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 186265

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186266)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItensSRV.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 186267

    ServicoSRV.PromptInclude = False
    ServicoSRV.Text = sProduto
    ServicoSRV.PromptInclude = True

    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ServicoSRV.Text

    'Faz o Tratamento do produto
    lErro = Servico_Saida_Celula()
    If lErro <> SUCESSO Then


        If Not (Me.ActiveControl Is ServicoSRV) Then

            GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ""

        End If

        gError 186268
    End If

    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 186267
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 186268

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186269)

    End Select

    Exit Sub

End Sub

'Private Sub objEventoTiposDeTributacao_evSelecao(obj1 As Object)
'
'Dim objTipoTrib As ClassTipoDeTributacaoMovto
'
'    Set objTipoTrib = obj1
'
'    If giFrameAtualTributacao = 1 Then
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

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    iAlterado = 0

    Me.Show

End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de tributacao
        If Opcao.SelectedItem.Index = TAB_Tributacao Then

            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError 186270

        '??? Alteração Daniel em 29/10/2002
        ElseIf Opcao.SelectedItem.Index = TAB_Cobranca Then

            'Recalcula as parcelas
            Call CobrancaAutomatica_Click

        End If

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = Opcao.SelectedItem.Index

    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr

        Case 186270

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186271)

    End Select

    Exit Sub

End Sub

'Public Sub OpcaoTributacao_Click()
'Dim lErro As Long
'On Error GoTo Erro_OpcaoTributacao_Click
'
'    'Se frame selecionado não for o atual
'    If OpcaoTributacao.SelectedItem.Index <> giFrameAtualTributacao Then
'
'        If TabStrip_PodeTrocarTab(giFrameAtualTributacao, OpcaoTributacao, Me) <> SUCESSO Then Exit Sub
'
'        'Esconde o frame atual, mostra o novo
'        FrameTributacao(OpcaoTributacao.SelectedItem.Index).Visible = True
'        FrameTributacao(giFrameAtualTributacao).Visible = False
'        'Armazena novo valor de giFrameAtualTributacao
'        giFrameAtualTributacao = OpcaoTributacao.SelectedItem.Index
'
'        'se abriu o tab de detalhamento
'        If OpcaoTributacao.SelectedItem.Index = 2 Then
'            lErro = TributacaoItem_InicializaTab
'            If lErro <> SUCESSO Then gError 186272
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
'        Case 186272
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186273)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub PercAcrescFin_Change()

    iAlterado = REGISTRO_ALTERADO
    giPercAcresFinAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercAcrescFin_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercAcrescFin_Validate

    If giPercAcresFinAlterado = 0 Then Exit Sub

    If Len(Trim(PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
        If lErro <> SUCESSO Then gError 186274
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 186280


    End If

    giPercAcresFinAlterado = 0

    Exit Sub

Erro_PercAcrescFin_Validate:

    Cancel = True


    Select Case gErr

        Case 186274, 186275, 186280

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186276)

    End Select

    Exit Sub


End Sub

Public Sub PercentDesc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PercentDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)


End Sub

Public Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrazoValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoTotal_Change()

    iAlterado = REGISTRO_ALTERADO


End Sub

Public Sub PrecoTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PrecoTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TabelaPreco_Click()

Dim lErro As Long

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    If TabelaPreco.ListIndex = -1 Then Exit Sub

    If objGridItensSRV.iLinhasExistentes <> 0 Or objGridItens.iLinhasExistentes <> 0 Then

        'Faz o tratamento para a Tabela de Preços escolhida
        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 186278

    End If
    

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 186277, 186278

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186279)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 186281

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 186282

        If lErro <> SUCESSO Then gError 186283

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 186284

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 186286

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

    Case 186281, 186282, 186284, 186285

    Case 84014, 84015, 84017

    Case 186283  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        End If

    Case 186286
        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186287)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtualItens Then

        If TabStrip_PodeTrocarTab(giFrameAtualItens, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameItens(TabStrip1.SelectedItem.Index).Visible = True
        FrameItens(giFrameAtualItens).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtualItens = TabStrip1.SelectedItem.Index

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186791)

    End Select

    Exit Sub

End Sub

Public Sub TipoDesconto1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesconto1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoDesconto1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoDesconto1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto1
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoDesconto2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesconto2_GotFocus()

     Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoDesconto2_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoDesconto2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto2
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoDesconto3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesconto3_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoDesconto3_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoDesconto3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto3
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'Public Sub TipoTributacao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giTipoTributacaoAlterado = 1
'
'End Sub
'
'Public Sub TipoTributacao_GotFocus()
'
'Dim iTipoTributacaoAux As Integer
'
'    iTipoTributacaoAux = giTipoTributacaoAlterado
'    Call MaskEdBox_TrataGotFocus(TipoTributacao, iAlterado)
'    giTipoTributacaoAlterado = iTipoTributacaoAux
'
'End Sub
'
'Public Sub TipoTributacao_Validate(Cancel As Boolean)
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
'            If lErro <> SUCESSO And lErro <> 27259 Then gError 186282
'
'            'Se não encontrou o Tipo da Tributação --> erro
'            If lErro = 27259 Then gError 186283
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
'        Case 186282
'
'        Case 186283
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
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186284)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TipoTributacaoItem_Change()
'
'    giTipoTributacaoItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'
'End Sub
'
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
'            If lErro <> SUCESSO And lErro <> 27259 Then gError 186285
'
'            'Se não encontrou o Tipo da Tributação --> erro
'            If lErro = 27259 Then gError 186286
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
'
'    Select Case gErr
'
'        Case 186285
'
'        Case 186286
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacaoItem.Text)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("TipoDeTributacao", objTributacaoTipo)
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186287)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 186288

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 186288

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186289)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 186290

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 186290

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186291)

    End Select

    Exit Sub

End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho

   RaiseEvent Unload

End Sub

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Desconto 1 Tipo")
    objGridInt.colColuna.Add ("Desc. 1 Data")
    objGridInt.colColuna.Add ("Desc. 1 Valor")
    objGridInt.colColuna.Add ("Desc. 1 %")
    objGridInt.colColuna.Add ("Desconto 2 Tipo")
    objGridInt.colColuna.Add ("Desc. 2 Data")
    objGridInt.colColuna.Add ("Desc. 2 Valor")
    objGridInt.colColuna.Add ("Desc. 2 %")
    objGridInt.colColuna.Add ("Desconto 3 Tipo")
    objGridInt.colColuna.Add ("Desc. 3 Data")
    objGridInt.colColuna.Add ("Desc. 3 Valor")
    objGridInt.colColuna.Add ("Desc. 3 %")

    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoDesconto1.Name)
    objGridInt.colCampo.Add (Desconto1Ate.Name)
    objGridInt.colCampo.Add (Desconto1Valor.Name)
    objGridInt.colCampo.Add (Desconto1Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto2.Name)
    objGridInt.colCampo.Add (Desconto2Ate.Name)
    objGridInt.colCampo.Add (Desconto2Valor.Name)
    objGridInt.colCampo.Add (Desconto2Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto3.Name)
    objGridInt.colCampo.Add (Desconto3Ate.Name)
    objGridInt.colCampo.Add (Desconto3Valor.Name)
    objGridInt.colCampo.Add (Desconto3Percentual.Name)


    'Controles que participam do Grid
    iGrid_Vencimento_Col = 1
    iGrid_ValorParcela_Col = 2
    iGrid_Desc1Codigo_Col = 3
    iGrid_Desc1Ate_Col = 4
    iGrid_Desc1Valor_Col = 5
    iGrid_Desc1Perc_Col = 6
    iGrid_Desc2Codigo_Col = 7
    iGrid_Desc2Ate_Col = 8
    iGrid_Desc2Valor_Col = 9
    iGrid_Desc2Perc_Col = 10
    iGrid_Desc3Codigo_Col = 11
    iGrid_Desc3Ate_Col = 12
    iGrid_Desc3Valor_Col = 13
    iGrid_Desc3Perc_Col = 14

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 700

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Trata_TabelaPreco

    If Not gbCarregandoTela Then

        For iLinha = 1 To objGridItens.iLinhasExistentes

            lErro = Trata_TabelaPreco_Item(iLinha)
            If lErro <> SUCESSO Then gError 186289

        Next

        For iLinha = 1 To objGridItensSRV.iLinhasExistentes

            lErro = Trata_TabelaPreco_Item_SRV(iLinha)
            If lErro <> SUCESSO Then gError 186290

        Next

        'Calcula o Valor Total da Nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186291

    End If

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 186289, 186290, 186291

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186292)

    End Select

    Exit Function

End Function

Public Function Trata_TabelaPreco_Item(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 186293

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 186294

        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 186295

        End If

        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            'Calcula o Preco Unitário do Ítem
            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 186296
            'Coloca no Grid
            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = ""
            End If
            'Calcula o Preco Total do Ítem
            Call PrecoTotal_Calcula(iLinha)

         End If

    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 186293 To 186296

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186297)

    End Select

    Exit Function

End Function

Public Function Trata_TabelaPreco_Item_SRV(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Trata_TabelaPreco_Item_SRV

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iLinha, iGrid_ServicoSRV_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 186298

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItensSRV.TextMatrix(iLinha, iGrid_UnidadeMedSRV_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 186299

        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 186300

        End If

        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            'Calcula o Preco Unitário do Ítem
            lErro = PrecoUnitario_Calcula(GridItensSRV.TextMatrix(iLinha, iGrid_UnidadeMedSRV_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 186301
            'Coloca no Grid
            If dPrecoUnitario > 0 Then
                GridItensSRV.TextMatrix(iLinha, iGrid_PrecoUnitarioSRV_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItensSRV.TextMatrix(iLinha, iGrid_PrecoUnitarioSRV_Col) = ""
            End If
            'Calcula o Preco Total do Ítem
            Call PrecoTotal_CalculaSRV(iLinha)

         End If

    End If

    Trata_TabelaPreco_Item_SRV = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item_SRV:

    Trata_TabelaPreco_Item_SRV = gErr

    Select Case gErr

        Case 186298 To 186301

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186302)

    End Select

    Exit Function

End Function

Function ValorTotal_Calcula() As Long
'Calcula o Valor Total do Pedido

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 186303

    ValorTotal.Caption = Format(dValorTotal, "Standard")

    Call ValorDescontoTit_Validate(bSGECancelDummy)

    Call ValorDescontoItens_Calcula

    If Not gbCarregandoTela Then

        'Faz o cálculo automático das comissões
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 186305

    End If

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 186303 To 186305

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186306)

    End Select

    Exit Function

End Function

Public Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'Calcula o Preço unitário do item de acordo com a UM e a tabela de preço

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 186307
    
    If lErro = 28030 Then gError 186308

    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 186309

    dPrecoUnitario = objTabelaPrecoItem.dPreco * dFator

    'Recolhe o percentual de acréscimo financeiro
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100

    'Calcula o Preço unitário
    If dPercAcresFin <> 0 Then
        dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin)
    End If

    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 186307, 186309

        Case 186308
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186310)

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
    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorProdutos = dValorProdutos + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
        End If
    Next

    For iIndice = 1 To objGridItens.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorProdutos = dValorProdutos + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        End If
    Next

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then

        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 186311

        dValorDesconto = CDbl(ValorDesconto.Text)

        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        If dValorDesconto > dValorProdutos Then gError 186312

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    'Para tributação
    gobjOrcamentoVenda.dValorDesconto = dValorDesconto
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 186313

    giValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 186311, 186313

        Case 186312
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186314)

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
        If lErro <> SUCESSO Then gError 186315

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorOutrasDespesas = dValorDespesas

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 186316

    giValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case 186315, 186316

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186317)

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
        If lErro <> SUCESSO Then gError 186318

        dValorFrete = CDbl(ValorFrete.Text)

        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorFrete = dValorFrete

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 186319

    giValorFreteAlterado = 0

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case gErr

        Case 186318, 186319

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186320)

    End Select

    Exit Sub

End Sub

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
'        If lErro <> SUCESSO Then gError 186321
'
'        dValorIRRF = CDbl(ValorIRRF.Text)
'
'        ValorIRRF.Text = Format(dValorIRRF, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = StrParaDbl(ValorTotal.Caption)
'
'        If dValorIRRF > dValorTotal Then gError 186322
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
'        Case 186321
'
'        Case 186322
'            Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186323)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub ValorParcela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorParcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorParcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorParcela_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

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
        If lErro <> SUCESSO Then gError 186324

        dValorSeguro = CDbl(ValorSeguro.Text)

        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorSeguro = dValorSeguro

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 186325

    giValorSeguroAlterado = 0

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True

    Select Case gErr

        Case 186324, 186325

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186326)

    End Select

    Exit Sub

End Sub

Private Function Move_GridItens_Memoria(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Move Grid Itens para memória

Dim lErro As Long, iIndice As Integer
Dim sProduto1 As String, sProduto As String, iPreenchido As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        sProduto1 = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
        
        'Verifica se o Produto está preenchido
        If Len(Trim(sProduto1)) > 0 Then
    
            'Formata o produto
            lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 186329
            
        End If
        
        lErro = Move_GridItem_Memoria(objOrcamentoVenda, iIndice, sProduto)
        If lErro <> SUCESSO Then gError 186329

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 186329

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186330)

    End Select

    Exit Function

End Function

Public Function Move_GridItem_Memoria(objOrcamentoVenda As ClassOrcamentoVenda, iIndice As Integer, sProduto1 As String) As Long
'Recolhe do Grid os dados do item orçamento no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemOrcamento As New ClassItemOV, objTributacaoOV As New ClassTributacaoDoc
Dim iPreenchido As Integer

On Error GoTo Erro_Move_GridItem_Memoria

    Set objItemOrcamento = New ClassItemOV

'    'Verifica se o Produto está preenchido
'    If Len(Trim(sProduto1)) > 0 Then
'
'        'Formata o produto
'        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 186331

    objItemOrcamento.sProduto = sProduto1
'    End If

    If Len(Trim(sProduto1)) = 0 Then gError 186332

    'Armazena os dados do item
    objItemOrcamento.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    objItemOrcamento.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
    objItemOrcamento.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
    objItemOrcamento.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    objItemOrcamento.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objItemOrcamento.sDescricao = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
    objItemOrcamento.iFilialEmpresa = giFilialEmpresa
    objItemOrcamento.dtDataEntrega = DATA_NULA
    
    If gobjOrcamentoVenda.colItens.Count >= iIndice Then
        Set objItemOrcamento.objTributacaoItemOV = gobjOrcamentoVenda.colItens.Item(iIndice).objTributacaoItemOV
    Else
        Set objItemOrcamento.objTributacaoItemOV = Nothing
    End If

    objItemOrcamento.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)

    'Adiciona o item na colecao de itens do orçamento de venda
     objOrcamentoVenda.colItens.Add objItemOrcamento

    Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItem_Memoria:

    Move_GridItem_Memoria = gErr

    Select Case gErr

        Case 186331

        Case 186332
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186333)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 186337

            'Se for o GridItens
            Case GridItensSRV.Name

                lErro = Saida_Celula_GridItensSRV(objGridInt)
                If lErro <> SUCESSO Then gError 186338

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 186339

            'se for o GridComissoes
            Case BLGGridBloqueio.Name

                lErro = gobjTelaTrataBloqGen.Saida_Celula(objGridInt)
                If lErro <> SUCESSO Then gError 198291

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 186340

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 186337 To 186340, 198291

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186341)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long
'Faz a crítica do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vencimento_Col
            lErro = Saida_Celula_DataVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 186342

        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 186343

        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 186344

        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError 186345

        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError 186346

        Case iGrid_Desc1Perc_Col, iGrid_Desc2Perc_Col, iGrid_Desc3Perc_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError 186347

    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 186342 To 186347

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186348)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Servico
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 186349

        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescProduto(objGridInt)
            If lErro <> SUCESSO Then gError 186350

        'Se for a de Unidade de Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 186351

        'Se for a de Quantidade Pedida
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 186352

        'Se for a de Preço Unitário
        Case iGrid_PrecoUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 186353

        'Se for a de Percentual de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 186354

        'Inserido por Wagner
        Case iGrid_Observacao_Col
            lErro = Saida_Celula_Observacao(objGridInt)
            If lErro <> SUCESSO Then gError 186355

        Case iGrid_VersaoKit_Col
            lErro = Saida_Celula_VersaoKit(objGridInt)
            If lErro <> SUCESSO Then gError 186371

        Case iGrid_VersaoKitBase_Col
            lErro = Saida_Celula_VersaoKitBase(objGridInt)
            If lErro <> SUCESSO Then gError 186372
            

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 186349 To 186357, 186388, 186389

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186358)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim dValorTotal As Double
Dim dValor As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 186377
    If Len(Trim(Cliente.Text)) = 0 Then gError 186378
    If Len(Trim(Filial.Text)) = 0 Then gError 186379
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 186380
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 186381

    dValor = CDbl(ValorTotal.Caption)

    If dValor < 0 Then gError 186382

    lErro = Valida_Grid_ItensSRV()
    If lErro <> SUCESSO Then gError 186390

    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 186391

    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError 186392

    'testa se os totais de servicos e pecas contidos em gcolProdSolicSRV esta de acordo com os grids de servicos e pecas
    lErro = Valida_Totais_Servicos_Pecas()
    If lErro <> SUCESSO Then gError 188234
    
    lErro = gobjTelaTrataBloqGen.Valida_Grid_Bloqueio()
    If lErro <> SUCESSO Then gError 188234

    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186399

'    lErro = Valida_Tributacao_Gravacao()
'    If lErro <> SUCESSO Then gError 186400

    lErro = gobjTribTab.Valida_Dados()
    If lErro <> SUCESSO Then gError 181786

    '###############################################
    'Inserido por Wagner 10/08/2006
'    lErro = gobjTelaProjetoInfo.Valida_Dados
'    If lErro <> SUCESSO Then gError 186401
    '###############################################
    
    'Grava no BD
    lErro = CF("OrcamentoSRV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186402

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then

        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then

            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_ORCSRV
            gobjAnotacao.sID = CStr(objOrcamentoVenda.iFilialEmpresa) & "," & CStr(objOrcamentoVenda.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje

            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 186403

        End If

    End If
    '########################################

    lErro = gobjTelaTrataBloqGen.Trata_Bloqueios_Gravacao(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186404
    
    'Incluído por Luiz Nogueira em 04/06/03
    'Se for para imprimir o orçamento depois da gravação
    If ImprimeOrcamentoGravacao.Value = vbChecked Then

        'Dispara função para imprimir orçamento
        lErro = Orcamento_Imprime(Trim(objOrcamentoVenda.lCodigo))
        If lErro <> SUCESSO Then gError 186404

    End If

    '########################################################
    'INSERIDO POR WAGNER
    'Se for para enviar por email o orçamento depois da gravação
    If EmailOrcamentoGravacao.Value = vbChecked Then

        'Dispara função para enviar por email orçamento
        lErro = Orcamento_Email(objOrcamentoVenda.lCodigo)
        If lErro <> SUCESSO And lErro <> 102237 Then gError 186405

    End If
    '########################################################

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 186377
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 186378
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 186379
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 186380
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 186381
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 186382
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_NEGATIVO", gErr)


        Case 186390 To 186392, 186399, 186400, 186401 To 186405, 188234

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186406)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iLinha, iGrid_ServicoSRV_Col), sServicoFormatado, iServicoPreenchido)
    If lErro <> SUCESSO Then gError 186404

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 186405

    Select Case objControl.Name

        Case ServicoSRV.Name
            'Se o produto estiver preenchido desabilita
            If iServicoPreenchido <> PRODUTO_VAZIO Then
                ServicoSRV.Enabled = False
            Else
                ServicoSRV.Enabled = True
            End If

        Case UnidadeMedSRV.Name
            'guarda a um go grid nessa coluna
            sUM = GridItensSRV.TextMatrix(iLinha, iGrid_UnidadeMedSRV_Col)

            UnidadeMedSRV.Enabled = True

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UnidadeMedSRV.Text

            UnidadeMedSRV.Clear

            If iServicoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sServicoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 186406

                If lErro = 28030 Then gError 186407

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 186408
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UnidadeMedSRV.AddItem objUM.sSigla
                Next

                'Tento selecionar na Combo a Unidade anterior
                If UnidadeMedSRV.ListCount <> 0 Then

                    For iIndice = 0 To UnidadeMedSRV.ListCount - 1

                        If UnidadeMedSRV.List(iIndice) = sUnidadeMed Then
                            UnidadeMedSRV.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If

            Else
                UnidadeMedSRV.Enabled = False
            End If

        Case QuantidadeSRV.Name, PrecoUnitarioSRV.Name, PercentDescSRV.Name, DataEntregaSRV.Name, MotivoPerdaItemSRV.Name, StatusItemSRV.Name, ObservacaoSRV.Name, VersaoKitSRV.Name, VersaoKitBaseSRV.Name
             
            'Se o produto estiver preenchido, habilita o controle
            If iServicoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Produto.Name
            
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If

        Case VersaoKit.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = True
            Else
                Produto.Enabled = False
            End If

        Case VersaoKitBase.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = True
            Else
                Produto.Enabled = False
            End If

        Case UnidadeMed.Name
            
            'guarda a um go grid nessa coluna
            sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

            UnidadeMed.Enabled = True

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UnidadeMed.Text

            UnidadeMed.Clear

            If iProdutoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sProdutoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 186409

                If lErro = 28030 Then gError 186410

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 186411
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next

                'Tento selecionar na Combo a Unidade anterior
                If UnidadeMed.ListCount <> 0 Then

                    For iIndice = 0 To UnidadeMed.ListCount - 1

                        If UnidadeMed.List(iIndice) = sUnidadeMed Then
                            UnidadeMed.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If

            Else
                UnidadeMed.Enabled = False
            End If

        Case PrecoUnitario.Name, PercentDesc.Name, Observacao.Name, _
             VersaoKit.Name, VersaoKitBase.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        '******** O TRATAMENTO DE QUANTIDADE FOI DESTACADO PARA TRATAR GRADE
        Case Quantidade.Name

            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If


        Case Desconto1Ate.Name, Desconto1Valor.Name, Desconto1Percentual.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))
                If objControl.Name = Desconto1Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto1Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto1Valor.Enabled = True
                ElseIf objControl.Name = Desconto1Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto1Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If

        Case Desconto2Ate.Name, Desconto2Valor.Name, Desconto2Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = Desconto2Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto2Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto2Valor.Enabled = True
                ElseIf objControl.Name = Desconto2Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto2Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If

        Case Desconto3Ate.Name, Desconto3Valor.Name, Desconto3Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = Desconto3Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto3Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto3Valor.Enabled = True
                ElseIf objControl.Name = Desconto3Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto3Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If

        Case ValorParcela.Name
            'Se o vencimento estiver preenchido, habilita o controle
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case TipoDesconto2.Name, TipoDesconto3.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select
    
    lErro = gobjTelaTrataBloqGen.Rotina_Grid_Enabled(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 198437

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 186404 To 186406, 186408, 186409, 186411, 198437

        Case 186407, 186410
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186412)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto Data que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 186416
        
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186417

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 186416, 186417
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186418)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidadede Medida que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UnidadeMed

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186419

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 186419
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186420)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_DescProduto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do Produto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescProduto

    Set objGridInt.objControle = DescricaoProduto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186425

    Saida_Celula_DescProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_DescProduto:

    Saida_Celula_DescProduto = gErr

    Select Case gErr

        Case 186425
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186426)

    End Select

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade que está deixando de ser a corrente

Dim lErro As Long
Dim bQuantidadeIgual As Boolean
Dim iIndice As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    bQuantidadeIgual = False

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 186431

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    If dQuantAnterior = StrParaDbl(Quantidade.Text) Then bQuantidadeIgual = True

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186432

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 And Not bQuantidadeIgual Then
    
        Call PrecoTotal_Calcula(GridItens.Row)
    
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186433

    End If

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 186431 To 186433
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186434)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
        If lErro <> SUCESSO Then gError 186435

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)

    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186436

    If Not bPrecoUnitarioIgual Then

        Call PrecoTotal_Calcula(GridItens.Row)

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186437
    
    End If
   
    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr


    Select Case gErr

        Case 186435 To 186437
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186438)

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
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    dPercentDesc = StrParaDbl(PercentDesc.Text)
    
    dPercentDescAnterior = PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)) * 100

    If Len(PercentDesc.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(PercentDesc.Text) 'Alterado por Wagner
        If lErro <> SUCESSO Then gError 186448

        If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 186449

            PercentDesc.Text = Format(dPercentDesc, "Fixed")

        End If

    Else

        dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
        dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    sValorPercAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)

    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186450
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then

        'Recalcula o preço total
        Call PrecoTotal_Calcula(GridItens.Row)

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186451

        'Preenche GridParcelas a partir da Condição de Pagto
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 186452

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 186448, 186450, 186451
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186449
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186452

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186453)

    End Select

    Exit Function

End Function

Function Produto_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult
Dim objItemOV As ClassItemOV
Dim iPossuiGrade As Integer
Dim objRomaneioGrade As ClassRomaneioGrade
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProdutoPai As String
Dim objGridItens1 As Object
Dim objKit As New ClassKit 'Inserido por Wagner
Dim sCodigoAntigo As String
Dim sCliente As String
Dim iFilialCli As Integer

On Error GoTo Erro_Produto_Saida_Celula

'***********  FUNÇÃO ALTERADA PARA TRATAMENTO DE GRADE  ******************

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 186472

    If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
        gError 186473
    End If

    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 186474

    If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then gError 188176

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 186475

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

    End If

    'Verifica se é um kit
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then

        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then

            objKit.sProdutoRaiz = objProduto.sCodigo
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 186476

            'Se encontrou => É UM KIT
            If lErro <> 106304 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_VersaoKit_Col) = Codigo.Text & SEPARADOR & CStr(GridItens.Row)
                GridItens.TextMatrix(GridItens.Row, iGrid_VersaoKitBase_Col) = objKit.sVersao
            End If

        End If
    End If

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao


    'Preço Unitário
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 186477

        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 186478

        End If

        If lErro <> 28014 Then

            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 186479

            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = ""
            End If

        End If

    End If

    'precodesc
    Set objGridItens1 = GridItens
    sCliente = Cliente.Text
    iFilialCli = Codigo_Extrai(Filial.Text)
    Call CF("Produto_Saida_Celula_PrecoDesc", objGridItens1, GridItens.Row, iGrid_PrecoUnitario_Col + 1, dPrecoUnitario, objProduto.sCodigo, iGrid_DescProduto_Col, sCliente, iFilialCli, gdDesconto, iGrid_PrecoUnitario_Col + 2)

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

'        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 186480

    End If

    'Atualiza a checkbox do grid para exibir a figura marcada/desmarcada
    Call Grid_Refresh_Checkbox(objGridItens)

    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 186472, 186473, 186476 To 186480

        Case 186474
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If

        Case 186475
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)

        Case 188176
            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_NAO_PODE_TER_NATUREZA_SERVICO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186481)

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
        dPercentDesc = CDbl(Format(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade)

    'Se percentual for >0 tira o desconto
    'If dPercentDesc > 0 Then dDesconto = dPercentDesc * dPrecoTotal
    dDesconto = dPercentDesc * dPrecoTotal
    dPrecoTotalReal = dPrecoTotal - dDesconto 'Inserido por Wagner

    'precodesc
    Set objGridItens1 = GridItens
    Call CF("PrecoTotal_Calcula_PrecoDesc", objGridItens1, iLinha, iGrid_PrecoUnitario_Col + 1, Format(dPrecoUnitario * (1 - dPercentDesc), "Standard"))

    'Coloca valor do desconto no Grid
'    If dDesconto > 0 Then
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
'    Else
'        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
'    End If

    'Coloca preco total do ítem no grid
    GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")

    GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dPrecoTotal, "Standard")

    'Calcula a soma dos valores de produtos
    dValorTotal = 0
    
'    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'        End If
'    Next
'
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'        End If
'    Next
'
'    If gdDesconto > 0 Then
'        dValorDesconto = gdDesconto * dValorTotal
'    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
'        dValorDesconto = CDbl(ValorDesconto.Text)
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
'        'Para tributação
'        gobjOrcamentoVenda.dValorDesconto = dValorDesconto
'
'        'Faz a atualização dos valores da tributação
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 186489
'
'        'Calcula a soma dos valores de produtos
'        dValorTotal = 0
'
'        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'            End If
'        Next
'
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'            End If
'        Next
'
'    End If
'
'    'Coloca valor total dos produtos na tela
'    ValorProdutos.Caption = Format(dValorTotal, "Standard")
'    ValorDesconto.Text = Format(dValorDesconto, "Standard")

    Call SubTotal_Calcula

    lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
    If lErro <> SUCESSO Then gError 186451
        
    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr

        Case 186489

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186490)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Vencimento que está deixando de serr a corrente

Dim lErro As Long
Dim dtDataReferencia As Date
Dim dtDataVencimento As Date
Dim sDataVencimento As String
Dim iCriouLinha As Boolean

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 186491

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = CDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 186492
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")

        iCriouLinha = False
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            iCriouLinha = True
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186493

    If iCriouLinha Then
        'Coloca desconto padrao (le em CPRConfig)
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 186494
    End If

    Saida_Celula_DataVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr

    Select Case gErr

        Case 186491, 186493, 186494
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186492
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186495)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorParcela(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Parcela que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim iIndice As Integer
Dim iColDescPerc As Integer
Dim iColTipoDesconto As Integer
Dim lTamanho As Long
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String
Dim iTipoDesconto As Integer

On Error GoTo Erro_Saida_Celula_ValorParcela

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 186496

        ValorParcela.Text = Format(ValorParcela.Text, "Standard")

        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then

            CobrancaAutomatica.Value = vbUnchecked

            '***Código para colocar valores de desconto
            dValorParcela = StrParaDbl(ValorParcela.Text)
            If dValorParcela > 0 Then

                'Vai varrer todos os 3 descontos para colocar valores
                For iIndice = 1 To 3

                    Select Case iIndice
                        Case 1
                            iColDescPerc = iGrid_Desc1Perc_Col
                            iColTipoDesconto = iGrid_Desc1Codigo_Col
                        Case 2
                            iColDescPerc = iGrid_Desc2Perc_Col
                            iColTipoDesconto = iGrid_Desc2Codigo_Col
                        Case 3
                            iColDescPerc = iGrid_Desc3Perc_Col
                            iColTipoDesconto = iGrid_Desc3Codigo_Col
                    End Select

                    iTipoDesconto = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iColTipoDesconto))
                    lTamanho = Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc)))

                    'Coloca valor de desconto na tela
                    If (iTipoDesconto = Percentual Or iTipoDesconto = PERC_ANT_DIA Or iTipoDesconto = PERC_ANT_DIA_UTIL) And lTamanho > 0 Then
                        dPercentual = PercentParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc))
                        sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
                        GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc - 1) = sValorDesconto
                    End If

                Next

            End If
            '***Fim Código para colocar valores de desconto

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            'Coloca DescontoPadrao
            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
            If lErro <> SUCESSO Then gError 186497

        End If

    Else

        '***Código para colocar valores de desconto
        'Limpa Valores de Desconto
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc1Valor_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc2Valor_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc3Valor_Col) = ""
        '***Fim Código para colocar valores de desconto

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186498

    Saida_Celula_ValorParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr

    Select Case gErr

        Case 186496 To 196498
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186499)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Desconto que está deixando de serr a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iTipo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_TipoDesconto

    If GridParcelas.Col = iGrid_Desc1Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto1
    ElseIf GridParcelas.Col = iGrid_Desc2Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto2
    ElseIf GridParcelas.Col = iGrid_Desc3Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto3
    End If

    'Verifica se o Tipo foi preenchido
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If objGridInt.objControle.Text <> objGridInt.objControle.List(objGridInt.objControle.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 186500

            'Não foi encontrado
            If lErro = 25085 Then gError 186501
            If lErro = 25086 Then gError 186502

        End If

        'Extrai o tipo de Desconto
        iTipo = Codigo_Extrai(objGridInt.objControle.Text)

        If (iTipo = VALOR_ANT_DIA) Or (iTipo = VALOR_ANT_DIA_UTIL) Or (iTipo = VALOR_FIXO) Then
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3) = ""
        ElseIf iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual Then
            '*** Acrescentado + 1 If para contabilizar com colocação de valores de desconto
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3))) = 0 Then
                GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 2) = ""
            End If
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    Else

        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Perc_Col
            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
        Next

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186503

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 186500, 186503
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186501
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186502
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186504)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoData(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Data que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_DescontoData

    If GridParcelas.Col = iGrid_Desc1Ate_Col Then
        Set objGridInt.objControle = Desconto1Ate
    ElseIf GridParcelas.Col = iGrid_Desc2Ate_Col Then
        Set objGridInt.objControle = Desconto2Ate
    ElseIf GridParcelas.Col = iGrid_Desc3Ate_Col Then
        Set objGridInt.objControle = Desconto3Ate
    End If

    If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then

        lErro = Data_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 186505
        'Se a data de vencimento estiver preenchida
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)) < CDate(objGridInt.objControle.Text) Then gError 186506
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186507

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 186505, 186507
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186506
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADESCONTO_MAIOR_DATAVENCIMENTO", gErr, CDate(objGridInt.objControle.Text), CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186508)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Valor que está deixando de serr a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_DescontoValor

    If GridParcelas.Col = iGrid_Desc1Valor_Col Then
        Set objGridInt.objControle = Desconto1Valor
    ElseIf GridParcelas.Col = iGrid_Desc2Valor_Col Then
        Set objGridInt.objControle = Desconto2Valor
    ElseIf GridParcelas.Col = iGrid_Desc3Valor_Col Then
        Set objGridInt.objControle = Desconto3Valor
    End If

    'Verifica se valor está preenchido
    If Len(objGridInt.objControle.ClipText) > 0 Then
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 186509

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186510

    Saida_Celula_DescontoValor = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr

    Select Case gErr

        Case 186509, 186510
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186511)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoPerc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Percentual que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Saida_Celula_DescontoPerc

    If GridParcelas.Col = iGrid_Desc1Perc_Col Then
        Set objGridInt.objControle = Desconto1Percentual
    ElseIf GridParcelas.Col = iGrid_Desc2Perc_Col Then
        Set objGridInt.objControle = Desconto2Percentual
    ElseIf GridParcelas.Col = iGrid_Desc3Perc_Col Then
        Set objGridInt.objControle = Desconto3Percentual
    End If

    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Critica porcentagem
        lErro = Porcentagem_Critica_Negativa(objGridInt.objControle.Text) ' Alterado por Wagner
        If lErro <> SUCESSO Then gError 186512

        '***Código para colocar valores de desconto
        dPercentual = CDbl(objGridInt.objControle.Text) / 100
        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col))

        'Coloca Valor do Desconto na tela
        If dValorParcela > 0 Then
            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = sValorDesconto
        End If

    Else

        'Limpa Valor de Desconto
        GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = ""
        '***Fim Código para colocar valores de desconto

    End If


    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186513

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 186512, 186513
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186514)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_Itens

    'para cada item do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then gError 186515

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 186516

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        If lErro <> SUCESSO Then gError 186517

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))) = 0 Then gError 186518

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
        If lErro <> SUCESSO Then gError 186519

        '##################################################
        'INSERIDO POR WAGNER
        If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col))) = 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col))) <> 0 Then gError 186520
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col))) = 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col))) <> 0 Then gError 186521
        End If
        '##################################################


    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr

        Case 186515
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 186516
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 186517, 186519

        Case 186518
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 186520
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAOKIT_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 186521
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAOKITBASE_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186522)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Parcelas() As Long
'Valida os dados do Grid de Parcelas

Dim lErro As Long
Dim iIndice As Integer
Dim dSomaParcelas As Double
'Dim dValorIRRF As Double, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
'Dim dValorTotal As Double
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim iTamanho As Integer
Dim iTipo As Integer
Dim dPercAcrecFin As Double
Dim iDesconto As Integer
Dim dtDataDesconto As Date
Dim dValorPagar As Double

On Error GoTo Erro_Valida_Grid_Parcelas

    'Verifica se alguma parcela foi informada
    If objGridParcelas.iLinhasExistentes = 0 Then gError 186523

    dSomaParcelas = 0

    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 186524
        
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 186525

        'Se o tipo de desconto 1 estiver preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data de desconto está preenchdida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 186526
            'Recolhe o Valor ou Percentual de desconto
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 186527
            Else
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))) = 0 Then gError 186528
            End If
            'Se o tipo de desconto 2 estiver preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchdida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 186529
                'Faz a crítica da ordem das datas de desconto
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 186530
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 186531
                'Recolhe o Valor ou Percentual de desconto
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 186532
                Else
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))) = 0 Then gError 186533
                End If
                'Se o tipo de desconto 3 estiver preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchdida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 186534
                    'Faz a crítica da ordem das datas de desconto
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 186535
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 186536
                    'Recolhe o Valor ou Percentual de desconto
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 186537
                    Else
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))) = 0 Then gError 186538
                    End If
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 186539
                Else
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 186540
                End If
            Else
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 186541
            End If
        End If



        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 186542

        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

    Next

'    dValorTotal = StrParaDbl(ValorTotal.Caption)
'    dValorIRRF = StrParaDbl(ValorIRRF.Text)
'    If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
'    If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
'    If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
'
'    If Format((dValorTotal - (dValorIRRF + dPISRetido + dCOFINSRetido + dCSLLRetido)), "Standard") <> Format(dSomaParcelas, "Standard") Then gError 186543

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 193062
    
    If Abs(dValorPagar - dSomaParcelas) > DELTA_VALORMONETARIO Then gError 186543

    Valida_Grid_Parcelas = SUCESSO

    Exit Function

Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr

    Select Case gErr

        Case 186523
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)

        Case 186524
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)

        Case 186525
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)

        Case 186526, 186529, 186534
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)

        Case 84222
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
        
        Case 186527, 186528, 186532, 186533, 186536, 186538
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)

        Case 186530, 186535
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 186531, 186536
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 186539, 186540, 186541
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_PARCELA_SUP_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 186542
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)

        Case 186543
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)
            
        Case 193062

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186544)

    End Select

    Exit Function

End Function

Private Function Move_OrcamentoVenda_Memoria(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Move os dados da tela para objOrcamentoVenda

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim dValorTotalParcelas As Double
Dim dValorIRRF As Double
Dim objVendedor As New ClassVendedor
Dim objProjetoInfo As Object
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_OrcamentoVenda_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objOrcamentoVenda.lCodigo = StrParaLong(Codigo.Text)

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objcliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 186392

        If lErro = SUCESSO Then
            'Guarda código do Cliente em objOrcamentoVenda
            objOrcamentoVenda.lCliente = objcliente.lCodigo
        End If

        objOrcamentoVenda.sNomeCli = objcliente.sNomeReduzido

    End If

    'Verifica se vendedor existe
    If Len(Trim(Vendedor.Text)) > 0 Then

        objVendedor.sNomeReduzido = Trim(Vendedor.Text)

        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 186393

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 186394

        objOrcamentoVenda.iVendedor = objVendedor.iCodigo

    End If

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then

        'Se o Cliente estiver cadastrado
        If objOrcamentoVenda.lCliente <> 0 Then

            'a filial tb deverá estar cadastrada e por isso teremos o código da filial na tela
            objOrcamentoVenda.iFilial = Codigo_Extrai(Filial.Text)
            objOrcamentoVenda.sNomeFilialCli = Nome_Extrai(Filial.Text)

        Else
            'se não, guardaremos o Texto digitado pelo usuário
            objOrcamentoVenda.sNomeFilialCli = Trim(Filial.Text)

        End If

    End If


    If Len(Trim(CodSolicSRV.Text)) > 0 Then
    
        objSolicSRV.iFilialEmpresa = giFilialEmpresa
        objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)
        
        lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 183258 Then gError 186966

        If lErro <> SUCESSO Then gError 186967
        
        objOrcamentoVenda.lNumIntSolicSRV = objSolicSRV.lNumIntDoc
        
    End If

    'Preenche objOrcamentoVenda com dados da tela
    objOrcamentoVenda.dtDataEmissao = MaskedParaDate(DataEmissao)
    objOrcamentoVenda.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    objOrcamentoVenda.sNaturezaOp = Trim(gobjTribTab.sNatOpInterna)
    objOrcamentoVenda.dValorFrete = StrParaDbl(ValorFrete.Text)
    objOrcamentoVenda.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objOrcamentoVenda.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objOrcamentoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objOrcamentoVenda.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objOrcamentoVenda.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    objOrcamentoVenda.iPrazoValidade = StrParaInt(PrazoValidade.Text)
    objOrcamentoVenda.dPercAcrescFinanceiro = StrParaDbl(PercAcrescFin.ClipText)
    objOrcamentoVenda.dtDataReferencia = MaskedParaDate(DataReferencia)
    objOrcamentoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.ClipText)
    objOrcamentoVenda.iCobrancaAutomatica = StrParaInt(CobrancaAutomatica.Value)
    objOrcamentoVenda.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    Set objOrcamentoVenda.colProdSolicSRV = gcolProdSolicSRV

    '########################################
    'Inserido por Wagner
    objOrcamentoVenda.sUsuario = gsUsuario
    objOrcamentoVenda.lCodigoBase = StrParaLong(CodigoBase.Text)
    '#######################################

    'Move Grid Itens para memória
    lErro = Move_GridItensSRV_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186395

    'Move Grid Itens para memória
    lErro = Move_GridItens_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186396

    'Move Tab Cobrança para memória
    Call Move_TabCobranca_Memoria(objOrcamentoVenda)

    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    
    lErro = gobjTelaTrataBloqGen.Move_GridBloqueio_Memoria(objOrcamentoVenda.colBloqueiosPV)
    If lErro <> SUCESSO Then gError 186396

    'Move Tributacao para objOrcamentoVenda
    Set objOrcamentoVenda.objTributacaoOV = gobjOrcamentoVenda.objTributacaoOV

    '#########################################
    'Inserido por Wagner 09/12/05
    If Status.ListIndex <> -1 Then
        objOrcamentoVenda.lStatus = Status.ItemData(Status.ListIndex)
    End If
    If MotivoPerda.ListIndex <> -1 Then
        objOrcamentoVenda.lMotivoPerda = MotivoPerda.ItemData(MotivoPerda.ListIndex)
    End If
    '#########################################

    '##################################################
    'Inserido por Wagner 09/08/2006
'    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_ORCSRV)
'    If lErro <> SUCESSO Then gError 186397
'
'    Set objOrcamentoVenda.objProjetoInfo = objProjetoInfo
    '##################################################

    Set objOrcamentoVenda.objTela = Me
    
    Set objOrcamentoVenda.objInfoAdic = gobjInfoAdic
    objOrcamentoVenda.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objOrcamentoVenda.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    Move_OrcamentoVenda_Memoria = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Memoria:

    Move_OrcamentoVenda_Memoria = gErr

    Select Case gErr

        Case 186392, 186393, 186395, 186396, 186397, 186966

        Case 186394
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)

        Case 186967
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186398)

    End Select

    Exit Function

End Function

Private Sub Move_TabCobranca_Memoria(objOrcamentoVenda As ClassOrcamentoVenda)
'Recolhe os dados do tab de cobrança

Dim lTamanho As Long
Dim iIndice As Integer
Dim objParcelaOV As ClassParcelaOV

    'Recolhe os dados da Cobrança
    objOrcamentoVenda.iCobrancaAutomatica = CobrancaAutomatica.Value
    objOrcamentoVenda.dtDataReferencia = MaskedParaDate(DataReferencia)
    objOrcamentoVenda.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    objOrcamentoVenda.dPercAcrescFinanceiro = StrParaDbl(PercAcrescFin.Text) / 100

    If objGridParcelas.iLinhasExistentes = 0 Then Exit Sub

    'Recolhe os Dados do Grid de Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaOV = New ClassParcelaOV

        objParcelaOV.iNumParcela = iIndice

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) > 0 Then
            objParcelaOV.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        Else
            objParcelaOV.dtDataVencimento = DATA_NULA
        End If

        objParcelaOV.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        objParcelaOV.iDesconto1Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
        objParcelaOV.iDesconto2Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
        objParcelaOV.iDesconto3Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto1Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
        Else
            objParcelaOV.dtDesconto1Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto2Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
        Else
            objParcelaOV.dtDesconto2Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto3Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
        Else
            objParcelaOV.dtDesconto3Ate = DATA_NULA
        End If

        If objParcelaOV.iDesconto1Codigo = VALOR_FIXO Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto1Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))
        ElseIf objParcelaOV.iDesconto1Codigo = Percentual Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto1Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))
        End If

        If objParcelaOV.iDesconto2Codigo = VALOR_FIXO Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto2Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))
        ElseIf objParcelaOV.iDesconto2Codigo = Percentual Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto2Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))
        End If

        If objParcelaOV.iDesconto3Codigo = VALOR_FIXO Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto3Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))
        ElseIf objParcelaOV.iDesconto3Codigo = Percentual Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto3Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))
        End If

        objOrcamentoVenda.colParcela.Add objParcelaOV

    Next
    
    objOrcamentoVenda.dValorDescontoTit = StrParaDbl(ValorDescontoTit.Text)

End Sub
'
'Public Sub NaturezaOpItem_Validate(Cancel As Boolean)
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
'        If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 186524
'
'        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'        If lErro <> SUCESSO And lErro <> 17958 Then gError 186525
'
'        'Se não achou a Natureza de Operação --> erro
'        If lErro <> SUCESSO Then gError 186526
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
'        Case 186524
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
'
'        Case 186525
'
'        Case 186526
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOpItem.Text)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186527)
'
'    End Select
'
'End Sub
'
'Public Sub NaturezaOpItem_GotFocus()
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
'Public Sub TributacaoRecalcular_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_TributacaoRecalcular_Click
'
'    giRecalculandoTributacao = 1
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 186528
'
'    giRecalculandoTributacao = 0
'
'    Exit Sub
'
'Erro_TributacaoRecalcular_Click:
'
'    Select Case gErr
'
'        Case 186528
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186529)
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

'Public Sub ISSIncluso_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_ISSIncluso_Click
'
'    iAlterado = REGISTRO_ALTERADO
'
'    Call BotaoGravarTrib
'
'    Exit Sub
'
'Erro_ISSIncluso_Click:
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186530)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function Traz_OrcamentoVenda_Tela(objOrcamentoVenda As ClassOrcamentoVenda, ByVal iBrowse As Integer) As Long
'Coloca na tela os dados do Orcamento de Venda

Dim lErro As Long
Dim objFilial As AdmFiliais
Dim bCancel As Boolean
Dim lCodigo As Long
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_Traz_OrcamentoVenda_Tela

    gbCarregandoTela = True

    lCodigo = StrParaLong(Codigo.Text)

    Call Limpa_OrcamentoVenda2

    'Lê os dados do Orcamento de Venda
    lErro = CF("OrcamentoSRV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186534 Then gError 186628
    
    If lErro <> SUCESSO Then gError 186629

    'Lê a parte de Tributação
    lErro = CF("OrcamentoSRV_Le_Tributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186623 Then gError 186630
    
    If lErro <> SUCESSO Then gError 186631

    lErro = CF("OrcamentoSRV_Le_Itens_ComTributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186602 Then gError 186632
    
    If lErro <> SUCESSO Then gError 186633

    Set gcolProdSolicSRV = objOrcamentoVenda.colProdSolicSRV

    lErro = CF("ParcelasOrcSRV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186617 Then gError 186634
    
    If lErro <> SUCESSO Then gError 186635

    lErro = gobjTribTab.Traz_NFiscal_Tela(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186636
    
    objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OSRV
    objOrcamentoVenda.objInfoAdic.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    lErro = CF("InfoAdicionais_Le", objOrcamentoVenda.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 186636

    ValorTotal.Caption = Format(objOrcamentoVenda.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objOrcamentoVenda.dValorProdutos, "Standard")

    PercAcrescFin.Text = ""
    'Coloca os dados do Orcamento na tela

    '#######################################
    'ALTERADO POR WAGNER
    Usuario.Caption = objOrcamentoVenda.sUsuario

    If iBrowse = VALOR_CODIGO Then
        Codigo.Text = objOrcamentoVenda.lCodigo
        If objOrcamentoVenda.lCodigoBase <> 0 Then CodigoBase.Text = objOrcamentoVenda.lCodigoBase
    Else
        CodigoBase.Text = objOrcamentoVenda.lCodigo
        If lCodigo <> 0 Then Codigo.Text = lCodigo
    End If
    '#######################################

    'Se existe um código para o Cliente
    If objOrcamentoVenda.lCliente <> 0 Then

        Call Cliente_Formata(objOrcamentoVenda.lCliente)
        Call Filial_Formata(Filial, objOrcamentoVenda.iFilial)

    Else

        'Preenche o Cliente e a Filial com os Nomes Informados.
        Cliente.Text = objOrcamentoVenda.sNomeCli
        Filial.Text = objOrcamentoVenda.sNomeFilialCli

    End If

    giFilialAlterada = 0

'    NatOpEspelho.Caption = objOrcamentoVenda.sNaturezaOp

    ValorFrete.Text = Format(objOrcamentoVenda.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objOrcamentoVenda.dValorSeguro, "Standard")
    ValorDesconto.Text = Format(objOrcamentoVenda.dValorDesconto, "Standard")
    ValorDespesas.Text = Format(objOrcamentoVenda.dValorOutrasDespesas, "Standard")

'    gobjTribTab.sNatOpInterna = objOrcamentoVenda.sNaturezaOp
'    Call NaturezaOp_Validate(bSGECancelDummy)

    giValorFreteAlterado = 0
    giValorSeguroAlterado = 0
    giValorDescontoAlterado = 0
    giValorDespesasAlterado = 0

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(objOrcamentoVenda.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    If objOrcamentoVenda.iPrazoValidade <> 0 Then
        PrazoValidade.Text = objOrcamentoVenda.iPrazoValidade
    End If

    'Se a tabela de preços estiver preenchida coloca na tela
    If objOrcamentoVenda.iTabelaPreco > 0 Then
        TabelaPreco.Text = objOrcamentoVenda.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    Else
        TabelaPreco.Text = ""
    End If

    'Preenche o campo de vendedores
    If objOrcamentoVenda.iVendedor <> 0 Then

        Vendedor.Text = objOrcamentoVenda.iVendedor
        Call Vendedor_Validate(bCancel)

    End If

    'Carrega o Tab Cobrança
    lErro = Carrega_Tab_Cobranca(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186637

    'Carrega o Grid de itens
    lErro = Carrega_Grid_Itens(objOrcamentoVenda, iBrowse)
    If lErro <> SUCESSO Then gError 186638

    'Carrega o Grid de itens
    lErro = Carrega_Grid_ItensSRV(objOrcamentoVenda, iBrowse)
    If lErro <> SUCESSO Then gError 186653

    ValorTotal.Caption = Format(objOrcamentoVenda.dValorTotal, "Standard")

'    'Carrega o Tab de Tributação
'    lErro = Carrega_Tab_Tributacao(objOrcamentoVenda)
'    If lErro <> SUCESSO Then gError 186639

    '##################################
    'Inserido por Wagner 09/12/05
    If objOrcamentoVenda.lStatus <> 0 Then Call Combo_Seleciona_ItemData(Status, objOrcamentoVenda.lStatus)
    If objOrcamentoVenda.lMotivoPerda <> 0 Then Call Combo_Seleciona_ItemData(MotivoPerda, objOrcamentoVenda.lMotivoPerda)
    '##################################

    If objOrcamentoVenda.lNumIntSolicSRV <> 0 Then
        
        objSolicSRV.lNumIntDoc = objOrcamentoVenda.lNumIntSolicSRV
        
        lErro = CF("SolicitacaoSRV_Le_NumIntDoc", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 186988 Then gError 186991

        If lErro = SUCESSO Then
            CodSolicSRV.Text = objSolicSRV.lCodigo
        End If
        
    End If
    
    lErro = gobjTelaTrataBloqGen.Carrega_GridBloqueio(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186991

    '###################################################
    'Inserido por Wagner 04/08/2006
'    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objOrcamentoVenda.objProjetoInfo, objOrcamentoVenda.lNumIntDoc, PRJ_CR_TIPO_ORCSRV)
'    If lErro <> SUCESSO Then gError 186640
    '###################################################
    
    Call SubTotal_Calcula

    Set gobjInfoAdic = objOrcamentoVenda.objInfoAdic
    
    ValorDescontoTit.Text = Format(objOrcamentoVenda.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)

    iAlterado = 0

    gbCarregandoTela = False

    Traz_OrcamentoVenda_Tela = SUCESSO

    Exit Function

Erro_Traz_OrcamentoVenda_Tela:

    gbCarregandoTela = False

    Traz_OrcamentoVenda_Tela = gErr

    Select Case gErr

        Case 186628 To 186640, 186653, 186991

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186641)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 186642

    Call TelaIndice_Preenche(Me)

    Exit Sub

Erro_Form_Activate:

    Select Case gErr

        Case 186642

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186643)

    End Select

    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OrcamentoSRV"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186644

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objOrcamentoVenda.lCodigo, 0, "Codigo"
    colCampoValor.Add "Cliente", objOrcamentoVenda.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objOrcamentoVenda.iFilial, 0, "Filial"
    colCampoValor.Add "CondicaoPagto", objOrcamentoVenda.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "PercAcrescFinanceiro", objOrcamentoVenda.dPercAcrescFinanceiro, 0, "PercAcrescFinanceiro"
    colCampoValor.Add "DataEmissao", objOrcamentoVenda.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorTotal", objOrcamentoVenda.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorFrete", objOrcamentoVenda.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorDesconto", objOrcamentoVenda.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorSeguro", objOrcamentoVenda.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "TabelaPreco", objOrcamentoVenda.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "Vendedor", objOrcamentoVenda.iVendedor, 0, "Vendedor"
    colCampoValor.Add "NomeCli", objOrcamentoVenda.sNomeCli, STRING_CLIENTE_NOME_REDUZIDO, "NomeCli"
    colCampoValor.Add "NomeFilialCli", objOrcamentoVenda.sNomeFilialCli, STRING_FILIAL_CLIENTE_NOME, "NomeFilialCli"
    colCampoValor.Add "NaturezaOp", objOrcamentoVenda.sNaturezaOp, STRING_NATUREZAOP_CODIGO, "NaturezaOp" 'por Leo em 02/05/02

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 186644

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186645)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Tela_Preenche

    objOrcamentoVenda.lCodigo = colCampoValor.Item("Codigo").vValor
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    If objOrcamentoVenda.lCodigo <> 0 Then

        'Mostra os dados do Pedido de Venda na tela
        lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGO)
        If lErro <> SUCESSO Then gError 186646

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 186646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186647)

    End Select

    Exit Sub

End Sub

Function Carrega_Tab_Cobranca(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Coloca os dados do tab de cobrança na tela

Dim objParcelaOV As ClassParcelaOV
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dValorDesconto As Double

On Error GoTo Erro_Carrega_Tab_Cobranca

    Call DateParaMasked(DataReferencia, objOrcamentoVenda.dtDataReferencia)
    giDataReferenciaAlterada = 0

    PercAcrescFin.Text = ""

    If objOrcamentoVenda.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objOrcamentoVenda.iCondicaoPagto
        Call CondicaoPagamento_Validate(bSGECancelDummy)
    Else
        CondicaoPagamento.Text = ""
    End If

    PercAcrescFin.Text = Format(objOrcamentoVenda.dPercAcrescFinanceiro * 100, "Fixed")

    CobrancaAutomatica.Value = objOrcamentoVenda.iCobrancaAutomatica

    'Limpa o Grid de Parcelas antes de preencher com dados da coleção
    Call Grid_Limpa(objGridParcelas)

    iIndice = 0

    For Each objParcelaOV In objOrcamentoVenda.colParcela

        iIndice = iIndice + 1
        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcelaOV.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcelaOV.dValor, "Standard")
        If objParcelaOV.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcelaOV.dtDesconto1Ate, "dd/mm/yyyy")
        If objParcelaOV.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcelaOV.dtDesconto2Ate, "dd/mm/yyyy")
        If objParcelaOV.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcelaOV.dtDesconto3Ate, "dd/mm/yyyy")
        If objParcelaOV.iDesconto1Codigo = VALOR_FIXO Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcelaOV.dDesconto1Valor, "Standard")
        ElseIf objParcelaOV.iDesconto1Codigo = Percentual Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col) = Format(objParcelaOV.dDesconto1Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto1Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaOV.iDesconto2Codigo = VALOR_FIXO Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcelaOV.dDesconto2Valor, "Standard")
        ElseIf objParcelaOV.iDesconto2Codigo = Percentual Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col) = Format(objParcelaOV.dDesconto2Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto2Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaOV.iDesconto3Codigo = VALOR_FIXO Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcelaOV.dDesconto3Valor, "Standard")
        ElseIf objParcelaOV.iDesconto3Codigo = Percentual Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col) = Format(objParcelaOV.dDesconto3Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto3Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        For iIndice2 = 0 To TipoDesconto1.ListCount - 1
            If TipoDesconto1.ItemData(iIndice2) = objParcelaOV.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = TipoDesconto1.List(iIndice2)
            If TipoDesconto2.ItemData(iIndice2) = objParcelaOV.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = TipoDesconto2.List(iIndice2)
            If TipoDesconto3.ItemData(iIndice2) = objParcelaOV.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = TipoDesconto3.List(iIndice2)
        Next

    Next

    objGridParcelas.iLinhasExistentes = iIndice

    Carrega_Tab_Cobranca = SUCESSO

    Exit Function

Erro_Carrega_Tab_Cobranca:

    Carrega_Tab_Cobranca = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186648)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_Itens(objOrcamentoVenda As ClassOrcamentoVenda, ByVal iBrowse As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objGridItens1 As Object

On Error GoTo Erro_Carrega_Grid_Itens

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For iIndice = 1 To objOrcamentoVenda.colItens.Count

        objOrcamentoVenda.colItens(iIndice).iItem = iIndice

        lErro = Mascara_RetornaProdutoEnxuto(objOrcamentoVenda.colItens(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 186649

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Calcula o percentual de desconto
        If objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto > 0 Then
            dPercDesc = objOrcamentoVenda.colItens(iIndice).dValorDesconto / (objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto)
        End If

        'Coloca os dados dos itens na tela
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objOrcamentoVenda.colItens(iIndice).sDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objOrcamentoVenda.colItens(iIndice).sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objOrcamentoVenda.colItens(iIndice).dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)

        '###############################################
        'INSERIDO POR WAGNER
        If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
            If iBrowse = VALOR_CODIGO Then
                GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col) = objOrcamentoVenda.colItens(iIndice).sVersaoKit
                GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col) = objOrcamentoVenda.colItens(iIndice).sVersaoKitBase
            Else
                GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col) = Codigo.Text & SEPARADOR & iIndice
                GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col) = objOrcamentoVenda.colItens(iIndice).sVersaoKit
            End If
        End If
        '###############################################

        'precoDesc
        Set objGridItens1 = GridItens
        Call CF("Carrega_Grid_Itens_PrecoDesc", objGridItens1, iIndice, iGrid_PrecoUnitario_Col + 1, Format(objOrcamentoVenda.colItens(iIndice).dPrecoUnitario * (1 - dPercDesc), gobjFAT.sFormatoPrecoUnitario), objOrcamentoVenda.colItens(iIndice).sProduto, iGrid_DescProduto_Col)

        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objOrcamentoVenda.colItens(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objOrcamentoVenda.colItens(iIndice).sObservacao

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = objOrcamentoVenda.colItens.Count

    'Atualiza o grid para mostrar as checkboxes marcadas / desmarcadas
    Call Grid_Refresh_Checkbox(objGridItens)

    Carrega_Grid_Itens = SUCESSO

    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = gErr

    Select Case gErr

        Case 186649
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objOrcamentoVenda.colItens(iIndice).sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186650)

    End Select

    Exit Function

End Function

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
    If lErro <> SUCESSO Then gError 186654

    lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 186655

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)

    'Se o Tipo estiver preenchido
    If objcliente.iTipo > 0 Then
        objTipoCliente.iCodigo = objcliente.iTipo
        'Lê o Tipo de Cliente
        lErro = CF("TipoCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 186656
    End If

    'Guarda o valor do desconto do cliente
    If objcliente.dDesconto > 0 Then
        gdDesconto = objcliente.dDesconto
    ElseIf objTipoCliente.dDesconto > 0 Then
        gdDesconto = objTipoCliente.dDesconto
    Else
        gdDesconto = 0
    End If

    'para fazer valer o que veio do bd
    giValorDescontoManual = 1

    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
    giClienteAlterado = 0

    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case 186654 To 186656

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186657)

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
    If lErro <> SUCESSO And lErro <> 17660 Then gError 186658

    If lErro = 17660 Then gError 186659

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 186658

        Case 186659
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186660)

    End Select

    Exit Sub

End Sub
'
'Public Sub ComboItensTrib_Click()
'
'Dim iIndice As Integer, objItemOV As ClassItemOV
'
'On Error GoTo Erro_ComboItensTrib_Click
'
'    iIndice = ComboItensTrib.ListIndex
'
'    If iIndice <> -1 Then
'
'        If iIndice + 1 > gobjOrcamentoVenda.colItensSRV.Count Then
'
'            Set objItemOV = gobjOrcamentoVenda.colItens.Item(iIndice + 1 - gobjOrcamentoVenda.colItensSRV.Count)
'
'        Else
'
'
'            Set objItemOV = gobjOrcamentoVenda.colItensSRV.Item(iIndice + 1)
'
'        End If
'
'        LabelValorItem.Caption = Format(objItemOV.dPrecoTotal, "Standard")
'        LabelQtdeItem.Caption = CStr(objItemOV.dQuantidade)
'        LabelUMItem.Caption = objItemOV.sUnidadeMed
'
'        Call TributacaoItem_TrazerTela(objItemOV.objTributacaoItemOV)
'
'    End If
'
'    Exit Sub
'
'Erro_ComboItensTrib_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186661)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TribSobreDesconto_Click()
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
'    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto)
'
'End Sub
'
'Public Sub TribSobreOutrasDesp_Click()
'
'   'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras)
'
'End Sub
'
'Public Sub TribSobreSeguro_Click()
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
'    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro)
'
'End Sub
'
'Public Sub TribSobreFrete_Click()
'
'    'exibir o frame de "outros"
'    FrameOutrosTrib.Visible = True
'    FrameItensTrib.Visible = False
'
'    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete)
'
'End Sub
'
'Public Sub TribSobreItem_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    'se houver itens na combo
'    If gobjOrcamentoVenda.colItens.Count <> 0 Or gobjOrcamentoVenda.colItensSRV.Count <> 0 Then
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

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iVendedorAlterado = 1


End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If iVendedorAlterado = 1 Then

        If Len(Trim(Vendedor.Text)) > 0 Then

            'Tenta ler o Vendedor (NomeReduzido ou Código)
            lErro = TP_Vendedor_Le(Vendedor, objVendedor)
            If lErro <> SUCESSO Then gError 186662

        End If

        iVendedorAlterado = 0

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case 186662

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 186663)

    End Select

End Sub

Public Sub VendedorLabel_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection

    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text

    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        ElseIf Me.ActiveControl Is ServicoSRV Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoPecas_Click
'        ElseIf Me.ActiveControl Is NatOpInterna Then
'            Call LblNatOpInterna_Click
'        ElseIf Me.ActiveControl Is NaturezaOpItem Then
'            Call NaturezaItemLabel_Click
'        ElseIf Me.ActiveControl Is TipoTributacaoItem Then
'            Call LblTipoTribItem_Click
'        ElseIf Me.ActiveControl Is TipoTributacao Then
'            Call LblTipoTrib_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call NumeroLabel_Click
        'Inserido por Wagner
        '#############################################
        ElseIf Me.ActiveControl Is CodigoBase Then
            Call NumeroBaseLabel_Click
        '#############################################
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
'        ElseIf Me.ActiveControl Is NaturezaOp Then 'por Leo em 02/05/02
'            Call NaturezaLabel_Click
        '#########################################
        'Inserido por Wagner 10/08/2006
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        '#########################################
        End If

    End If

End Sub

Private Function Tipo_Cliente(ByVal sCliente As String) As enumTipo

    If Len(Trim(sCliente)) = 0 Then
        Tipo_Cliente = TIPO_VAZIO
    ElseIf Not IsNumeric(sCliente) Then
        Tipo_Cliente = TIPO_STRING
    ElseIf Int(CDbl(sCliente)) <> CDbl(sCliente) Then
        Tipo_Cliente = TIPO_DECIMAL
    ElseIf CDbl(sCliente) <= 0 Then
        Tipo_Cliente = TIPO_NAO_POSITIVO
    ElseIf Len(Trim(sCliente)) > STRING_CGC Then
        Tipo_Cliente = TIPO_OVERFLOW
    ElseIf Len(Trim(sCliente)) > STRING_CPF Then
        Tipo_Cliente = TIPO_CGC
    ElseIf CDbl(sCliente) > NUM_MAX_CLIENTES Then
        Tipo_Cliente = TIPO_CPF
    Else
        Tipo_Cliente = TIPO_CODIGO
    End If

End Function

Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)

    Nome_Extrai = sString

    Exit Function

End Function
'
'Private Function TributacaoOrcSRV_Reset(Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long
''cria ou atualiza gobjOrcamentoVenda, com dados correspondentes a objOrcamentoVenda (se este for passado) ou com dados "padrao"
'
'Dim lErro As Long
'Dim objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoOrcSRV_Reset
'
'    'se gobjOrcamentoVenda já foi inicializado
'    If Not (gobjOrcamentoVenda Is Nothing) Then
'
'        Set objTributoDoc = gobjOrcamentoVenda
'
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then gError 186627
'
'        Set gobjOrcamentoVenda = Nothing
'
'    End If
'
'    'se o pedido de venda veio preenchido
'    If Not (objOrcamentoVenda Is Nothing) Then
'
'        Set gobjOrcamentoVenda = objOrcamentoVenda
'
'    Else
'
'        Set gobjOrcamentoVenda = New ClassOrcamentoVenda
'        gobjOrcamentoVenda.dtDataEmissao = gdtDataAtual
'
'    End If
'
'    Set objTributoDoc = gobjOrcamentoVenda
'    lErro = objTributoDoc.Ativar
'    If lErro <> SUCESSO Then gError 186628
'
'    giNaturezaOpAlterada = 0
'    giISSAliquotaAlterada = 0
'    giISSValorAlterado = 0
'    giValorIRRFAlterado = 0
'    giTipoTributacaoAlterado = 0
'    giAliqIRAlterada = 0
'    iPISRetidoAlterado = 0
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
'    TributacaoOrcSRV_Reset = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoOrcSRV_Reset:
'
'    TributacaoOrcSRV_Reset = gErr
'
'    Select Case gErr
'
'        Case 186627, 186628
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186629)
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
'    If lErro <> SUCESSO Then gError 186664
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 186665
'
'    lErro = Carrega_Tab_Tributacao(gobjOrcamentoVenda)
'    If lErro <> SUCESSO Then gError 186666
'
'    Exit Sub
'
'Erro_BotaoGravarTrib:
'
'    Select Case gErr
'
'        Case 186664 To 186666
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186667)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Tributacao_GravarTela() As Long
''transfere dados de tributacao da tela para gobjOrcamentoVenda
''os dados que estiverem diferentes devem ser marcados como "manuais"
'
'Dim lErro As Long
'Dim iIndice As Integer, iTemp As Integer, dTemp As Double, objTributacaoOV As ClassTributacaoOV
'
'On Error GoTo Erro_Tributacao_GravarTela
'
'    Set objTributacaoOV = gobjOrcamentoVenda.objTributacaoOV
'
'    If gobjOrcamentoVenda.sNaturezaOp <> gobjTribTab.sNatOpInterna Then
'
'        gobjOrcamentoVenda.sNaturezaOp = gobjTribTab.sNatOpInterna
'        gobjOrcamentoVenda.iNaturezaOpManual = VAR_PREENCH_MANUAL
'
'    End If
'
'    iTemp = StrParaInt(TipoTributacao.Text)
'    If iTemp <> objTributacaoOV.iTipoTributacao Then
'        objTributacaoOV.iTipoTributacao = iTemp
'        objTributacaoOV.iTipoTributacaoManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados de ISS
'    iTemp = ISSIncluso.Value
'    If iTemp <> objTributacaoOV.iISSIncluso Then
'        objTributacaoOV.iISSIncluso = iTemp
'        objTributacaoOV.iISSInclusoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ISSAliquota.Text <> CStr(objTributacaoOV.dISSAliquota * 100) Then
'        dTemp = StrParaDbl(ISSAliquota.Text) / 100
'        If objTributacaoOV.dISSAliquota <> dTemp Then
'            objTributacaoOV.dISSAliquota = dTemp
'            objTributacaoOV.iISSAliquotaManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If ISSValor.Text <> CStr(objTributacaoOV.dISSValor) Then
'        dTemp = StrParaDbl(ISSValor.Text)
'        If objTributacaoOV.dISSValor <> dTemp Then
'            objTributacaoOV.dISSValor = dTemp
'            objTributacaoOV.iISSValorManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    'setar dados de IR
'    If IRAliquota.Text <> CStr(objTributacaoOV.dIRRFAliquota * 100) Then
'        dTemp = StrParaDbl(IRAliquota.Text) / 100
'        If objTributacaoOV.dIRRFAliquota <> dTemp Then
'            objTributacaoOV.dIRRFAliquota = dTemp
'            objTributacaoOV.iIRRFAliquotaManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If ValorIRRF.Text <> CStr(objTributacaoOV.dIRRFValor) Then
'        dTemp = StrParaDbl(ValorIRRF.Text)
'        If objTributacaoOV.dIRRFValor <> dTemp Then
'            objTributacaoOV.dIRRFValor = dTemp
'            objTributacaoOV.iIRRFValorManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If PISRetido.Text <> CStr(objTributacaoOV.dPISRetido) Then
'        dTemp = StrParaDbl(PISRetido.Text)
'        If objTributacaoOV.dPISRetido <> dTemp Then
'            objTributacaoOV.dPISRetido = dTemp
'            objTributacaoOV.iPISRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If COFINSRetido.Text <> CStr(objTributacaoOV.dCOFINSRetido) Then
'        dTemp = StrParaDbl(COFINSRetido.Text)
'        If objTributacaoOV.dCOFINSRetido <> dTemp Then
'            objTributacaoOV.dCOFINSRetido = dTemp
'            objTributacaoOV.iCOFINSRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If CSLLRetido.Text <> CStr(objTributacaoOV.dCSLLRetido) Then
'        dTemp = StrParaDbl(CSLLRetido.Text)
'        If objTributacaoOV.dCSLLRetido <> dTemp Then
'            objTributacaoOV.dCSLLRetido = dTemp
'            objTributacaoOV.iCSLLRetidoManual = VAR_PREENCH_MANUAL
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
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186668)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Carrega_Tab_Tributacao(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'
'Dim lErro As Long
'Dim objTributacaoOV As ClassTributacaoOV
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'
'On Error GoTo Erro_Carrega_Tab_Tributacao
'
'    giTrazendoTribTela = 1
'
'    Set objTributacaoOV = objOrcamentoVenda.objTributacaoOV
'
'    If gobjTribTab.sNatOpInterna <> objOrcamentoVenda.sNaturezaOp Then
'
'        gobjTribTab.sNatOpInterna = objOrcamentoVenda.sNaturezaOp
'        Call NaturezaOp_Validate(bSGECancelDummy)
'
'    End If
'
'    'no frame de "resumo"
'    objTributacaoTipo.iTipo = objTributacaoOV.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        TipoTributacao.Text = CStr(objTributacaoOV.iTipoTributacao)
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then gError 186669
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
'    IPIBase.Caption = Format(objTributacaoOV.dIPIBase, "Standard")
'    IPIValor.Caption = Format(objTributacaoOV.dIPIValor, "Standard")
'    ISSBase.Caption = Format(objTributacaoOV.dISSBase, "Standard")
'    ISSAliquota.Text = CStr(objTributacaoOV.dISSAliquota * 100)
'    ISSValor.Text = CStr(objTributacaoOV.dISSValor)
'    ISSIncluso.Value = objTributacaoOV.iISSIncluso
'    IRBase.Caption = Format(objTributacaoOV.dIRRFBase, "Standard")
'    IRAliquota.Text = CStr(objTributacaoOV.dIRRFAliquota * 100)
'    ValorIRRF.Text = CStr(objTributacaoOV.dIRRFValor)
'    ICMSBase.Caption = Format(objTributacaoOV.dICMSBase, "Standard")
'    ICMSValor.Caption = Format(objTributacaoOV.dICMSValor, "Standard")
'    ICMSSubstBase.Caption = Format(objTributacaoOV.dICMSSubstBase, "Standard")
'    ICMSSubstValor.Caption = Format(objTributacaoOV.dICMSSubstValor, "Standard")
'    PISRetido.Text = CStr(objTributacaoOV.dPISRetido)
'    COFINSRetido.Text = CStr(objTributacaoOV.dCOFINSRetido)
'    CSLLRetido.Text = CStr(objTributacaoOV.dCSLLRetido)
'
'    'o frame de "detalhamento" vou deixar p/carregar qdo o usuario entrar nele
'
'    giISSAliquotaAlterada = 0
'    giISSValorAlterado = 0
'    giValorIRRFAlterado = 0
'    giTipoTributacaoAlterado = 0
'    giAliqIRAlterada = 0
'    iPISRetidoAlterado = 0
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
'        Case 186669
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186670)
'
'    End Select
'
'End Function
'
'Private Sub BotaoGravarTribItem_Click()
'
'Dim lErro As Long, objTributacaoItemOV As ClassTributacaoItemPV, iIndice As Integer
'
'On Error GoTo Erro_BotaoGravarTribItem_Click
'
'    'atualizar dados da colecao p/o item ou complemento corrente
'
'    'se um item estiver selecionado
'    If TribSobreItem.Value = True Then
'        iIndice = ComboItensTrib.ListIndex
'       If iIndice <> -1 Then
'            If iIndice + 1 > gobjOrcamentoVenda.colItensSRV.Count Then
'                Set objTributacaoItemOV = gobjOrcamentoVenda.colItens.Item(iIndice + 1 - gobjOrcamentoVenda.colItensSRV.Count).objTributacaoItemOV
'            Else
'                Set objTributacaoItemOV = gobjOrcamentoVenda.colItensSRV.Item(iIndice + 1).objTributacaoItemOV
'            End If
'        Else
'            gError 186671
'        End If
'    Else
'        If TribSobreDesconto.Value = True Then
'            Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto
'        Else
'            If TribSobreFrete.Value = True Then
'                Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete
'            Else
'                If TribSobreSeguro.Value = True Then
'                    Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro
'                Else
'                    If TribSobreOutrasDesp.Value = True Then
'                        Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    lErro = TributacaoItem_GravarTela(objTributacaoItemOV)
'    If lErro <> SUCESSO Then gError 186672
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 186673
'
'    lErro = TributacaoItem_TrazerTela(objTributacaoItemOV)
'    If lErro <> SUCESSO Then gError 186674
'
'    Exit Sub
'
'Erro_BotaoGravarTribItem_Click:
'
'    Select Case gErr
'
'        Case 186671
'            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEM_TRIB_SEL", gErr, Error)
'
'        Case 186672 To 186674
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186675)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function TributacaoItem_GravarTela(objTributacaoItemOV As ClassTributacaoItemPV) As Long
''transfere dados de tributacao de um item da tela para objTributacaoItemOV
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
'    If Trim(objTributacaoItemOV.sNaturezaOp) <> sTemp Then
'        objTributacaoItemOV.sNaturezaOp = sTemp
'        objTributacaoItemOV.iNaturezaOpManual = VAR_PREENCH_MANUAL
'    End If
'
'    iTemp = StrParaInt(TipoTributacaoItem.Text)
'    If iTemp <> objTributacaoItemOV.iTipoTributacao Then
'        objTributacaoItemOV.iTipoTributacao = iTemp
'        objTributacaoItemOV.iTipoTributacaoManual = VAR_PREENCH_MANUAL
'    End If
'
'    'Setar dados de ICMS
'
'    iTemp = ComboICMSTipo.ItemData(ComboICMSTipo.ListIndex)
'    If iTemp <> objTributacaoItemOV.iICMSTipo Then
'        objTributacaoItemOV.iICMSTipo = iTemp
'        objTributacaoItemOV.iICMSTipoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSBaseItem.Text <> CStr(objTributacaoItemOV.dICMSBase) Then
'        dTemp = StrParaDbl(ICMSBaseItem.Text)
'        objTributacaoItemOV.dICMSBase = dTemp
'        objTributacaoItemOV.iICMSBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSPercRedBaseItem.Text <> CStr(objTributacaoItemOV.dICMSPercRedBase * 100) Then
'        dTemp = StrParaDbl(ICMSPercRedBaseItem.Text) / 100
'        objTributacaoItemOV.dICMSPercRedBase = dTemp
'        objTributacaoItemOV.iICMSPercRedBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSAliquotaItem.Text <> CStr(objTributacaoItemOV.dICMSAliquota * 100) Then
'        dTemp = StrParaDbl(ICMSAliquotaItem.Text) / 100
'        objTributacaoItemOV.dICMSAliquota = dTemp
'        objTributacaoItemOV.iICMSAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSValorItem.Text <> CStr(objTributacaoItemOV.dICMSValor) Then
'        dTemp = StrParaDbl(ICMSValorItem.Text)
'        objTributacaoItemOV.dICMSValor = dTemp
'        objTributacaoItemOV.iICMSValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados ICMS Substituicao
'
'    If ICMSSubstBaseItem.Text <> CStr(objTributacaoItemOV.dICMSSubstBase) Then
'        dTemp = StrParaDbl(ICMSSubstBaseItem.Text)
'        objTributacaoItemOV.dICMSSubstBase = dTemp
'        objTributacaoItemOV.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSSubstAliquotaItem.Text <> CStr(objTributacaoItemOV.dICMSSubstAliquota * 100) Then
'        dTemp = StrParaDbl(ICMSSubstAliquotaItem.Text) / 100
'        objTributacaoItemOV.dICMSSubstAliquota = dTemp
'        objTributacaoItemOV.iICMSSubstAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSSubstValorItem.Text <> CStr(objTributacaoItemOV.dICMSSubstValor) Then
'        dTemp = StrParaDbl(ICMSSubstValorItem.Text)
'        objTributacaoItemOV.dICMSSubstValor = dTemp
'        objTributacaoItemOV.iICMSSubstValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados de IPI
'    iTemp = ComboIPITipo.ItemData(ComboIPITipo.ListIndex)
'    If iTemp <> objTributacaoItemOV.iIPITipo Then
'        objTributacaoItemOV.iIPITipo = iTemp
'        objTributacaoItemOV.iIPITipoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIBaseItem.Text <> CStr(objTributacaoItemOV.dIPIBaseCalculo) Then
'        dTemp = StrParaDbl(IPIBaseItem.Text)
'        objTributacaoItemOV.dIPIBaseCalculo = dTemp
'        objTributacaoItemOV.iIPIBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIPercRedBaseItem.Text <> CStr(objTributacaoItemOV.dIPIPercRedBase * 100) Then
'        dTemp = StrParaDbl(IPIPercRedBaseItem.Text) / 100
'        objTributacaoItemOV.dIPIPercRedBase = dTemp
'        objTributacaoItemOV.iIPIPercRedBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIAliquotaItem.Text <> CStr(objTributacaoItemOV.dIPIAliquota * 100) Then
'        dTemp = StrParaDbl(IPIAliquotaItem.Text) / 100
'        objTributacaoItemOV.dIPIAliquota = dTemp
'        objTributacaoItemOV.iIPIAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIValorItem.Text <> CStr(objTributacaoItemOV.dIPIValor) Then
'        dTemp = StrParaDbl(IPIValorItem.Text)
'        objTributacaoItemOV.dIPIValor = dTemp
'        objTributacaoItemOV.iIPIValorManual = VAR_PREENCH_MANUAL
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
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186676)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TributacaoItem_TrazerTela(objTributacaoItemOV As ClassTributacaoItemPV) As Long
''Traz para a tela dados de tributacao de um item
'
'Dim iIndice As Integer
'Dim objItemOrcamento As ClassItemOV
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
'    NaturezaOpItem.Text = objTributacaoItemOV.sNaturezaOp
'
'    sNatOp = Trim(NaturezaOpItem.Text)
'
'    If sNatOp <> "" Then
'
'        objNaturezaOp.sCodigo = sNatOp
'        'Lê a Natureza de Operação
'        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'        If lErro <> SUCESSO And lErro <> 17958 Then gError 186677
'
'        'Se não achou a Natureza de Operação --> erro
'        If lErro <> SUCESSO Then gError 186678
'
'        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
'    Else
'        LabelDescrNatOpItem.Caption = ""
'    End If
'
'    objTributacaoTipo.iTipo = objTributacaoItemOV.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then gError 186679
'
'        TipoTributacaoItem.Text = CStr(objTributacaoItemOV.iTipoTributacao)
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
'    Call Combo_Seleciona_ItemData(ComboICMSTipo, objTributacaoItemOV.iICMSTipo)
'
'    ICMSBaseItem.Text = CStr(objTributacaoItemOV.dICMSBase)
'    ICMSPercRedBaseItem.Text = CStr(objTributacaoItemOV.dICMSPercRedBase * 100)
'    ICMSAliquotaItem.Text = CStr(objTributacaoItemOV.dICMSAliquota * 100)
'    ICMSValorItem.Text = CStr(objTributacaoItemOV.dICMSValor)
'
'    'setar dados ICMS Substituicao
'    ICMSSubstBaseItem.Text = CStr(objTributacaoItemOV.dICMSSubstBase)
'    ICMSSubstAliquotaItem.Text = CStr(objTributacaoItemOV.dICMSSubstAliquota * 100)
'    ICMSSubstValorItem.Text = CStr(objTributacaoItemOV.dICMSSubstValor)
'
'    For Each objTipoTribICMS In gcolTiposTribICMS
'        If objTipoTribICMS.iTipo = objTributacaoItemOV.iICMSTipo Then Exit For
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
'    Call Combo_Seleciona_ItemData(ComboIPITipo, objTributacaoItemOV.iIPITipo)
'
'    IPIBaseItem.Text = CStr(objTributacaoItemOV.dIPIBaseCalculo)
'    IPIPercRedBaseItem.Text = CStr(objTributacaoItemOV.dIPIPercRedBase * 100)
'    IPIAliquotaItem.Text = CStr(objTributacaoItemOV.dIPIAliquota * 100)
'    IPIValorItem.Text = CStr(objTributacaoItemOV.dIPIValor)
'
'    For Each objTipoTribIPI In gcolTiposTribIPI
'        If objTipoTribIPI.iTipo = objTributacaoItemOV.iIPITipo Then Exit For
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
'        Case 186677, 186679
'
'        Case 186678
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", objNaturezaOp.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186680)
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
'Dim objItemOrcamento As ClassItemOV
'Dim sItem As String
'
'On Error GoTo Erro_TributacaoItem_InicializaTab
'
'    'preencher o valor de frete, seguro, descontos e outras desp no frameOutros
'    LabelValorFrete.Caption = Format(gobjOrcamentoVenda.dValorFrete, "Standard")
'    LabelValorDesconto.Caption = Format(gobjOrcamentoVenda.dValorDesconto, "Standard")
'    LabelValorSeguro.Caption = Format(gobjOrcamentoVenda.dValorSeguro, "Standard")
'    LabelValorOutrasDespesas.Caption = Format(gobjOrcamentoVenda.dValorOutrasDespesas, "Standard")
'
'    'esvaziar a combo de itens
'    ComboItensTrib.Clear
'
'    'preencher a combo de itens: com "codigo do produto - descricao"
'
'    For Each objItemOrcamento In gobjOrcamentoVenda.colItensSRV
'
'        lErro = Mascara_MascararProduto(objItemOrcamento.sProduto, sItem)
'        If lErro <> SUCESSO Then gError 186681
'
'        sItem = sItem & " - " & objItemOrcamento.sDescricao
'        ComboItensTrib.AddItem sItem
'
'    Next
'
'
'    For Each objItemOrcamento In gobjOrcamentoVenda.colItens
'
'        lErro = Mascara_MascararProduto(objItemOrcamento.sProduto, sItem)
'        If lErro <> SUCESSO Then gError 186682
'
'        sItem = sItem & " - " & objItemOrcamento.sDescricao
'        ComboItensTrib.AddItem sItem
'
'    Next
'
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
'        Case 186681, 186682
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186683)
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
'    If lErro <> SUCESSO Then gError 186684
'
'    'Atualiza os valores de tributação
'    lErro = AtualizarTributacao()
'    If lErro <> SUCESSO Then gError 186685
'
'    Exit Sub
'
'Erro_BotaoGravarTribCarga:
'
'    Select Case gErr
'
'        Case 186684, 186685
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186686)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function TributacaoOV_Terminar() As Long
'
'Dim lErro As Long, objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoOV_Terminar
'
'    If Not (gobjOrcamentoVenda Is Nothing) Then
'        Set objTributoDoc = gobjOrcamentoVenda
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then gError 186687
'        Set gobjOrcamentoVenda = Nothing
'    End If
'
'    TributacaoOV_Terminar = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoOV_Terminar:
'
'    TributacaoOV_Terminar = gErr
'
'    Select Case gErr
'
'        Case 186687
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186688)
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
'    If lErro <> SUCESSO Then gError 186689
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
'    If lErro <> SUCESSO Then gError 186690
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
'        Case 186689, 186690
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186691)
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
'
'    If Not (gobjOrcamentoVenda Is Nothing) Then
'
'        'Atualiza os impostos
'        lErro = gobjTributacao.AtualizaImpostos(gobjOrcamentoVenda, giRecalculandoTributacao)
'        If lErro <> SUCESSO Then gError 186692
'
'        'joga dados do obj atualizado p/a tela
'        lErro = Carrega_Tab_Tributacao(gobjOrcamentoVenda)
'        If lErro <> SUCESSO Then gError 186693
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
'        Case 186692, 186693
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186694)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Tributacao_Inclusao_Item_Grid(iLinha As Integer, sProduto As String) As Long
''trata a inclusao de uma linha de item no grid
'Dim lErro As Long
'Dim objTributoDocItem As ClassTributoDocItem
'Dim objItemOV As ClassItemOV
'On Error GoTo Erro_Tributacao_Inclusao_Item_Grid
'
'    lErro = Move_GridItem_Memoria(gobjOrcamentoVenda, iLinha, sProduto)
'    If lErro <> SUCESSO Then gError 186695
'
'    Set objItemOV = gobjOrcamentoVenda.colItens.Item(iLinha)
'    Set objTributoDocItem = objItemOV
'
'    lErro = objTributoDocItem.Ativar(gobjOrcamentoVenda)
'    If lErro <> SUCESSO Then gError 186696
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
'        Case 186695, 186696
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186697)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function gobjTribTab.Exclusao_Item_Grid(iLinha As Integer) As Long
''trata a exclusao de uma linha de item no grid
'Dim objItemOV As ClassItemOV, objTributoDocItem As ClassTributoDocItem
'
'        Set objItemOV = gobjOrcamentoVenda.colItens(iLinha)
'        Set objTributoDocItem = objItemOV
'        Call objTributoDocItem.Desativar
'        Call gobjOrcamentoVenda.RemoverItem(iLinha)
'
'End Function
'
'Function gobjTribTab.Exclusao_Item_GridSRV(iLinha As Integer) As Long
''trata a exclusao de uma linha de item no grid
'Dim objItemOV As ClassItemOV, objTributoDocItem As ClassTributoDocItem
'
'        Set objItemOV = gobjOrcamentoVenda.colItensSRV(iLinha)
'        Set objTributoDocItem = objItemOV
'        Call objTributoDocItem.Desativar
'        Call gobjOrcamentoVenda.RemoverItemSRV(iLinha)
'
'End Function
'
'
'Function Tributacao_Alteracao_Item_Grid(iIndice As Integer) As Long
''trata a alteracao de uma linha de item no grid
'
'Dim lErro As Long, sProduto As String, iPreenchido As Integer
'Dim objItemOV As ClassItemOV
'
'On Error GoTo Erro_Tributacao_Alteracao_Item_Grid
'
'    Set objItemOV = gobjOrcamentoVenda.colItens.Item(iIndice)
'
'    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then
'
'        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 186485
'
'        objItemOV.sProduto = sProduto
'
'    End If
'
'    objItemOV.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
'
'    objItemOV.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'    'objItemOV.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
'    objItemOV.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'    objItemOV.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
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
'        Case 186485
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186486)
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
'Dim objItemOV As ClassItemOV
'Dim iIndice As Integer, dtDataRef As Date
'
'On Error GoTo Erro_Valida_Tributacao_Gravacao
'
'    If gobjOrcamentoVenda.objTributacaoOV.iTipoTributacao = 0 Then gError 186704
'
'    dtDataRef = gobjOrcamentoVenda.dtDataEmissao
'
'    iIndice = 0
'
'    For Each objItemOV In gobjOrcamentoVenda.colItensSRV
'
'        iIndice = iIndice + 1
'        If Len(Trim(objItemOV.objTributacaoItemOV.sNaturezaOp)) = 0 Then gError 186698
'        If objItemOV.objTributacaoItemOV.iTipoTributacao = 0 Then gError 186699
'        If Natop_ErroTamanho(dtDataRef, objItemOV.objTributacaoItemOV.sNaturezaOp) Then gError 186700
'
'    Next
'
'    For Each objItemOV In gobjOrcamentoVenda.colItens
'
'        iIndice = iIndice + 1 + gobjOrcamentoVenda.colItensSRV.Count
'        If Len(Trim(objItemOV.objTributacaoItemOV.sNaturezaOp)) = 0 Then gError 186701
'        If objItemOV.objTributacaoItemOV.iTipoTributacao = 0 Then gError 186702
'        If Natop_ErroTamanho(dtDataRef, objItemOV.objTributacaoItemOV.sNaturezaOp) Then gError 186703
'
'    Next
'
'    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto.sNaturezaOp)) = 0 Then gError 186705
'    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto.iTipoTributacao = 0 Then gError 186706
'
'    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete.sNaturezaOp)) = 0 Then gError 186707
'    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete.iTipoTributacao = 0 Then gError 186708
'
'    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras.sNaturezaOp)) = 0 Then gError 186709
'    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras.iTipoTributacao = 0 Then gError 186710
'
'    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro.sNaturezaOp)) = 0 Then gError 186711
'    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro.iTipoTributacao = 0 Then gError 186712
'
'    If Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras.sNaturezaOp) Or _
'        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro.sNaturezaOp) Then gError 186715
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
'        Case 186698, 186701
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ITEM_TRIBUTACAO_NAO_PREENCHIDA", iIndice)
'
'        Case 186699, 186702
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
'
'        Case 186700, 186703, 186715
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_TAMANHO_INCORRETO", gErr)
'
'        Case 186704
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)
'
'        Case 186705
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESCONTO_NAO_PRENCHIDA", gErr)
'
'        Case 186706
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESCONTO_NAO_PREENCHIDO", gErr)
'
'        Case 186707
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_FRETE_NAO_PRENCHIDA", gErr)
'
'        Case 186708
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_FRETE_NAO_PREENCHIDO", gErr)
'
'        Case 186709
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESPESAS_NAO_PRENCHIDA", gErr)
'
'        Case 186710
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESPESAS_NAO_PREENCHIDO", gErr)
'
'        Case 186711
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SEGURO_NAO_PRENCHIDA", gErr)
'
'        Case 186712
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_SEGURO_NAO_PREENCHIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186713)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents
        
        Set gobjTelaTrataBloqGen = New ClassTelaTrataBloqGen
        
        lErro = gobjTelaTrataBloqGen.Inicializa(MAPBLOQGEN_TIPOTELA_ORCAMENTOSRV, Me)
        If lErro <> SUCESSO Then gError 186716

        gbCarregandoTela = True
        
        ''*** incluidos p/tratamento de tributacao *******************************
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, , , gobjTribTab.TIPOTELA_OV, , True)
        If lErro <> SUCESSO Then gError 186716
        
        Set gobjOrcamentoVenda = New ClassOrcamentoVenda
        gobjOrcamentoVenda.iServico = MARCADO
        gobjOrcamentoVenda.dtDataEmissao = gdtDataHoje
        lErro = gobjTribTab.TributacaoNF_Reset(gobjOrcamentoVenda)
        If lErro <> SUCESSO Then gError 186716

'        lErro = TributacaoOrcSRV_Reset()
'        If lErro <> SUCESSO Then gError 186716
'
'        Call BotaoGravarTribCarga
'
'        lErro = CarregaTiposTrib()
'        If lErro <> SUCESSO Then gError 186717

        'Carrega a combo combo de Tabela de Preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 186718

        'Carrega a combo de Condição de Pagamento
        lErro = Carrega_CondicaoPagamento()
        If lErro <> SUCESSO Then gError 186719

        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario

        Quantidade.Format = FORMATO_ESTOQUE

        'Preenche Data Referencia e Data de Emissão coma Data Atual
        DataReferencia.PromptInclude = False
        DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataReferencia.PromptInclude = True
        giDataReferenciaAlterada = 0

        Set objGridItens = New AdmGrid
        Set objGridParcelas = New AdmGrid
        Set objGridItensSRV = New AdmGrid
        'Set objGridBloqueio = New AdmGrid

        Set objEventoCliente = New AdmEvento
        Set objEventoNumero = New AdmEvento
        Set objEventoCondPagto = New AdmEvento
        Set objEventoServico = New AdmEvento
        Set objEventoVendedor = New AdmEvento
'        Set objEventoNaturezaOp = New AdmEvento
'        Set objEventoTiposDeTributacao = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoCodSolicSRV = New AdmEvento


'        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
'        Set gobjTelaProjetoInfo.objUserControl = objUserControl
'        Set gobjTelaProjetoInfo.objTela = Me

        'Faz as Inicializações dos Grids
        lErro = Inicializa_Grid_ItensSRV(objGridItensSRV)
        If lErro <> SUCESSO Then gError 186768

        'Faz as Inicializações dos Grids
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 186720

'        lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
'        If lErro <> SUCESSO Then gError 186720

        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 186721

        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ServicoSRV)
        If lErro <> SUCESSO Then gError 186769

        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError 186722

        gbCarregandoTela = False

        iAlterado = 0
        iDescontoAlterado = 0

        giPosCargaOk = 1

        Call ValorTotal_Calcula

    End If

    CargaPosFormLoad = SUCESSO

    Exit Function

Erro_CargaPosFormLoad:

    gbCarregandoTela = False

    CargaPosFormLoad = gErr

    Select Case gErr

        Case 186716 To 186722, 186768, 186769
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186723)

    End Select

    Exit Function

End Function

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Preenche campo CondicaoPagamento
    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida

    'Altera PercAcrescFin
    If Len(Trim(PercAcrescFin.ClipText)) > 0 Then

        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100

        If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
    Else
        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
        Call PercAcrescFin_Validate(bSGECancelDummy)
    End If

    If Len(Trim(ValorTotal.Caption)) > 0 Then
        'Se DataReferencia estiver preenchida e Valor for positivo
        If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then

            'Preenche GridParcelas a partir da Condição de Pagto
            lErro = Cobranca_Automatica()
            If lErro <> SUCESSO Then gError 186724

        End If
    End If

    Me.Show

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case 186724

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186725)

     End Select

     Exit Sub

End Sub

Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o código do orçamento não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 186726

    'Dispara função para imprimir orçamento
    lErro = Orcamento_Imprime(Trim(Codigo.Text))
    If lErro <> SUCESSO Then gError 186727

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 186726
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ORCAMENTO_NAO_PREENCHIDO", gErr)

        Case 186727

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186728)

    End Select

    Exit Sub

End Sub

Private Function Orcamento_Imprime(ByVal lOrcamento As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Orcamento_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Guarda no obj o código do orçamento passado como parâmetro
    objOrcamentoVenda.lCodigo = lOrcamento

    'Guarda a FilialEmpresa ativa como filial do orçamento
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    'Lê os dados do orçamento para verificar se o mesmo existe no BD
    lErro = CF("OrcamentoSRV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186534 Then gError 186729

    'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
    If lErro <> SUCESSO Then gError 186730

    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDireto("Orçamento de Serviço", "OrcamentoSRV >= @NORCINIC E OrcamentoSRV <= @NORCFIM", 1, "OrcSrv", "NORCINIC", Trim(Codigo.Text), "NORCFIM", Trim(Codigo.Text), "NEXIBIRORC", IIf(ImprimirOVComPreco.Value = vbChecked, 1, 0), "NEXIBIRCODPROD", IIf(ImprimirOVComCodProd.Value = vbChecked, 1, 0))
    If lErro <> SUCESSO Then gError 186731

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

    Orcamento_Imprime = SUCESSO

    Exit Function

Erro_Orcamento_Imprime:

    Orcamento_Imprime = gErr

    Select Case gErr

        Case 186729, 186731

        Case 186730
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186732)

    End Select

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function
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
'        If lErro <> SUCESSO Then gError 186733
'
'        dValor = CDbl(PISRetido.Text)
'
'        PISRetido.Text = Format(dValor, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 186734
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
'        Case 186733
'
'        Case 186734
'            Call Rotina_Erro(vbOKOnly, "ERRO_PIS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186735)
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
'        If lErro <> SUCESSO Then gError 186736
'
'        dValor = CDbl(COFINSRetido.Text)
'
'        COFINSRetido.Text = Format(dValor, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 186737
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
'        Case 186736
'
'        Case 186737
'            Call Rotina_Erro(vbOKOnly, "ERRO_COFINS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186738)
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
'        If lErro <> SUCESSO Then gError 186739
'
'        dValor = CDbl(CSLLRetido.Text)
'
'        CSLLRetido.Text = Format(dValor, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 186740
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
'        Case 186739
'
'        Case 186740
'            Call Rotina_Erro(vbOKOnly, "ERRO_CSLL_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186741)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'#############################################################
'INSERIDO POR WAGNER
Public Sub VersaoKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VersaoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub VersaoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub VersaoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = VersaoKit
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub VersaoKitBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VersaoKitBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub VersaoKitBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub VersaoKitBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = VersaoKitBase
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_VersaoKitBase(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit base que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_VersaoKitBase

    Set objGridInt.objControle = VersaoKitBase

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186742

    Saida_Celula_VersaoKitBase = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoKitBase:

    Saida_Celula_VersaoKitBase = gErr

    Select Case gErr

        Case 186742
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186743)

    End Select

End Function

Private Function Saida_Celula_VersaoKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_VersaoKit

    Set objGridInt.objControle = VersaoKit

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186744

    Saida_Celula_VersaoKit = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoKit:

    Saida_Celula_VersaoKit = gErr

    Select Case gErr

        Case 186744
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186745)

    End Select

End Function

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    VersaoKitBase.Text = objKit.sVersao

    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        GridItens.TextMatrix(GridItens.Row, iGrid_VersaoKitBase_Col) = VersaoKitBase.Text
    End If

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186753)

    End Select

    Exit Sub

End Sub

Private Function Atualiza_VersaoKit_Grid() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_VersaoKit_Grid


    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then

        For iIndice = 1 To objGridItens.iLinhasExistentes

            GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col) = Codigo.Text & SEPARADOR & iIndice

        Next

    End If

    Atualiza_VersaoKit_Grid = SUCESSO

    Exit Function

Erro_Atualiza_VersaoKit_Grid:

    Atualiza_VersaoKit_Grid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186754)

    End Select

    Exit Function

End Function

Public Sub NumeroBaseLabel_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As Collection

On Error GoTo Erro_NumeroLabel_Click

    objOrcamentoVenda.lCodigo = StrParaLong(CodigoBase.Text)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("OrcamentoSRV1Lista", colSelecao, objOrcamentoVenda, objEventoNumeroBase)
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186755)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumeroBase_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_objEventoNumeroBase_evSelecao

    Set objOrcamentoVenda = obj1

    lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGOBASE)
    If lErro <> SUCESSO Then gError 186756

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumeroBase_evSelecao:

    Select Case gErr

        Case 186756

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186757)

    End Select

    Exit Sub

End Sub

Public Sub CodigoBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodigoBase_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoBase, iAlterado)

End Sub

Private Function Orcamento_Email(ByVal lOrcamento As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim sMailTo As String, objFilialCliente As New ClassFilialCliente, objEndereco As New ClassEndereco

On Error GoTo Erro_Orcamento_Email

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Guarda no obj o código do orçamento passado como parâmetro
    objOrcamentoVenda.lCodigo = lOrcamento

    'Guarda a FilialEmpresa ativa como filial do orçamento
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    'Lê os dados do orçamento para verificar se o mesmo existe no BD
    lErro = CF("OrcamentoSRV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186534 Then gError 186758

    'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
    If lErro <> SUCESSO Then gError 186759

    objFilialCliente.lCodCliente = objOrcamentoVenda.lCliente
    objFilialCliente.iCodFilial = objOrcamentoVenda.iFilial

    If objFilialCliente.lCodCliente <> 0 And objFilialCliente.iCodFilial Then

        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 186760

        If lErro = SUCESSO Then

            objEndereco.lCodigo = objFilialCliente.lEndereco

            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError 186761

            sMailTo = objEndereco.sEmail

        End If

    End If

    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDiretoEmail("Orçamento de Serviço", "OrcamentoSRV >= @NORCINIC E OrcamentoSRV <= @NORCFIM", 0, "OrcSrv", "NORCINIC", Trim(Codigo.Text), "NORCFIM", Trim(Codigo.Text), "NEXIBIRORC", 1, "TTO_EMAIL", sMailTo, "TSUBJECT", "Orçamento " & CStr(lOrcamento), "TALIASATTACH", "OrcSrv" & CStr(lOrcamento))
    If lErro <> SUCESSO Then gError 186762

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

    Orcamento_Email = SUCESSO

    Exit Function

Erro_Orcamento_Email:

    Orcamento_Email = gErr

    Select Case gErr

        Case 186758

        Case 186760 To 186762

        Case 186759
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_SRV_NAO_CASTRADO", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186763)

    End Select

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function

Public Sub BotaoEmail_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoEmail_Click

    'Se o código do orçamento não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 186764

    'Dispara função para imprimir orçamento
    lErro = Orcamento_Email(Trim(Codigo.Text))
    If lErro <> SUCESSO And lErro <> 102237 Then gError 186765

    Exit Sub

Erro_BotaoEmail_Click:

    Select Case gErr

        Case 186764
            Call Rotina_Erro(vbOKOnly, "'ERRO_NUMERO_ORCAMENTO_NAO_PREENCHIDO", gErr)

        Case 186765

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186766)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object

On Error GoTo Erro_Cliente_Preenche

    Set objcliente = Cliente

    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 186767

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 186767

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187767)

    End Select

    Exit Sub

End Sub
'###############################################################

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIncremento As Integer
Dim objUserControl As Object

    Set objGridInt.objForm = Me
    Set objUserControl = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")

    'Inserido por Wagner
    '########################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colColuna.Add ("Versão do Kit")
        objGridInt.colColuna.Add ("Versao do Kit Base")
    End If
    '########################################

    'precodesc
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    objGridInt.colColuna.Add ("Observação")
    '###############################

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)

    'Inserido por Wagner
    '############################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colCampo.Add (VersaoKit.Name)
        objGridInt.colCampo.Add (VersaoKitBase.Name)

    Else

        VersaoKit.left = POSICAO_FORA_TELA
        VersaoKit.TabStop = False

        VersaoKitBase.left = POSICAO_FORA_TELA
        VersaoKitBase.TabStop = False

    End If
    '############################################

    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    'precodesc
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)

    objGridInt.colCampo.Add (Observacao.Name)
    '###############################

    'Colunas do Grid
    iGrid_ItemProduto_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2

    'Inserido por Wagner
    '#############################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        iGrid_VersaoKit_Col = 3
        iGrid_VersaoKitBase_Col = 4
        iIncremento = iIncremento + 2

    End If
    '#############################################

    iGrid_UnidadeMed_Col = 3 + iIncremento
    iGrid_Quantidade_Col = 4 + iIncremento
    iGrid_PrecoUnitario_Col = 5 + iIncremento
    iGrid_PercDesc_Col = 6 + iIncremento
    iGrid_Desconto_Col = 7 + iIncremento
    iGrid_PrecoTotal_Col = 8 + iIncremento
    iGrid_PrecoTotalB_Col = 9 + iIncremento
    iGrid_Observacao_Col = 10 + iIncremento
    '###############################

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub Status_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub MotivoPerda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub Status_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub MotivoPerda_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_Status(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Status

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSOV, objComboBox)
    If lErro <> SUCESSO Then gError 186769

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0

    Carrega_Status = SUCESSO

    Exit Function

Erro_Carrega_Status:

    Carrega_Status = gErr

    Select Case gErr

        Case 186769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186770)

    End Select

    Exit Function

End Function

Private Function Carrega_Motivo(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Motivo

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_MOTIVOSOV, objComboBox)
    If lErro <> SUCESSO Then gError 186771

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0

    Carrega_Motivo = SUCESSO

    Exit Function

Erro_Carrega_Motivo:

    Carrega_Motivo = gErr

    Select Case gErr

        Case 186771

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186772)

    End Select

    Exit Function

End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_ORCSRV
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = CStr(giFilialEmpresa) & "," & Codigo.Text
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If

    Exit Sub

Erro_Anotacao_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186773)

    End Select

    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao

    Exit Sub

Erro_Anotacao_Preenche:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186774)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Observacao que está deixando de ser a corrente

Dim lErro As Long
Dim sObservacao As String

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186775

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 186775
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186776)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function


Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub




Public Sub GridItensSRV_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensSRV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensSRV, iAlterado)
    End If

End Sub

Public Sub GridItensSRV_EnterCell()

    Call Grid_Entrada_Celula(objGridItensSRV, iAlterado)

End Sub

Public Sub GridItensSRV_GotFocus()

    Call Grid_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub GridItensSRV_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorTotal As Double
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult 'Inserido por Wagner
Dim sServico As String
Dim iPreenchido As Integer

On Error GoTo Erro_GridItensSRV_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItensSRV.iLinhasExistentes
    iItemAtual = GridItensSRV.Row

    lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iItemAtual, iGrid_ServicoSRV_Col), sServico, iPreenchido)
    If lErro <> SUCESSO Then gError 188007

    Call Grid_Trata_Tecla1(KeyCode, objGridItensSRV)

    If objGridItensSRV.iLinhasExistentes < iLinhasExistentesAnterior Then

        '************ grade ************
        'Retira a "#" caso o item excluído tenha sido um de grade
        GridItensSRV.TextMatrix(GridItensSRV.Row, 0) = GridItensSRV.Row
        '*******************************

'        Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)

        Call gobjTribTab.Exclusao_Item_Grid(iItemAtual, True)

        Call ProdSolicSRV_Remover_Item(sServico, gcolProdSolicSRV)
        
        'Calcula a soma dos valores de produtos
        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
            End If
        Next

        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
        Next

        '#################################
        'Inserido por Wagner
        If objGridItensSRV.iLinhasExistentes <> 0 Then
            Call PrecoTotal_CalculaSRV(objGridItensSRV.iLinhasExistentes)
        Else
            If StrParaDbl(ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)

                'Limpa o valor de desconto
                gdDesconto = 0
                ValorDesconto.Text = ""
                giValorDescontoAlterado = 0

                'Para tributação
                gobjOrcamentoVenda.dValorDesconto = 0

            End If
        End If
        '#################################

        'Coloca valor total dos produtos na tela
        ValorProdutos.Caption = Format(dValorTotal, "Standard")

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186229

    End If

    Exit Sub

Erro_GridItensSRV_KeyDown:

    Select Case gErr

        Case 186229, 188007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186230)

    End Select

    Exit Sub

End Sub

Public Sub GridItensSRV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensSRV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensSRV, iAlterado)
    End If


End Sub

Public Sub GridItensSRV_LeaveCell()

    Call Saida_Celula(objGridItensSRV)

End Sub

Public Sub GridItensSRV_RowColChange()

    Call Grid_RowColChange(objGridItensSRV)

End Sub

Public Sub GridItensSRV_Scroll()

    Call Grid_Scroll(objGridItensSRV)

End Sub

Public Sub GridItensSRV_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensSRV)

End Sub

Public Sub PercentDescSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDescSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub PercentDescSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)


End Sub

Public Sub PercentDescSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentDescSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoTotalSRV_Change()

    iAlterado = REGISTRO_ALTERADO


End Sub

Public Sub PrecoTotalSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub PrecoTotalSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub PrecoTotalSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = PrecoTotalSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoUnitarioSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoUnitarioSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub PrecoUnitarioSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub PrecoUnitarioSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = PrecoUnitarioSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ServicoSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ServicoSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub ServicoSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub ServicoSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = ServicoSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantidadeSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantidadeSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub QuantidadeSRV_KeyPress(KeyAscii As Integer)

    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub QuantidadeSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = QuantidadeSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UnidadeMedSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMedSRV_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMedSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub UnidadeMedSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub UnidadeMedSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMedSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Move_GridItensSRV_Memoria(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Move Grid Itens para memória

Dim lErro As Long, iIndice As Integer
Dim sProduto1 As String, sProduto As String, iPreenchido As Integer

On Error GoTo Erro_Move_GridItensSRV_Memoria

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        sProduto1 = GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col)
        
        'Verifica se o Produto está preenchido
        If Len(Trim(sProduto1)) > 0 Then
    
            'Formata o produto
            lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 84102
            
        End If
        
        lErro = Move_GridItem_MemoriaSRV(objOrcamentoVenda, iIndice, sProduto)
        If lErro <> SUCESSO Then gError 186327

    Next

    Move_GridItensSRV_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItensSRV_Memoria:

    Move_GridItensSRV_Memoria = gErr

    Select Case gErr

        Case 186327

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186328)

    End Select

    Exit Function

End Function

Public Function Move_GridItem_MemoriaSRV(objOrcamentoVenda As ClassOrcamentoVenda, iIndice As Integer, sProduto1 As String) As Long
'Recolhe do Grid os dados do item orçamento no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemOrcamento As New ClassItemOV, objTributacaoOV As New ClassTributacaoDoc
Dim iPreenchido As Integer

On Error GoTo Erro_Move_GridItem_MemoriaSRV

    Set objItemOrcamento = New ClassItemOV
'
'    'Verifica se o Produto está preenchido
'    If Len(Trim(sProduto1)) > 0 Then
'
'        'Formata o produto
'        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 186334

    objItemOrcamento.sProduto = sProduto1
'    End If

    If Len(Trim(sProduto1)) = 0 Then gError 186335  'Or iPreenchido = PRODUTO_VAZIO Then gError 186335

    'Armazena os dados do item
    objItemOrcamento.sUnidadeMed = GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col)
    objItemOrcamento.dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))
    objItemOrcamento.dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoUnitarioSRV_Col))
    objItemOrcamento.dPrecoTotal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
    objItemOrcamento.dValorDesconto = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
    objItemOrcamento.dtDataEntrega = StrParaDate(GridItensSRV.TextMatrix(iIndice, iGrid_DataEntregaSRV_Col))
    objItemOrcamento.dValorDesconto = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
    objItemOrcamento.sDescricao = GridItensSRV.TextMatrix(iIndice, iGrid_DescServicoSRV_Col)
    objItemOrcamento.iFilialEmpresa = giFilialEmpresa

    objItemOrcamento.sVersaoKit = GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitSRV_Col)
    objItemOrcamento.sVersaoKitBase = GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitBaseSRV_Col)

    If gobjOrcamentoVenda.colItensSRV.Count >= iIndice Then
        Set objItemOrcamento.objTributacaoItemOV = gobjOrcamentoVenda.colItensSRV.Item(iIndice).objTributacaoItemOV
    Else
        Set objItemOrcamento.objTributacaoItemOV = Nothing
    End If

    objItemOrcamento.sObservacao = GridItensSRV.TextMatrix(iIndice, iGrid_ObservacaoSRV_Col)
    objItemOrcamento.lStatus = LCodigo_Extrai(GridItensSRV.TextMatrix(iIndice, iGrid_StatusItemSRV_Col))
    objItemOrcamento.lMotivoPerda = LCodigo_Extrai(GridItensSRV.TextMatrix(iIndice, iGrid_MotivoPerdaItemSRV_Col))

    'Adiciona o item na colecao de itens do orçamento de venda
     objOrcamentoVenda.colItensSRV.Add objItemOrcamento

    Move_GridItem_MemoriaSRV = SUCESSO

    Exit Function

Erro_Move_GridItem_MemoriaSRV:

    Move_GridItem_MemoriaSRV = gErr

    Select Case gErr

        Case 186334

        Case 186335
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186336)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItensSRV(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensSRV

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Servico
        Case iGrid_ServicoSRV_Col
            lErro = Saida_Celula_ServicoSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186359

        Case iGrid_DescServicoSRV_Col
            lErro = Saida_Celula_DescServicoSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186360

        'Se for a de Unidade de Medida
        Case iGrid_UnidadeMedSRV_Col
            lErro = Saida_Celula_UMSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186361

        'Se for a de Quantidade Pedida
        Case iGrid_QuantidadeSRV_Col
            lErro = Saida_Celula_QuantidadeSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186362

        'Se for a de Preço Unitário
        Case iGrid_PrecoUnitarioSRV_Col
            lErro = Saida_Celula_PrecoUnitarioSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186363

        'Se for a de Percentual de Desconto
        Case iGrid_PercDescSRV_Col
            lErro = Saida_Celula_PercentDescSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186364

        Case iGrid_DataEntregaSRV_Col
            lErro = Saida_Celula_DataEntregaSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186365

        Case iGrid_ObservacaoSRV_Col
            lErro = Saida_Celula_ObservacaoSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186366

        Case iGrid_StatusItemSRV_Col
            lErro = Saida_Celula_StatusItemSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186367

        Case iGrid_MotivoPerdaItemSRV_Col
            lErro = Saida_Celula_MotivoPerdaItemSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186368

        Case iGrid_VersaoKitSRV_Col
            lErro = Saida_Celula_VersaoKitSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186371

        Case iGrid_VersaoKitBaseSRV_Col
            lErro = Saida_Celula_VersaoKitBaseSRV(objGridInt)
            If lErro <> SUCESSO Then gError 186372
            
    End Select

    Saida_Celula_GridItensSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensSRV:

    Saida_Celula_GridItensSRV = gErr

    Select Case gErr

        Case 186359 To 186375, 186387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186376)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_ItensSRV() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_ItensSRV

    'Verifica se há itens no grid
    If objGridItensSRV.iLinhasExistentes = 0 Then gError 186378

    'para cada item do grid
    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col))) = 0 Then gError 186385

        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))) = 0 Then gError 186379

        lErro = Valor_Positivo_Critica(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))
        If lErro <> SUCESSO Then gError 186380

        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoUnitarioSRV_Col))) = 0 Then gError 186381

        lErro = Valor_Positivo_Critica(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoUnitarioSRV_Col))
        If lErro <> SUCESSO Then gError 186382

        If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitSRV_Col))) = 0 And Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitBaseSRV_Col))) <> 0 Then gError 186383
            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitBaseSRV_Col))) = 0 And Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitSRV_Col))) <> 0 Then gError 186384
        End If


    Next

    Valida_Grid_ItensSRV = SUCESSO

    Exit Function

Erro_Valida_Grid_ItensSRV:

    Valida_Grid_ItensSRV = gErr

    Select Case gErr

        Case 186378
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_SERVICOS_NAO_PREENCHIDO", gErr)

        Case 186379
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SERVICO_NAO_PREENCHIDA", gErr, iIndice)

        Case 186380, 186382

        Case 186381
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 186383
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAOKIT_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 186384
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAOKITBASE_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 186385
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_GRID_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186386)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ServicoSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Servico que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_ServicoSRV

    Set objGridInt.objControle = ServicoSRV

    If Len(Trim(ServicoSRV.ClipText)) > 0 Then

        lErro = Servico_Saida_Celula()
        If lErro <> SUCESSO Then gError 186413
        
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186414

    Saida_Celula_ServicoSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_ServicoSRV:

    Saida_Celula_ServicoSRV = gErr

    Select Case gErr

        Case 186413, 186414
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186415)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UMSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidadede Medida que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UMSRV

    Set objGridInt.objControle = UnidadeMedSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186421

    Saida_Celula_UMSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_UMSRV:

    Saida_Celula_UMSRV = gErr

    Select Case gErr

        Case 186421
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186422)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_DescServicoSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do Servico que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescServicoSRV

    Set objGridInt.objControle = DescServicoSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186423

    Saida_Celula_DescServicoSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_DescServicoSRV:

    Saida_Celula_DescServicoSRV = gErr

    Select Case gErr

        Case 186423
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186424)

    End Select

End Function

Private Function Saida_Celula_QuantidadeSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade que está deixando de ser a corrente

Dim lErro As Long
Dim bQuantidadeIgual As Boolean
Dim iIndice As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Saida_Celula_QuantidadeSRV

    Set objGridInt.objControle = QuantidadeSRV

    bQuantidadeIgual = False

    If Len(QuantidadeSRV.Text) > 0 Then

        lErro = Valor_Positivo_Critica(QuantidadeSRV.Text)
        If lErro <> SUCESSO Then gError 186427

        QuantidadeSRV.Text = Formata_Estoque(QuantidadeSRV.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantidadeSRV_Col))
    If dQuantAnterior = StrParaDbl(QuantidadeSRV.Text) Then bQuantidadeIgual = True

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186428

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoUnitarioSRV_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 And Not bQuantidadeIgual Then
    
        Call PrecoTotal_CalculaSRV(GridItensSRV.Row)
    
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186429

    End If

    Saida_Celula_QuantidadeSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeSRV:

    Saida_Celula_QuantidadeSRV = gErr

    Select Case gErr

        Case 186427 To 186429
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186430)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitarioSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoUnitarioSRV

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitarioSRV

    If Len(Trim(PrecoUnitarioSRV.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitarioSRV.Text)
        If lErro <> SUCESSO Then gError 186439

        PrecoUnitarioSRV.Text = Format(PrecoUnitarioSRV.Text, gobjFAT.sFormatoPrecoUnitario)

    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoUnitarioSRV_Col)) = StrParaDbl(PrecoUnitarioSRV.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186440

    If Not bPrecoUnitarioIgual Then

        Call PrecoTotal_CalculaSRV(GridItensSRV.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186441

    End If
    
   Saida_Celula_PrecoUnitarioSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitarioSRV:

    Saida_Celula_PrecoUnitarioSRV = gErr


    Select Case gErr

        Case 186439 To 186441
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186442)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDescSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim sValorPercAnterior As String
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDescSRV

    Set objGridInt.objControle = PercentDescSRV

    dPercentDesc = StrParaDbl(PercentDescSRV.Text)
    
    dPercentDescAnterior = PercentParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PercDescSRV_Col)) * 100

    If Len(PercentDescSRV.Text) > 0 Then
    
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(PercentDescSRV.Text) 'Alterado por Wagner
        If lErro <> SUCESSO Then gError 186443

        dPercentDesc = CDbl(PercentDescSRV.Text)

        If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 186444

            PercentDescSRV.Text = Format(dPercentDesc, "Fixed")

        End If

    Else

        dDesconto = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DescontoSRV_Col))
        dValorTotal = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoTotalSRV_Col))

        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DescontoSRV_Col) = ""
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoTotalSRV_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    sValorPercAnterior = GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PercDescSRV_Col)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186445
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then

        'Recalcula o preço total
        Call PrecoTotal_CalculaSRV(GridItensSRV.Row)

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 186446

        'Preenche GridParcelas a partir da Condição de Pagto
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 186447

    End If

    Saida_Celula_PercentDescSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDescSRV:

    Saida_Celula_PercentDescSRV = gErr

    Select Case gErr

        Case 186443, 186445, 186446
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186444
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186447

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186448)

    End Select

    Exit Function

End Function

Function Saida_Celula_DataEntregaSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Entrega que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataEntrega As Date
Dim dtDataEmissao As Date

On Error GoTo Erro_Saida_Celula_DataEntregaSRV

    Set objGridInt.objControle = DataEntregaSRV

    If Len(Trim(DataEntregaSRV.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(DataEntregaSRV.Text)
        If lErro <> SUCESSO Then gError 186454
        
        'Se data de emissão estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEntrega = CDate(DataEntregaSRV.Text)
            dtDataEmissao = CDate(DataEmissao.Text)
            
            'Veerifica se a data de emissão é maior que a data de entrega
            If dtDataEntrega < dtDataEmissao Then gError 186455

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186456

    Saida_Celula_DataEntregaSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_DataEntregaSRV:

    Saida_Celula_DataEntregaSRV = gErr

    Select Case gErr

        Case 186454, 186456
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186455
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAENTREGA", gErr, dtDataEntrega, dtDataEmissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186457)

    End Select

    Exit Function

End Function

Function Servico_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult
Dim objItemOV As ClassItemOV
Dim iPossuiGrade As Integer
Dim objRomaneioGrade As ClassRomaneioGrade
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProdutoPai As String
Dim objGridItens1 As Object
Dim objKit As New ClassKit 'Inserido por Wagner
Dim sCodigoAntigo As String
Dim sCliente As String
Dim iFilialCli As Integer

On Error GoTo Erro_Servico_Saida_Celula

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", ServicoSRV.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 186458

    If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
        gError 186459
    End If

    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 186460

    If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 188175

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 186461

        ServicoSRV.PromptInclude = False
        ServicoSRV.Text = sProduto
        ServicoSRV.PromptInclude = True

    End If

    'Verifica se é um kit
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then

        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then

            objKit.sProdutoRaiz = objProduto.sCodigo
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 186462

            'Se encontrou => É UM KIT
            If lErro <> 106304 Then
                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_VersaoKitSRV_Col) = Codigo.Text & SEPARADOR & CStr(GridItensSRV.Row)
                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_VersaoKitBaseSRV_Col) = objKit.sVersao
            End If

        End If
    End If

    'Unidade de Medida
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col) = objProduto.sSiglaUMVenda
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DescServicoSRV_Col) = objProduto.sDescricao
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_StatusItemSRV_Col) = Status.Text
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_MotivoPerdaItemSRV_Col) = MotivoPerda.Text

    'Preço Unitário
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ServicoSRV.Text

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 186467

        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 186468

        End If

        If lErro <> 28014 Then

            lErro = PrecoUnitario_Calcula(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 186469

            If dPrecoUnitario > 0 Then
                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoUnitarioSRV_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoUnitarioSRV_Col) = ""
            End If

        End If

    End If

    'precodesc
    Set objGridItens1 = GridItensSRV
    sCliente = Cliente.Text
    iFilialCli = Codigo_Extrai(Filial.Text)
    
    Call CF("Produto_Saida_Celula_PrecoDesc", objGridItens1, GridItensSRV.Row, iGrid_PrecoUnitarioSRV_Col + 1, dPrecoUnitario, objProduto.sCodigo, iGrid_DescServicoSRV_Col, sCliente, iFilialCli, gdDesconto, iGrid_PrecoUnitarioSRV_Col + 2)

    'Acrescenta uma linha no Grid se for o caso
    If GridItensSRV.Row - GridItensSRV.FixedRows = objGridItensSRV.iLinhasExistentes Then
        objGridItensSRV.iLinhasExistentes = objGridItensSRV.iLinhasExistentes + 1

'        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItensSRV.Row)

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItensSRV.Row, objProduto.sCodigo, True)
        If lErro <> SUCESSO Then gError 186470

    End If

    'Atualiza a checkbox do grid para exibir a figura marcada/desmarcada
    Call Grid_Refresh_Checkbox(objGridItensSRV)

    Servico_Saida_Celula = SUCESSO

    Exit Function

Erro_Servico_Saida_Celula:

    Servico_Saida_Celula = gErr

    Select Case gErr

        Case 186461
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, ServicoSRV.Text)

        Case 186458, 186459, 186462, 186467 To 186470

        Case 186460
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ServicoSRV.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItensSRV)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItensSRV)
            End If

        Case 188175
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_NATUREZA_SERVICO", gErr, ServicoSRV.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186471)

    End Select

    Exit Function

End Function

'Private Function Tributacao_Inclusao_Item_GridSRV(iLinha As Integer, sProduto As String) As Long
''trata a inclusao de uma linha de item no grid
'Dim lErro As Long
'Dim objTributoDocItem As ClassTributoDocItem
'Dim objItemOV As ClassItemOV
'
'On Error GoTo Erro_Tributacao_Inclusao_Item_GridSRV
'
'    lErro = Move_GridItem_MemoriaSRV(gobjOrcamentoVenda, iLinha, sProduto)
'    If lErro <> SUCESSO Then gError 186462
'
'    Set objItemOV = gobjOrcamentoVenda.colItensSRV.Item(iLinha)
'
'    Set objTributoDocItem = objItemOV
'
'    lErro = objTributoDocItem.Ativar(gobjOrcamentoVenda)
'    If lErro <> SUCESSO Then gError 186463
'
'    Tributacao_Inclusao_Item_GridSRV = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Inclusao_Item_GridSRV:
'
'    Tributacao_Inclusao_Item_GridSRV = gErr
'
'    Select Case gErr
'
'        Case 186459, 186460
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186464)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub PrecoTotal_CalculaSRV(iLinha As Integer)

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objGridItens1 As Object

On Error GoTo Erro_PrecoTotal_CalculaSRV

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_PrecoUnitarioSRV_Col))
    dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantidadeSRV_Col))

    'Cálculo do desconto
    lTamanho = Len(Trim(GridItensSRV.TextMatrix(iLinha, iGrid_PercDescSRV_Col)))
    If lTamanho > 0 Then
        dPercentDesc = CDbl(Format(GridItensSRV.TextMatrix(iLinha, iGrid_PercDescSRV_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade)

    'Se percentual for >0 tira o desconto
    'If dPercentDesc > 0 Then dDesconto = dPercentDesc * dPrecoTotal
    dDesconto = dPercentDesc * dPrecoTotal
    dPrecoTotalReal = dPrecoTotal - dDesconto 'Inserido por Wagner

    'precodesc
    Set objGridItens1 = GridItensSRV
    Call CF("PrecoTotal_Calcula_PrecoDesc", objGridItens1, iLinha, iGrid_PrecoUnitarioSRV_Col + 1, Format(dPrecoUnitario * (1 - dPercentDesc), "Standard"))

    'Coloca valor do desconto no Grid
    GridItensSRV.TextMatrix(iLinha, iGrid_DescontoSRV_Col) = Format(dDesconto, "Standard")

    'Coloca preco total do ítem no grid
    GridItensSRV.TextMatrix(iLinha, iGrid_PrecoTotalSRV_Col) = Format(dPrecoTotalReal, "Standard")

    GridItensSRV.TextMatrix(iLinha, iGrid_PrecoTotalSRVB_Col) = Format(dPrecoTotal, "Standard")

'    'Calcula a soma dos valores de produtos
'    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'        End If
'    Next
'
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'        End If
'    Next
'
'    If gdDesconto > 0 Then
'        dValorDesconto = gdDesconto * dValorTotal
'    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
'        dValorDesconto = CDbl(ValorDesconto.Text)
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
'        'Para tributação
'        gobjOrcamentoVenda.dValorDesconto = dValorDesconto
'
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 186487
'
'        'Calcula a soma dos valores de produtos
'        dValorTotal = 0
'        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'            End If
'        Next
'
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'            End If
'        Next
'
'    End If
'
'    'Coloca valor total dos produtos na tela
'    ValorProdutos.Caption = Format(dValorTotal, "Standard")
'    ValorDesconto.Text = Format(dValorDesconto, "Standard")

    Call SubTotal_Calcula

    lErro = gobjTribTab.Alteracao_Item_Grid(iLinha, True)
    If lErro <> SUCESSO Then gError 186429
        
    Exit Sub

Erro_PrecoTotal_CalculaSRV:

    Select Case gErr

        Case 186487

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186488)

    End Select

    Exit Sub

End Sub
'
'Function Tributacao_Alteracao_Item_GridSRV(iIndice As Integer) As Long
''trata a alteracao de uma linha de item no grid
'
'Dim lErro As Long, sProduto As String, iPreenchido As Integer
'Dim objItemOV As ClassItemOV
'
'On Error GoTo Erro_Tributacao_Alteracao_Item_GridSRV
'
'    Set objItemOV = gobjOrcamentoVenda.colItensSRV.Item(iIndice)
'
'    If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col))) > 0 Then
'
'        lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col), sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 186483
'
'        objItemOV.sProduto = sProduto
'
'    End If
'
'    objItemOV.sUnidadeMed = GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col)
'
'    objItemOV.dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))
'    'objItemOV.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
'    objItemOV.dPrecoTotal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'    objItemOV.dValorDesconto = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
'
'    Tributacao_Alteracao_Item_GridSRV = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Alteracao_Item_GridSRV:
'
'    Tributacao_Alteracao_Item_GridSRV = gErr
'
'    Select Case gErr
'
'        Case 186483
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186484)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_Grid_ItensSRV(objOrcamentoVenda As ClassOrcamentoVenda, ByVal iBrowse As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objGridItens1 As Object

On Error GoTo Erro_Carrega_Grid_ItensSRV

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItensSRV)

    For iIndice = 1 To objOrcamentoVenda.colItensSRV.Count

        objOrcamentoVenda.colItensSRV(iIndice).iItem = iIndice

        lErro = Mascara_RetornaProdutoEnxuto(objOrcamentoVenda.colItensSRV(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 186651

        'Mascara o produto enxuto
        ServicoSRV.PromptInclude = False
        ServicoSRV.Text = sProdutoEnxuto
        ServicoSRV.PromptInclude = True

        'Calcula o percentual de desconto
        If objOrcamentoVenda.colItensSRV(iIndice).dPrecoTotal + objOrcamentoVenda.colItensSRV(iIndice).dValorDesconto > 0 Then
            dPercDesc = objOrcamentoVenda.colItensSRV(iIndice).dValorDesconto / (objOrcamentoVenda.colItensSRV(iIndice).dPrecoTotal + objOrcamentoVenda.colItensSRV(iIndice).dValorDesconto)
        End If

        'Coloca os dados dos itens na tela
        GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col) = ServicoSRV.Text
        GridItensSRV.TextMatrix(iIndice, iGrid_DescServicoSRV_Col) = objOrcamentoVenda.colItensSRV(iIndice).sDescricao
        GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col) = objOrcamentoVenda.colItensSRV(iIndice).sUnidadeMed
        GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col) = Formata_Estoque(objOrcamentoVenda.colItensSRV(iIndice).dQuantidade)
        GridItensSRV.TextMatrix(iIndice, iGrid_PrecoUnitarioSRV_Col) = Format(objOrcamentoVenda.colItensSRV(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)

        '###############################################
        'INSERIDO POR WAGNER
        If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
            If iBrowse = VALOR_CODIGO Then
                GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitSRV_Col) = objOrcamentoVenda.colItensSRV(iIndice).sVersaoKit
                GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitBaseSRV_Col) = objOrcamentoVenda.colItensSRV(iIndice).sVersaoKitBase
            Else
                GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitSRV_Col) = Codigo.Text & SEPARADOR & iIndice
                GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitBaseSRV_Col) = objOrcamentoVenda.colItensSRV(iIndice).sVersaoKit
            End If
        End If
        '###############################################

        'precoDesc
        Set objGridItens1 = GridItensSRV
        Call CF("Carrega_Grid_Itens_PrecoDesc", objGridItens1, iIndice, iGrid_PrecoUnitarioSRV_Col + 1, Format(objOrcamentoVenda.colItensSRV(iIndice).dPrecoUnitario * (1 - dPercDesc), gobjFAT.sFormatoPrecoUnitario), objOrcamentoVenda.colItensSRV(iIndice).sProduto, iGrid_DescServicoSRV_Col)

        GridItensSRV.TextMatrix(iIndice, iGrid_PercDescSRV_Col) = Format(dPercDesc, "Percent")
        GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col) = Format(objOrcamentoVenda.colItensSRV(iIndice).dValorDesconto, "Standard")
        GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col) = Format(objOrcamentoVenda.colItensSRV(iIndice).dPrecoTotal, "Standard")
        GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col) = Format(objOrcamentoVenda.colItensSRV(iIndice).dPrecoTotal + objOrcamentoVenda.colItensSRV(iIndice).dValorDesconto, "Standard")
        If objOrcamentoVenda.colItensSRV(iIndice).dtDataEntrega <> DATA_NULA Then GridItensSRV.TextMatrix(iIndice, iGrid_DataEntregaSRV_Col) = Format(objOrcamentoVenda.colItensSRV(iIndice).dtDataEntrega, "dd/mm/yyyy")

        GridItensSRV.TextMatrix(iIndice, iGrid_ObservacaoSRV_Col) = objOrcamentoVenda.colItensSRV(iIndice).sObservacao
        If objOrcamentoVenda.colItensSRV(iIndice).lStatus <> 0 Then
            Call Combo_Seleciona_ItemData(StatusItemSRV, objOrcamentoVenda.colItensSRV(iIndice).lStatus)
            GridItensSRV.TextMatrix(iIndice, iGrid_StatusItemSRV_Col) = StatusItemSRV.Text
        End If
        If objOrcamentoVenda.colItensSRV(iIndice).lMotivoPerda <> 0 Then
            Call Combo_Seleciona_ItemData(MotivoPerdaItemSRV, objOrcamentoVenda.colItensSRV(iIndice).lMotivoPerda)
            GridItensSRV.TextMatrix(iIndice, iGrid_MotivoPerdaItemSRV_Col) = MotivoPerdaItemSRV.Text
        End If

    Next

    'Atualiza o número de linhas existentes
    objGridItensSRV.iLinhasExistentes = objOrcamentoVenda.colItensSRV.Count

    'Atualiza o grid para mostrar as checkboxes marcadas / desmarcadas
    Call Grid_Refresh_Checkbox(objGridItensSRV)

    Carrega_Grid_ItensSRV = SUCESSO

    Exit Function

Erro_Carrega_Grid_ItensSRV:

    Carrega_Grid_ItensSRV = gErr

    Select Case gErr

        Case 186651
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objOrcamentoVenda.colItensSRV(iIndice).sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186652)

    End Select

    Exit Function

End Function

Public Sub VersaoKitSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VersaoKitSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub VersaoKitSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub VersaoKitSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = VersaoKitSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub VersaoKitBaseSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VersaoKitBaseSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub VersaoKitBaseSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub VersaoKitBaseSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = VersaoKitBaseSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_VersaoKitBaseSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit base que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_VersaoKitBaseSRV

    Set objGridInt.objControle = VersaoKitBaseSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186746

    Saida_Celula_VersaoKitBaseSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoKitBaseSRV:

    Saida_Celula_VersaoKitBaseSRV = gErr

    Select Case gErr

        Case 186746
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186747)

    End Select

End Function

Private Function Saida_Celula_VersaoKitSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_VersaoKitSRV

    Set objGridInt.objControle = VersaoKitSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186748

    Saida_Celula_VersaoKitSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoKitSRV:

    Saida_Celula_VersaoKitSRV = gErr

    Select Case gErr

        Case 186748
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186749)

    End Select

End Function

Private Sub objEventoVersaoSRV_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersaoSRV_evSelecao

    Set objKit = obj1

    VersaoKitBaseSRV.Text = objKit.sVersao

    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_VersaoKitBaseSRV_Col) = VersaoKitBaseSRV.Text
    End If

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersaoSRV_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186750)

    End Select

    Exit Sub

End Sub

Private Function Atualiza_VersaoKitSRV_Grid() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_VersaoKitSRV_Grid


    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then

        For iIndice = 1 To objGridItensSRV.iLinhasExistentes

            GridItensSRV.TextMatrix(iIndice, iGrid_VersaoKitSRV_Col) = Codigo.Text & SEPARADOR & iIndice

        Next

    End If

    Atualiza_VersaoKitSRV_Grid = SUCESSO

    Exit Function

Erro_Atualiza_VersaoKitSRV_Grid:

    Atualiza_VersaoKitSRV_Grid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186751)

    End Select

    Exit Function

End Function

Public Function Inicializa_Grid_ItensSRV(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIncremento As Integer
Dim objUserControl As Object

    Set objGridInt.objForm = Me
    Set objUserControl = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Descrição")

    'Inserido por Wagner
    '########################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colColuna.Add ("Versão do Kit")
        objGridInt.colColuna.Add ("Versao do Kit Base")
    End If
    '########################################

    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Motivo Perda")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ServicoSRV.Name)
    objGridInt.colCampo.Add (DescServicoSRV.Name)

    'Inserido por Wagner
    '############################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colCampo.Add (VersaoKitSRV.Name)
        objGridInt.colCampo.Add (VersaoKitBaseSRV.Name)

    Else

        VersaoKitSRV.left = POSICAO_FORA_TELA
        VersaoKitSRV.TabStop = False

        VersaoKitBaseSRV.left = POSICAO_FORA_TELA
        VersaoKitBaseSRV.TabStop = False

    End If
    '############################################

    objGridInt.colCampo.Add (UnidadeMedSRV.Name)
    objGridInt.colCampo.Add (QuantidadeSRV.Name)
    objGridInt.colCampo.Add (PrecoUnitarioSRV.Name)
    'precodesc
    objGridInt.colCampo.Add (PercentDescSRV.Name)
    objGridInt.colCampo.Add (DescontoSRV.Name)
    objGridInt.colCampo.Add (PrecoTotalSRV.Name)
    objGridInt.colCampo.Add (PrecoTotalSRVB.Name)
    objGridInt.colCampo.Add (DataEntregaSRV.Name)

    '###############################
    'Inserido por Wagner 09/12/05
    objGridInt.colCampo.Add (StatusItemSRV.Name)
    objGridInt.colCampo.Add (MotivoPerdaItemSRV.Name)
    objGridInt.colCampo.Add (ObservacaoSRV.Name)
    '###############################

    'Colunas do Grid
    iGrid_ItemServicoSRV_Col = 0
    iGrid_ServicoSRV_Col = 1
    iGrid_DescServicoSRV_Col = 2

    'Inserido por Wagner
    '#############################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        iGrid_VersaoKitSRV_Col = 3
        iGrid_VersaoKitBaseSRV_Col = 4
        iIncremento = iIncremento + 2

    End If
    '#############################################

    iGrid_UnidadeMedSRV_Col = 3 + iIncremento
    iGrid_QuantidadeSRV_Col = 4 + iIncremento
    iGrid_PrecoUnitarioSRV_Col = 5 + iIncremento
    iGrid_PercDescSRV_Col = 6 + iIncremento
    iGrid_DescontoSRV_Col = 7 + iIncremento
    iGrid_PrecoTotalSRV_Col = 8 + iIncremento
    iGrid_PrecoTotalSRVB_Col = 9 + iIncremento
    iGrid_DataEntregaSRV_Col = 10 + iIncremento

    iGrid_StatusItemSRV_Col = 11 + iIncremento
    iGrid_MotivoPerdaItemSRV_Col = 12 + iIncremento
    iGrid_ObservacaoSRV_Col = 13 + iIncremento
    '###############################

    'Grid do GridInterno
    objGridInt.objGrid = GridItensSRV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItensSRV.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensSRV = SUCESSO

    Exit Function

End Function

Private Function Saida_Celula_MotivoPerdaItemSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula MotivoPerdaItem que está deixando de ser a corrente

Dim lErro As Long
Dim sMotivoPerdaItem As String

On Error GoTo Erro_Saida_Celula_MotivoPerdaItemSRV

    Set objGridInt.objControle = MotivoPerdaItemSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186778

    Saida_Celula_MotivoPerdaItemSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoPerdaItemSRV:

    Saida_Celula_MotivoPerdaItemSRV = gErr

    Select Case gErr

        Case 186778
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186779)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_StatusItemSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula StatusItem que está deixando de ser a corrente

Dim lErro As Long
Dim sStatusItem As String

On Error GoTo Erro_Saida_Celula_StatusItemSRV

    Set objGridInt.objControle = StatusItemSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186780

    Saida_Celula_StatusItemSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_StatusItemSRV:

    Saida_Celula_StatusItemSRV = gErr

    Select Case gErr

        Case 186780
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186781)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub ObservacaoSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ObservacaoSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub ObservacaoSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub ObservacaoSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = ObservacaoSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ObservacaoSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Observacao que está deixando de ser a corrente

Dim lErro As Long
Dim sObservacao As String

On Error GoTo Erro_Saida_Celula_ObservacaoSRV

    Set objGridInt.objControle = ObservacaoSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186789

    Saida_Celula_ObservacaoSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoSRV:

    Saida_Celula_ObservacaoSRV = gErr

    Select Case gErr

        Case 186789
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186790)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub StatusItemSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub StatusItemSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub StatusItemSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub StatusItemSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = StatusItemSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MotivoPerdaItemSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub MotivoPerdaItemSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub MotivoPerdaItemSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub MotivoPerdaItemSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = MotivoPerdaItemSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoPecas_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoPecas_Click

    If Me.ActiveControl Is Produto Then

        sProduto1 = Produto.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 186782

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 186783

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza<>?"

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto, sSelecaoSQL)

    Exit Sub

Erro_BotaoPecas_Click:

    Select Case gErr

        Case 186782
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 186783

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186784)

    End Select

    Exit Sub

End Sub


Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 186785

    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

    'Faz o Tratamento do produto
    lErro = Produto_Saida_Celula()
    If lErro <> SUCESSO Then


        If Not (Me.ActiveControl Is Produto) Then

            GridItens.TextMatrix(GridItensSRV.Row, iGrid_Produto_Col) = ""

        End If

        gError 186786
        
    End If

    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 186785
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 186786

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186787)

    End Select

    Exit Sub

End Sub

Public Sub DataEntregaSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEntregaSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub DataEntregaSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub DataEntregaSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataEntregaSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Private Sub BotaoProjetos_Click()
        Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Sub

Private Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Private Sub Projeto_GotFocus()
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Sub

Private Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        Set objItensPRJCR = New ClassItensPRJCR

        objItensPRJCR.sItem = CStr(iIndice)
        objItensPRJCR.sDescricao = GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_ORCSRV

        colItensPRJ.Add objItensPRJCR

    Next

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItensPRJCR = New ClassItensPRJCR

        objItensPRJCR.sItem = CStr(iIndice + objGridItensSRV.objGrid.Count)
        objItensPRJCR.sDescricao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_ORCSRV

        colItensPRJ.Add objItensPRJCR

    Next

    Call gobjTribTab.Preenche_Itens_Projeto_NF(colItensPRJ)

    Obter_ItensPRJCR = SUCESSO

    Exit Function

Erro_Obter_ItensPRJCR:

    Obter_ItensPRJCR = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186788)

    End Select

    Exit Function
    
End Function

Public Function Preenche_Itens_Projeto_Aux(ByVal colItensPRJ As Collection, ByVal sItem As String, ByVal sDescricao As String, ByVal dValor As Double, ByVal iTipo As Integer, Optional ByVal iNegativo As Integer = 0, Optional ByVal sOBS As String = "") As Long

Dim objItensPRJCR As New ClassItensPRJCR

On Error GoTo Erro_Preenche_Itens_Projeto_Aux

    objItensPRJCR.sItem = sItem
    objItensPRJCR.sDescricao = sDescricao
    objItensPRJCR.dValorOriginal = dValor
    objItensPRJCR.iNegativo = iNegativo
    objItensPRJCR.iTipoOrigem = iTipo
    objItensPRJCR.sObservacao = sOBS

    colItensPRJ.Add objItensPRJCR

    Exit Function

Erro_Preenche_Itens_Projeto_Aux:

    Preenche_Itens_Projeto_Aux = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
    
End Function
'#####################################

Public Function Traz_Dados_SolicSRV_Tela(objSolicSRV As ClassSolicSRV) As Long

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim objItensSolicSRV As ClassItensSolicSRV
Dim sServicoEnxuto As String
Dim iIndice As Integer
Dim objProdSolicSRV As ClassProdSolicSRV
Dim colItens As New Collection
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim dPrecoUnitario As Double

On Error GoTo Erro_Traz_Dados_SolicSRV_Tela

    'Limpa a tela
    Call Limpa_OrcamentoVenda2

    If objSolicSRV.iVendedor <> 0 Then
        
        objVendedor.iCodigo = objSolicSRV.iVendedor
            
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 188003
        
        If lErro <> SUCESSO Then gError 188004
    
        Vendedor.Text = objVendedor.sNomeReduzido
        
    End If
        
    'Se existe um código para o Cliente
    If objSolicSRV.lCliente <> 0 Then

        Call Cliente_Formata(objSolicSRV.lCliente)
        Call Filial_Formata(Filial, objSolicSRV.iFilial)
        Call Cliente_Validate(bSGECancelDummy)
        Call Filial_Validate(bSGECancelDummy)

    End If
        
    CodSolicSRV.Text = objSolicSRV.lCodigo
    
    lErro = CF("Solicitacao_Le_Gastos", objSolicSRV, colItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iIndice = 0
    For Each objItensSolicSRV In colItens
    
        If objItensSolicSRV.iTipo = ITEMNF_TIPO_SERVICO Then
    
            lErro = Mascara_RetornaProdutoEnxuto(objItensSolicSRV.sProduto, sServicoEnxuto)
            If lErro <> SUCESSO Then gError 188005
    
            'Mascara o produto enxuto
            ServicoSRV.PromptInclude = False
            ServicoSRV.Text = sServicoEnxuto
            ServicoSRV.PromptInclude = True
    
            iIndice = iIndice + 1
    
            'Coloca os dados dos itens na tela
            GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col) = ServicoSRV.Text
            GridItensSRV.TextMatrix(iIndice, iGrid_DescServicoSRV_Col) = objItensSolicSRV.sProdutoDesc
            GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col) = objItensSolicSRV.sUM
            GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col) = Formata_Estoque(objItensSolicSRV.dQuantidade)
    
'            Set objProdSolicSRV = New ClassProdSolicSRV
'
'            objProdSolicSRV.sServicoOrcSRV = objItensSolicSRV.sServico
'            objProdSolicSRV.dQuantidade = objItensSolicSRV.dQuantidade
'            objProdSolicSRV.sProduto = objItensSolicSRV.sProduto
'            objProdSolicSRV.sContrato = objItensSolicSRV.sContrato
'            objProdSolicSRV.lGarantia = objItensSolicSRV.lGarantia
'            objProdSolicSRV.sLote = objItensSolicSRV.sLote
'            objProdSolicSRV.iFilialOP = objItensSolicSRV.iFilialOP
'
'            gcolProdSolicSRV.Add objProdSolicSRV
        
            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objItensSolicSRV.sProduto, True)
            If lErro <> SUCESSO Then gError 195756
            
            'Preço Unitário
            If Len(Trim(TabelaPreco.Text)) > 0 Then
            
                lErro = Trata_TabelaPreco_Item_SRV(iIndice)
                If lErro <> SUCESSO Then gError 186290
        
            End If
            
        End If
    
    Next
    objGridItensSRV.iLinhasExistentes = iIndice

    
    iIndice = 0
    For Each objItensSolicSRV In colItens
    
        If objItensSolicSRV.iTipo = ITEMNF_TIPO_PECA Then
    
            lErro = Mascara_RetornaProdutoEnxuto(objItensSolicSRV.sProduto, sServicoEnxuto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Mascara o produto enxuto
            Produto.PromptInclude = False
            Produto.Text = sServicoEnxuto
            Produto.PromptInclude = True
    
            iIndice = iIndice + 1
    
            'Coloca os dados dos itens na tela
            GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
            GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItensSolicSRV.sProdutoDesc
            GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItensSolicSRV.sUM
            GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItensSolicSRV.dQuantidade)

'            Set objProdSolicSRV = New ClassProdSolicSRV
'
'            objProdSolicSRV.sServicoOrcSRV = objItensSolicSRV.sServico
'            objProdSolicSRV.dQuantidade = objItensSolicSRV.dQuantidade
'            objProdSolicSRV.sProduto = objItensSolicSRV.sProduto
'            objProdSolicSRV.sContrato = objItensSolicSRV.sContrato
'            objProdSolicSRV.lGarantia = objItensSolicSRV.lGarantia
'            objProdSolicSRV.sLote = objItensSolicSRV.sLote
'            objProdSolicSRV.iFilialOP = objItensSolicSRV.iFilialOP
'
'            gcolProdSolicSRV.Add objProdSolicSRV

            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objItensSolicSRV.sProduto)
            If lErro <> SUCESSO Then gError 195756
            
                'Preço Unitário
            If Len(Trim(TabelaPreco.Text)) > 0 Then
            
                lErro = Trata_TabelaPreco_Item(iIndice)
                If lErro <> SUCESSO Then gError 186290
        
            End If
            
        End If
    
    Next
        
    objGridItens.iLinhasExistentes = iIndice
        
    Traz_Dados_SolicSRV_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_SolicSRV_Tela:

    'Indica que o sistema não está mais carregando dados para a tela
    gbCarregandoTela = False
    
    Traz_Dados_SolicSRV_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 188003, 195756

        Case 188004
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)

        Case 188005
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItensSolicSRV.sServico)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188006)

    End Select

    Exit Function

End Function

Sub ProdSolicSRV_Remover_Item(ByVal sServico As String, ByVal colProdSolicSRV As Collection)

Dim objProdSolicSRV As ClassProdSolicSRV
Dim iIndice As Integer

On Error GoTo Erro_ProdSolicSRV_Remover_Item

    For iIndice = colProdSolicSRV.Count To 1 Step -1
    
        Set objProdSolicSRV = colProdSolicSRV.Item(iIndice)
    
        If objProdSolicSRV.sServicoOrcSRV = sServico Then
            colProdSolicSRV.Remove (iIndice)
        End If
    
    Next

    Exit Sub

Erro_ProdSolicSRV_Remover_Item:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188008)

    End Select

End Sub

'Sub ProdutoSRV_Remover_Servico(ByVal sServico As String, ByVal colProdutoSRV As Collection)
'
'Dim objProdutoSRV As ClassProdutoSRV
'Dim iIndice As Integer
'
'On Error GoTo Erro_ProdutoSRV_Remover_Servico
'
'    For iIndice = colProdutoSRV.Count To 1 Step -1
'
'        Set objProdutoSRV = colProdutoSRV.Item(iIndice)
'
'        If objProdutoSRV.sServicoSRV = sServico Then
'            colProdutoSRV.Remove (iIndice)
'        End If
'
'    Next
'
'    Exit Sub
'
'Erro_ProdutoSRV_Remover_Servico:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188008)
'
'    End Select
'
'End Sub

'Sub ProdutoSRV_Remover_Produto(ByVal sProduto As String, ByVal colProdutoSRV As Collection)
'
'Dim objProdutoSRV As ClassProdutoSRV
'Dim iIndice As Integer
'
'On Error GoTo Erro_ProdutoSRV_Remover_Produto
'
'    For iIndice = colProdutoSRV.Count To 1 Step -1
'
'        Set objProdutoSRV = colProdutoSRV.Item(iIndice)
'
'        If objProdutoSRV.sProdutoSRV = sProduto Then
'            colProdutoSRV.Remove (iIndice)
'        End If
'
'    Next
'
'    Exit Sub
'
'Erro_ProdutoSRV_Remover_Produto:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188008)
'
'    End Select
'
'End Sub


Public Function Atualiza_Tela() As Long
'atualiza a tela a partir dos dados oriundos da chamada de ProdSolicSRV

Dim objItemOVServico As ClassItemOV
Dim colItemOVServico As New Collection
Dim objItemOVPeca As ClassItemOV
Dim colItemOVPeca As New Collection
Dim iAchou As Integer
Dim objProdutoSRV As ClassProdutoSRV
Dim objProdSolicSRV As ClassProdSolicSRV
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sServico As String
Dim iPreenchido As Integer
Dim dQuantidade As Double
Dim dPrecoUnitario As Double
Dim sPeca As String
Dim lErro As Long
Dim s
Dim sServicoEnxuto As String
Dim sProdutoEnxuto As String

On Error GoTo Erro_Atualiza_Tela

    lErro = Totaliza_Servicos_Pecas(colItemOVServico, colItemOVPeca)
    If lErro <> SUCESSO Then gError 188229
    
    'verifica se alguma linha do grid de servico tem a quantidade menor do que a apurada pela tela ProdSolicSRV
    'se for menor, coloca o quantidade maior e atualiza os valores
    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        For iIndice1 = colItemOVServico.Count To 1 Step -1
        
            Set objItemOVServico = colItemOVServico.Item(iIndice1)
        
            lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col), sServico, iPreenchido)
            If lErro <> SUCESSO Then gError 188208
        
            If objItemOVServico.sProduto = sServico Then
            
                dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))
                
                If objItemOVServico.dQuantidade > dQuantidade Then
                    
                    GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col) = Formata_Estoque(objItemOVServico.dQuantidade)
                                                
                    'Preço unitário
                    dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoUnitarioSRV_Col))
                
                    'Recalcula preço do ítem e valor total da nota
                    If dPrecoUnitario > 0 Then
                        Call PrecoTotal_CalculaSRV(iIndice)
                        lErro = ValorTotal_Calcula()
                        If lErro <> SUCESSO Then gError 188209
                    End If
                                                
                End If
                
                colItemOVServico.Remove (iIndice1)
                Exit For
                
            End If
            
        Next
        
    Next
    
    'para cada um dos servicos que nao foram encontrados no grid de servicos ==> adiciona-os ao grid
    For Each objItemOVServico In colItemOVServico
    
        GridItensSRV.Row = objGridItensSRV.iLinhasExistentes + 1
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemOVServico.sProduto, sServicoEnxuto)
        If lErro <> SUCESSO Then gError 188240

        'Mascara o produto enxuto
        ServicoSRV.PromptInclude = False
        ServicoSRV.Text = sServicoEnxuto
        ServicoSRV.PromptInclude = True
        
        lErro = Servico_Saida_Celula()
        If lErro <> SUCESSO Then gError 188210
    
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ServicoSRV.Text
    
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantidadeSRV_Col) = Formata_Estoque(objItemOVServico.dQuantidade)
    
        'Preço unitário
        dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItensSRV.Row, iGrid_PrecoUnitarioSRV_Col))
    
        'Recalcula preço do ítem e valor total da nota
        If dPrecoUnitario > 0 Then
            Call PrecoTotal_CalculaSRV(GridItensSRV.Row)
            lErro = ValorTotal_Calcula()
            If lErro <> SUCESSO Then gError 188211
        End If
    
    Next
        
    'verifica se alguma linha do grid de pecas tem a quantidade menor do que a apurada pela tela ProdSolicSRV
    'se for menor, coloca o quantidade maior e atualiza os valores
    For iIndice = 1 To objGridItens.iLinhasExistentes

        For iIndice1 = colItemOVPeca.Count To 1 Step -1
        
            Set objItemOVPeca = colItemOVPeca.Item(iIndice1)
        
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sPeca, iPreenchido)
            If lErro <> SUCESSO Then gError 188212
        
            If objItemOVPeca.sProduto = sPeca Then
            
                dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
                
                If objItemOVPeca.dQuantidade > dQuantidade Then
                    
                    GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemOVPeca.dQuantidade)
                                                
                    'Preço unitário
                    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
                
                    'Recalcula preço do ítem e valor total da nota
                    If dPrecoUnitario > 0 Then
                        Call PrecoTotal_Calcula(iIndice)
                        lErro = ValorTotal_Calcula()
                        If lErro <> SUCESSO Then gError 188213
                    End If
                                                
                End If
                
                colItemOVPeca.Remove (iIndice1)
                Exit For
                
            End If
            
        Next
        
    Next
    
    'para cada uma das pecas que nao foram encontrados no grid de pecas ==> adiciona-as ao grid
    For Each objItemOVPeca In colItemOVPeca
    
        GridItens.Row = objGridItens.iLinhasExistentes + 1
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemOVPeca.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 188241
        
        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 188214
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(objItemOVPeca.dQuantidade)
    
        'Preço unitário
        dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
    
        If dPrecoUnitario > 0 Then
            Call PrecoTotal_Calcula(GridItens.Row)
            lErro = ValorTotal_Calcula()
            If lErro <> SUCESSO Then gError 188215
        End If
    
    Next

    Atualiza_Tela = SUCESSO

    Exit Function

Erro_Atualiza_Tela:

    Atualiza_Tela = gErr

    Select Case gErr

        Case 188208 To 188215, 188229

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188216)

    End Select

    Exit Function

End Function

Private Function Totaliza_Servicos_Pecas(colItemOVServico As Collection, colItemOVPeca) As Long
'totaliza os servicos e pecas contidos em gcolProdSolicSRV

Dim objItemOVServico As ClassItemOV
Dim objItemOVPeca As ClassItemOV
Dim iAchou As Integer
Dim objProdutoSRV As ClassProdutoSRV
Dim objProdSolicSRV As ClassProdSolicSRV

On Error GoTo Erro_Totaliza_Servicos_Pecas

    'totaliza os servicos
    For Each objProdSolicSRV In gcolProdSolicSRV
                        
        iAchou = 0
                
        For Each objItemOVServico In colItemOVServico
        
            If objItemOVServico.sProduto = objProdSolicSRV.sServicoOrcSRV Then
                objItemOVServico.dQuantidade = objItemOVServico.dQuantidade + objProdSolicSRV.dQuantidade
                iAchou = 1
                Exit For
            End If
    
        Next
        
        If iAchou = 0 Then
        
            Set objItemOVServico = New ClassItemOV
            
            objItemOVServico.sProduto = objProdSolicSRV.sServicoOrcSRV
            objItemOVServico.dQuantidade = objProdSolicSRV.dQuantidade
            
            colItemOVServico.Add objItemOVServico
            
        End If
        
        'totaliza as pecas associadas ao servico
        For Each objProdutoSRV In objProdSolicSRV.colProdutoSRV
        
            iAchou = 0
        
            For Each objItemOVPeca In colItemOVPeca
            
                If objItemOVPeca.sProduto = objProdutoSRV.sProduto Then
                    objItemOVPeca.dQuantidade = objItemOVPeca.dQuantidade + objProdutoSRV.dQuantidade
                    iAchou = 1
                    Exit For
                End If
        
            Next
            
            If iAchou = 0 Then
            
                Set objItemOVPeca = New ClassItemOV
                
                objItemOVPeca.sProduto = objProdutoSRV.sProduto
                objItemOVPeca.dQuantidade = objProdutoSRV.dQuantidade
                
                colItemOVPeca.Add objItemOVPeca
                
            End If
    
        Next
    
    Next

    Totaliza_Servicos_Pecas = SUCESSO

    Exit Function

Erro_Totaliza_Servicos_Pecas:

    Totaliza_Servicos_Pecas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188228)

    End Select

    Exit Function

End Function

Private Sub ProdutoSRV_Remover_Item(ByVal sPeca As String, ByVal colProdSolicSRV As Collection)

Dim objProdSolicSRV As ClassProdSolicSRV
Dim iIndice As Integer
Dim objProdutoSRV As ClassProdutoSRV

On Error GoTo Erro_ProdutoSRV_Remover_Item

    For Each objProdSolicSRV In colProdSolicSRV
    
        For iIndice = objProdSolicSRV.colProdutoSRV.Count To 1 Step -1
            
            Set objProdutoSRV = objProdSolicSRV.colProdutoSRV.Item(iIndice)
    
            If objProdutoSRV.sProduto = sPeca Then colProdSolicSRV.Remove (iIndice)
        
        Next
    
    Next

    Exit Sub

Erro_ProdutoSRV_Remover_Item:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188227)

    End Select

End Sub

Private Function Valida_Totais_Servicos_Pecas() As Long
'testa se os totais de servicos e pecas contidos em gcolProdSolicSRV esta de acordo com os grids de servicos e pecas

Dim objItemOVServico As ClassItemOV
Dim colItemOVServico As New Collection
Dim objItemOVPeca As ClassItemOV
Dim colItemOVPeca As New Collection
Dim iAchou As Integer
Dim objProdutoSRV As ClassProdutoSRV
Dim objProdSolicSRV As ClassProdSolicSRV
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sServico As String
Dim iPreenchido As Integer
Dim dQuantidade As Double
Dim dPrecoUnitario As Double
Dim sPeca As String
Dim lErro As Long
Dim sMsg As String

On Error GoTo Erro_Valida_Totais_Servicos_Pecas

    lErro = Totaliza_Servicos_Pecas(colItemOVServico, colItemOVPeca)
    If lErro <> SUCESSO Then gError 188230
    
    'verifica se alguma linha do grid de servico tem a quantidade diferente do que a apurada pela tela ProdSolicSRV
    'se for coloca a informacao na mensagem
    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        For iIndice1 = colItemOVServico.Count To 1 Step -1
        
            Set objItemOVServico = colItemOVServico.Item(iIndice1)
        
            lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col), sServico, iPreenchido)
            If lErro <> SUCESSO Then gError 188231
        
            If objItemOVServico.sProduto = sServico Then
            
                dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantidadeSRV_Col))
                
                If objItemOVServico.dQuantidade <> dQuantidade Then
                    
                    sMsg = sMsg & "Serviço: " & sServico & " Quant.Grid Itens: " & CStr(dQuantidade) & " Quant. Detalhada: " & CStr(objItemOVServico.dQuantidade) & Chr(10)
                                                
                End If
                
                colItemOVServico.Remove (iIndice1)
                
                Exit For
                
            End If
            
        Next
        
    Next
    
    'para cada um dos servicos que nao foram encontrados no grid de servicos ==> adiciona-os na msg
    For Each objItemOVServico In colItemOVServico
    
        sMsg = sMsg & "Serviço: " & objItemOVServico.sProduto & " Não está no grid de itens e Quant. Detalhada: " & CStr(objItemOVServico.dQuantidade) & Chr(10)
    
    Next
        
    'verifica se alguma linha do grid de pecas tem a quantidade menor do que a apurada pela tela ProdSolicSRV
    'se for menor, coloca o quantidade maior e atualiza os valores
    For iIndice = 1 To objGridItens.iLinhasExistentes

        For iIndice1 = colItemOVPeca.Count To 1 Step -1
        
            Set objItemOVPeca = colItemOVPeca.Item(iIndice1)
        
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sPeca, iPreenchido)
            If lErro <> SUCESSO Then gError 188232
        
            If objItemOVPeca.sProduto = sPeca Then
            
                dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
                
                If objItemOVPeca.dQuantidade <> dQuantidade Then
                    
                    sMsg = sMsg & "Peça: " & sPeca & " Quant.Grid Itens: " & CStr(dQuantidade) & " Quant. Detalhada: " & CStr(objItemOVPeca.dQuantidade) & Chr(10)
                                                
                End If
                
                colItemOVPeca.Remove (iIndice1)
                
                Exit For
                
            End If
            
        Next
        
    Next
    
    'para cada uma das pecas que nao foram encontrados no grid de pecas ==> adiciona-as ao grid
    For Each objItemOVPeca In colItemOVPeca
    
        sMsg = sMsg & "Peça: " & objItemOVPeca.sProduto & " Não está no grid de itens e Quant. Detalhada: " & CStr(objItemOVPeca.dQuantidade) & Chr(10)
    
    Next
    
    If Len(sMsg) > 0 Then
        sMsg = Chr(10) & sMsg
        gError 188235
    End If

    Valida_Totais_Servicos_Pecas = SUCESSO

    Exit Function

Erro_Valida_Totais_Servicos_Pecas:

    Valida_Totais_Servicos_Pecas = gErr

    Select Case gErr

        Case 188230 To 188232

        Case 188235
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTAIS_SERVICOS_PECAS", gErr, sMsg)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188233)

    End Select

    Exit Function

End Function

'#################################################
'Bloqueio

Public Sub BLGTipoBloqueio_Change()
    Call gobjTelaTrataBloqGen.TipoBloqueio_Change
End Sub

Public Sub BLGTipoBloqueio_GotFocus()
    Call gobjTelaTrataBloqGen.TipoBloqueio_GotFocus
End Sub

Public Sub BLGTipoBloqueio_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.TipoBloqueio_KeyPress(KeyAscii)
End Sub

Public Sub BLGTipoBloqueio_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.TipoBloqueio_Validate(Cancel)
End Sub

Public Sub BLGDataBloqueio_Change()
    Call gobjTelaTrataBloqGen.DataBloqueio_Change
End Sub

Public Sub BLGDataBloqueio_GotFocus()
    Call gobjTelaTrataBloqGen.DataBloqueio_GotFocus
End Sub

Public Sub BLGDataBloqueio_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.DataBloqueio_KeyPress(KeyAscii)
End Sub

Public Sub BLGDataBloqueio_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.DataBloqueio_Validate(Cancel)
End Sub

Public Sub BLGObservacao_Change()
    Call gobjTelaTrataBloqGen.Observacao_Change
End Sub

Public Sub BLGObservacao_GotFocus()
    Call gobjTelaTrataBloqGen.Observacao_GotFocus
End Sub

Public Sub BLGObservacao_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.Observacao_KeyPress(KeyAscii)
End Sub

Public Sub BLGObservacao_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.Observacao_Validate(Cancel)
End Sub

Public Sub BLGUsuario_Change()
    Call gobjTelaTrataBloqGen.Usuario_Change
End Sub

Public Sub BLGUsuario_GotFocus()
    Call gobjTelaTrataBloqGen.Usuario_GotFocus
End Sub

Public Sub BLGUsuario_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.Usuario_KeyPress(KeyAscii)
End Sub

Public Sub BLGUsuario_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.Usuario_Validate(Cancel)
End Sub

Public Sub BLGResponsavel_Change()
    Call gobjTelaTrataBloqGen.Responsavel_Change
End Sub

Public Sub BLGResponsavel_GotFocus()
    Call gobjTelaTrataBloqGen.Responsavel_GotFocus
End Sub

Public Sub BLGResponsavel_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.Responsavel_KeyPress(KeyAscii)
End Sub

Public Sub BLGResponsavel_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.Responsavel_Validate(Cancel)
End Sub

Public Sub BLGGridBloqueio_Click()
    Call gobjTelaTrataBloqGen.GridBloqueio_Click
End Sub

Public Sub BLGGridBloqueio_GotFocus()
    Call gobjTelaTrataBloqGen.GridBloqueio_GotFocus
End Sub

Public Sub BLGGridBloqueio_EnterCell()
    Call gobjTelaTrataBloqGen.GridBloqueio_EnterCell
End Sub

Public Sub BLGGridBloqueio_LeaveCell()
    Call gobjTelaTrataBloqGen.GridBloqueio_LeaveCell
End Sub

Public Sub BLGGridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
    Call gobjTelaTrataBloqGen.GridBloqueio_KeyDown(KeyCode, Shift)
End Sub

Public Sub BLGGridBloqueio_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.GridBloqueio_KeyPress(KeyAscii)
End Sub

Public Sub BLGGridBloqueio_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.GridBloqueio_Validate(Cancel)
End Sub

Public Sub BLGGridBloqueio_RowColChange()
    Call gobjTelaTrataBloqGen.GridBloqueio_RowColChange
End Sub

Public Sub BLGGridBloqueio_Scroll()
    Call gobjTelaTrataBloqGen.GridBloqueio_Scroll
End Sub

Public Sub BLGBotaoLibBloq_Click()
    Call gobjTelaTrataBloqGen.BotaoLibBloq_Click(Codigo)
End Sub

'Bloqueio
'#################################################

Private Sub SubTotal_Calcula()
'Faz a soma da Coluna passado no Grid passado e devolve em dValorTotal

Dim iIndice As Integer
Dim dValorTotal As Double, dValorTotalB As Double, dValorDesconto As Double

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
        End If
        dValorTotalB = dValorTotalB + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))
        dValorDesconto = dValorDesconto + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
    Next

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        End If
        dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
        dValorDesconto = dValorDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    Next

    'Coloca valor total dos produtos na tela
    ValorProdutos.Caption = Format(dValorTotal, "Standard")
    
    ValorProdutos2.Caption = Format(dValorTotalB, "Standard")
    
    ValorDescontoItens.Text = Format(dValorDesconto, "Standard")
    dValorDescontoItensAnt = dValorDesconto
    If dValorTotalB > 0 Then
        PercDescontoItens.Text = Format(dValorDesconto / dValorTotalB * 100, "FIXED")
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDesconto / dValorTotalB, 4)
    Else
        PercDescontoItens.Text = Format(0, "FIXED")
        dPercDescontoItensAnt = 0
    End If

    Exit Sub

End Sub


Public Sub ValorDescontoTit_Change()
    'Seta iComissoesAlterada
    'iComissoesAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorDescontoTit_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim dValorPagar As Double

On Error GoTo Erro_ValorDescontoTit_Validate

    dValorDesconto = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDescontoTit.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDescontoTit.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorDesconto = CDbl(ValorDescontoTit.Text)
        
        'Coloca o Valor formatado na tela
        ValorDescontoTit.Text = Format(dValorDesconto, "Standard")

    End If
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ValorTit.Caption = Format(dValorPagar, "Standard")
    ValorOriginalTit.Caption = Format(dValorPagar + dValorDesconto, "Standard")
    
    If Not gbCarregandoTela Then
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Exit Sub

Erro_ValorDescontoTit_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Public Sub BotaoInfoAdic_Click()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoInfoAdic_Click

    If gobjInfoAdic Is Nothing Then Set gobjInfoAdic = New ClassInfoAdic

    Call Chama_Tela_Modal("RetiradaEntrega", gobjInfoAdic, Me)

    Exit Sub

Erro_BotaoInfoAdic_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207733)

    End Select

    Exit Sub

End Sub

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

    If objGridItens.iLinhasExistentes > 0 Or objGridItensSRV.iLinhasExistentes > 0 Then
    
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
        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))) > 0 Then
                dValorItem = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))
                dDescontoItem = Arredonda_Moeda(dValorItem * dFator)
                dDescontoAplicado = dDescontoAplicado + dDescontoItem
                GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col) = Format(dDescontoItem, "STANDARD")
            End If
        Next
        If Abs(dDesconto - dDescontoAplicado) > DELTA_VALORMONETARIO Then
            If objGridItens.iLinhasExistentes > 0 Then
                GridItens.TextMatrix(1, iGrid_Desconto_Col) = Format(StrParaDbl(GridItens.TextMatrix(1, iGrid_Desconto_Col)) + dDescontoAplicado - dDesconto, "STANDARD")
            Else
                GridItensSRV.TextMatrix(1, iGrid_DescontoSRV_Col) = Format(StrParaDbl(GridItensSRV.TextMatrix(1, iGrid_DescontoSRV_Col)) + dDescontoAplicado - dDesconto, "STANDARD")
            End If
        End If
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
            dPercDesc = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
            Call PrecoTotal_Calcula(iIndice)
            
            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        Next
    
        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            dPercDesc = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col)) / StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))
            GridItensSRV.TextMatrix(iIndice, iGrid_PercDescSRV_Col) = Format(dPercDesc, "Percent")
            Call PrecoTotal_CalculaSRV(iIndice)
            
            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice, True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            dValorTotal = dValorTotal + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
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
        If objGridItens.iLinhasExistentes > 0 Or objGridItensSRV.iLinhasExistentes > 0 Then
            For iIndice = 1 To objGridItens.iLinhasExistentes
                dDesconto = dDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
            Next
            For iIndice = 1 To objGridItensSRV.iLinhasExistentes
                dDesconto = dDesconto + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
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
