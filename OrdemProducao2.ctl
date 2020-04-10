VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrdemProducao 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "'"
      Height          =   8415
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   585
      Width           =   16725
      Begin VB.Frame FrameClienteFornec 
         Caption         =   "Dono do Material (Material de Terceiros)"
         Height          =   825
         Left            =   60
         TabIndex        =   186
         Top             =   4350
         Width           =   10650
         Begin VB.ComboBox FilialTerc 
            Height          =   315
            Left            =   6675
            TabIndex        =   18
            Top             =   375
            Width           =   3165
         End
         Begin VB.OptionButton OptionTipoTerc 
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
            Height          =   210
            Index           =   2
            Left            =   4860
            TabIndex        =   188
            Top             =   150
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.OptionButton OptionTipoTerc 
            Caption         =   "Cliente"
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
            Left            =   3315
            TabIndex        =   187
            Top             =   150
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1020
         End
         Begin MSMask.MaskEdBox FornecedorTerc 
            Height          =   300
            Left            =   810
            TabIndex        =   189
            Top             =   375
            Visible         =   0   'False
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteTerc 
            Height          =   300
            Left            =   810
            TabIndex        =   17
            Top             =   375
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label ClienteLabel 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   135
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   185
            Top             =   435
            Width           =   660
         End
         Begin VB.Label LabelFilial 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6180
            TabIndex        =   191
            Top             =   435
            Width           =   465
         End
         Begin VB.Label FornecedorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fornec.:"
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
            Left            =   90
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   190
            Top             =   435
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Projeto"
         Height          =   855
         Left            =   60
         TabIndex        =   178
         Top             =   5520
         Width           =   10620
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
            Left            =   4320
            TabIndex        =   20
            Top             =   330
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   6645
            TabIndex        =   21
            Top             =   330
            Width           =   3195
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   795
            TabIndex        =   19
            Top             =   345
            Width           =   3540
            _ExtentX        =   6244
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
            Index           =   62
            Left            =   6000
            TabIndex        =   180
            Top             =   390
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
            Left            =   105
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   179
            Top             =   390
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Padrões "
         Height          =   1590
         Index           =   0
         Left            =   60
         TabIndex        =   67
         Top             =   2550
         Width           =   10650
         Begin VB.ComboBox EscaninhoTerc 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrdemProducao2.ctx":0000
            Left            =   6675
            List            =   "OrdemProducao2.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   990
            Width           =   3135
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data de Previsão"
            Height          =   690
            Left            =   75
            TabIndex        =   181
            Top             =   780
            Width           =   4035
            Begin MSComCtl2.UpDown UpDownInicio 
               Height          =   300
               Left            =   1845
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataInicioPadrao 
               Height          =   300
               Left            =   750
               TabIndex        =   11
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
            Begin MSComCtl2.UpDown UpDownFim 
               Height          =   300
               Left            =   3705
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataFimPadrao 
               Height          =   300
               Left            =   2610
               TabIndex        =   13
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
            Begin VB.Label DataPrevIniLbl 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   135
               TabIndex        =   183
               Top             =   300
               Width           =   555
            End
            Begin VB.Label DataPrevFimLbl 
               AutoSize        =   -1  'True
               Caption         =   "Fim:"
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
               Left            =   2250
               TabIndex        =   182
               Top             =   300
               Width           =   360
            End
         End
         Begin VB.ComboBox DestinacaoPadrao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrdemProducao2.ctx":0032
            Left            =   7470
            List            =   "OrdemProducao2.ctx":0034
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   315
            Width           =   2340
         End
         Begin MSMask.MaskEdBox CclPadrao 
            Height          =   315
            Left            =   825
            TabIndex        =   8
            Top             =   330
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxPadrao 
            Height          =   315
            Left            =   4290
            TabIndex        =   9
            Top             =   330
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrioridadePadrao 
            Height          =   315
            Left            =   5100
            TabIndex        =   15
            Top             =   990
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Escaninho:"
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
            Left            =   5715
            TabIndex        =   184
            Top             =   1050
            Width           =   960
         End
         Begin VB.Label PrioridadePadraoLbl 
            AutoSize        =   -1  'True
            Caption         =   " Prioridade:"
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
            Left            =   4095
            TabIndex        =   71
            Top             =   1050
            Width           =   990
         End
         Begin VB.Label DestPadraoLbl 
            AutoSize        =   -1  'True
            Caption         =   "Destinação:"
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
            Left            =   6405
            TabIndex        =   70
            Top             =   375
            Width           =   1020
         End
         Begin VB.Label AlmoxPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado:"
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
            Left            =   3135
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   69
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label CclPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Ccl:"
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
            TabIndex        =   68
            Top             =   375
            Width           =   345
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem de Produção"
         Height          =   2250
         Left            =   60
         TabIndex        =   61
         Top             =   90
         Width           =   10650
         Begin VB.CheckBox IgnoraEst 
            Caption         =   "Ignora o estoque no cálculo das necessidades"
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
            Left            =   5160
            TabIndex        =   192
            Top             =   1245
            Width           =   4410
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
            Left            =   9255
            TabIndex        =   175
            Top             =   855
            Width           =   495
         End
         Begin VB.CommandButton botaoProxNum2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4125
            Picture         =   "OrdemProducao2.ctx":0036
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Numeração Automática"
            Top             =   1230
            Width           =   300
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2715
            Picture         =   "OrdemProducao2.ctx":0120
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Numeração Automática"
            Top             =   375
            Width           =   300
         End
         Begin VB.TextBox PrestadorServico 
            Height          =   285
            Left            =   7635
            MaxLength       =   6
            TabIndex        =   7
            Top             =   1560
            Width           =   2130
         End
         Begin VB.TextBox Codigo 
            Height          =   285
            Left            =   1365
            MaxLength       =   9
            TabIndex        =   0
            Top             =   375
            Width           =   1350
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
            TabIndex        =   3
            Top             =   915
            Width           =   3270
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
            Height          =   270
            Left            =   5160
            TabIndex        =   4
            Top             =   915
            Width           =   4215
         End
         Begin VB.TextBox CodOPGerada 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2820
            MaxLength       =   9
            TabIndex        =   5
            Top             =   1230
            Width           =   1305
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   5370
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   360
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   4305
            TabIndex        =   1
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
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
            TabIndex        =   6
            Top             =   1575
            Width           =   3720
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
            Left            =   6300
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   94
            Top             =   1605
            Width           =   1290
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
            TabIndex        =   66
            Top             =   420
            Width           =   480
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
            TabIndex        =   65
            Top             =   420
            Width           =   1095
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
            TabIndex        =   64
            Top             =   1275
            Width           =   2220
         End
         Begin VB.Label StatusOP 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7620
            TabIndex        =   63
            Top             =   345
            Width           =   2175
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
            Left            =   6960
            TabIndex        =   62
            Top             =   405
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relatórios"
         Height          =   1170
         Left            =   60
         TabIndex        =   60
         Top             =   6660
         Width           =   10650
         Begin VB.Frame Frame6 
            Caption         =   "Imprimir"
            Height          =   615
            Left            =   225
            TabIndex        =   177
            Top             =   240
            Width           =   6345
            Begin VB.OptionButton OpcaoRelatorio 
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
               Index           =   2
               Left            =   5025
               TabIndex        =   24
               Top             =   270
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.OptionButton OpcaoRelatorio 
               Caption         =   " Rótulos para Produção"
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
               Left            =   2370
               TabIndex        =   23
               Top             =   255
               Width           =   2355
            End
            Begin VB.OptionButton OpcaoRelatorio 
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
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   22
               Top             =   255
               Width           =   2775
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Layout OP"
            Height          =   615
            Left            =   7305
            TabIndex        =   95
            Top             =   240
            Width           =   2580
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
               Left            =   165
               TabIndex        =   25
               Top             =   300
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.OptionButton OpcaoDetalhado 
               Caption         =   "Detalhado"
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
               Left            =   1275
               TabIndex        =   26
               Top             =   315
               Width           =   1215
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8325
      Index           =   3
      Left            =   60
      TabIndex        =   96
      Top             =   690
      Visible         =   0   'False
      Width           =   16740
      Begin VB.Frame Frame2 
         Caption         =   "Mão de Obra"
         Height          =   3675
         Index           =   5
         Left            =   75
         TabIndex        =   160
         Top             =   4500
         Visible         =   0   'False
         Width           =   16350
         Begin VB.ComboBox NomeMaquinaRealMO 
            Height          =   315
            Left            =   6435
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   2100
            Width           =   3825
         End
         Begin VB.TextBox DescricaoTipoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1320
            TabIndex        =   167
            Top             =   2025
            Width           =   4800
         End
         Begin VB.TextBox TipoMaoDeObraReal 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   195
            TabIndex        =   165
            Top             =   1410
            Width           =   990
         End
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
            TabIndex        =   161
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   3270
            Width           =   1305
         End
         Begin MSMask.MaskEdBox HorasRealMO 
            Height          =   315
            Left            =   6000
            TabIndex        =   162
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
            TabIndex        =   163
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
            TabIndex        =   164
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
            Height          =   2910
            Left            =   75
            TabIndex        =   166
            Top             =   195
            Width           =   16155
            _ExtentX        =   28496
            _ExtentY        =   5133
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Máquinas"
         Height          =   3660
         Index           =   4
         Left            =   75
         TabIndex        =   152
         Top             =   4500
         Visible         =   0   'False
         Width           =   16350
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
            Left            =   8940
            TabIndex        =   158
            ToolTipText     =   "Traz as máquinas que tiveram o uso previsto pelo MRP."
            Top             =   3270
            Width           =   2220
         End
         Begin VB.TextBox NomeMaquinaReal 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   180
            TabIndex        =   156
            Top             =   1425
            Width           =   4035
         End
         Begin VB.CommandButton BotaoMaquinasReal 
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
            Left            =   60
            TabIndex        =   153
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   3270
            Width           =   1305
         End
         Begin MSMask.MaskEdBox HorasReal 
            Height          =   315
            Left            =   4905
            TabIndex        =   154
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
            TabIndex        =   155
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
            TabIndex        =   157
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
            Height          =   2910
            Left            =   90
            TabIndex        =   159
            Top             =   210
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   5133
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3675
         Index           =   3
         Left            =   75
         TabIndex        =   97
         Top             =   4500
         Visible         =   0   'False
         Width           =   16350
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
            Left            =   9495
            TabIndex        =   151
            ToolTipText     =   "Abre a tela do Cronograma Gráfico das Etapas de Produção do Item da OP"
            Top             =   3075
            Width           =   1245
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
            Left            =   4680
            TabIndex        =   150
            ToolTipText     =   "Abre a tela de Centro de Trabalho"
            Top             =   3075
            Width           =   1950
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
            Left            =   7440
            TabIndex        =   149
            ToolTipText     =   "Abre a Tela de Competências"
            Top             =   3075
            Width           =   1500
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
            Left            =   2520
            TabIndex        =   148
            ToolTipText     =   "Abre a tela de Máquinas, Habilidades e Processos"
            Top             =   3075
            Width           =   1305
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
            TabIndex        =   147
            ToolTipText     =   "Abre a tela de Ordem de Produção para a Sub O.P."
            Top             =   3075
            Width           =   1590
         End
         Begin MSMask.MaskEdBox HorasPO 
            Height          =   315
            Left            =   5475
            TabIndex        =   128
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
            TabIndex        =   129
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
         Begin VB.Frame FrameDatas 
            Caption         =   "Datas"
            Height          =   1155
            Left            =   375
            TabIndex        =   100
            Top             =   1260
            Width           =   1995
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
               TabIndex        =   135
               Top             =   270
               Width           =   525
            End
            Begin VB.Label DataInicio 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   675
               TabIndex        =   134
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label DataFinal 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   675
               TabIndex        =   125
               Top             =   690
               Width           =   1155
            End
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
               TabIndex        =   101
               Top             =   720
               Width           =   525
            End
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
            Left            =   405
            TabIndex        =   54
            Top             =   405
            Width           =   225
         End
         Begin VB.TextBox OPCodigoMRP 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1050
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   55
            Top             =   810
            Width           =   1170
         End
         Begin VB.TextBox NomeMaquina 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2565
            TabIndex        =   99
            Top             =   1425
            Width           =   1950
         End
         Begin VB.TextBox TaxaProducao 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   7530
            MaxLength       =   15
            TabIndex        =   98
            Top             =   1425
            Width           =   1185
         End
         Begin MSMask.MaskEdBox QuantidadeMaquina 
            Height          =   315
            Left            =   4515
            TabIndex        =   102
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
            Height          =   2415
            Left            =   3855
            TabIndex        =   56
            Top             =   435
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   4260
            _Version        =   393216
         End
         Begin VB.Label LabelMRP 
            Caption         =   "Plano Mestre de Produção Gerado"
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
            Left            =   660
            TabIndex        =   139
            Top             =   420
            Width           =   2910
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
            Left            =   555
            TabIndex        =   104
            Top             =   840
            Width           =   420
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
            Left            =   3840
            TabIndex        =   103
            Top             =   210
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Insumos por Operações"
         Height          =   3660
         Index           =   2
         Left            =   75
         TabIndex        =   105
         Top             =   4500
         Visible         =   0   'False
         Width           =   16350
         Begin VB.TextBox OrigemProduto 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3195
            TabIndex        =   146
            Top             =   690
            Width           =   1020
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2025
            TabIndex        =   110
            Top             =   1995
            Width           =   4000
         End
         Begin VB.ComboBox UMProduto 
            Height          =   315
            Left            =   4830
            TabIndex        =   109
            Top             =   1035
            Width           =   1305
         End
         Begin VB.ComboBox ComposicaoInsumos 
            Height          =   315
            ItemData        =   "OrdemProducao2.ctx":020A
            Left            =   6975
            List            =   "OrdemProducao2.ctx":0214
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   1020
            Width           =   1635
         End
         Begin VB.ComboBox VersaoKitComp 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OrdemProducao2.ctx":0228
            Left            =   5475
            List            =   "OrdemProducao2.ctx":022A
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   1035
            Width           =   1440
         End
         Begin MSMask.MaskEdBox ProdutoInsumos 
            Height          =   315
            Left            =   435
            TabIndex        =   106
            Top             =   1035
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PerdaInsumos 
            Height          =   315
            Left            =   6285
            TabIndex        =   111
            Top             =   1035
            Width           =   1335
            _ExtentX        =   2355
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
            TabIndex        =   112
            Top             =   1020
            Width           =   1410
            _ExtentX        =   2487
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
            TabIndex        =   113
            Top             =   1035
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridOperacaoInsumos 
            Height          =   3240
            Left            =   90
            TabIndex        =   53
            Top             =   255
            Width           =   16155
            _ExtentX        =   28496
            _ExtentY        =   5715
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operação"
         Height          =   3675
         Index           =   1
         Left            =   75
         TabIndex        =   114
         Top             =   4500
         Width           =   16350
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
            TabIndex        =   172
            ToolTipText     =   "Abre a Tela de Kit"
            Top             =   2640
            Width           =   1305
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
            TabIndex        =   171
            ToolTipText     =   "Abre a Tela de Roteiro de Fabricação"
            Top             =   1980
            Width           =   1305
         End
         Begin VB.Frame FrameConsideraAlgoritmo 
            Caption         =   "Algorítmo"
            Height          =   480
            Left            =   5175
            TabIndex        =   141
            Top             =   3060
            Width           =   2475
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
               TabIndex        =   50
               Tag             =   "Controla a alocação de máquinas."
               Top             =   210
               Width           =   870
            End
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
               TabIndex        =   49
               Tag             =   "Não controla a alocação de máquinas."
               Top             =   210
               Width           =   870
            End
         End
         Begin VB.TextBox Observacao 
            Height          =   735
            Left            =   1365
            MaxLength       =   255
            TabIndex        =   46
            Top             =   1950
            Width           =   6285
         End
         Begin MSMask.MaskEdBox CodigoCTPadrao 
            Height          =   315
            Left            =   1380
            TabIndex        =   45
            Top             =   1545
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
            TabIndex        =   44
            Top             =   1125
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
            TabIndex        =   47
            Top             =   2760
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
            TabIndex        =   48
            Top             =   2760
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
            Left            =   5175
            TabIndex        =   174
            Top             =   2790
            Width           =   2010
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
            TabIndex        =   170
            Top             =   2790
            Width           =   2550
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
            TabIndex        =   140
            Top             =   765
            Width           =   480
         End
         Begin VB.Label UMLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8235
            TabIndex        =   138
            Top             =   705
            Width           =   780
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
            TabIndex        =   137
            Top             =   765
            Width           =   1050
         End
         Begin VB.Label QtdeLabel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6015
            TabIndex        =   136
            Top             =   705
            Width           =   1470
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
            TabIndex        =   133
            Top             =   315
            Width           =   810
         End
         Begin VB.Label ProdutoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   132
            Top             =   285
            Width           =   7635
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
            TabIndex        =   131
            Top             =   765
            Width           =   690
         End
         Begin VB.Label VersaoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   130
            Top             =   705
            Width           =   1665
         End
         Begin VB.Label LabelCodigoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3930
            TabIndex        =   127
            Top             =   3225
            Width           =   420
         End
         Begin VB.Label LabelCodigoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3480
            TabIndex        =   126
            Top             =   3225
            Width           =   420
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
            TabIndex        =   123
            Top             =   1155
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
            TabIndex        =   122
            Top             =   1575
            Width           =   1110
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
            TabIndex        =   121
            Top             =   2010
            Width           =   1155
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
            TabIndex        =   120
            Top             =   3270
            Width           =   1020
         End
         Begin VB.Label Nivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   119
            Top             =   3225
            Width           =   420
         End
         Begin VB.Label Sequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3030
            TabIndex        =   118
            Top             =   3225
            Width           =   420
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
            TabIndex        =   117
            Top             =   3255
            Width           =   540
         End
         Begin VB.Label DescricaoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   116
            Top             =   1125
            Width           =   5145
         End
         Begin VB.Label DescricaoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   115
            Top             =   1545
            Width           =   5145
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
         Left            =   15315
         TabIndex        =   145
         ToolTipText     =   "Abre o Relatório de Operações"
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Frame FrameItemOP 
         Caption         =   "Item da OP"
         Height          =   660
         Left            =   15300
         TabIndex        =   142
         Top             =   -15
         Width           =   1275
         Begin MSComCtl2.UpDown UpDownItemOP 
            Height          =   315
            Left            =   780
            TabIndex        =   144
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
            TabIndex        =   143
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
      Begin VB.Frame FrameRoteiro 
         Caption         =   "Roteiro de Fabricação:"
         Height          =   4080
         Left            =   30
         TabIndex        =   124
         Top             =   -15
         Width           =   14895
         Begin MSComctlLib.TreeView Roteiro 
            Height          =   3660
            Left            =   120
            TabIndex        =   43
            Top             =   270
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   6456
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
         Left            =   15300
         Picture         =   "OrdemProducao2.ctx":022C
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Altera a Operação da Árvore do Roteiro"
         Top             =   750
         Width           =   1275
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   4185
         Left            =   30
         TabIndex        =   51
         Top             =   4110
         Width           =   16530
         _ExtentX        =   29157
         _ExtentY        =   7382
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
      Height          =   8400
      Index           =   2
      Left            =   60
      TabIndex        =   72
      Top             =   600
      Visible         =   0   'False
      Width           =   16695
      Begin VB.CommandButton BotaoConsultaPV 
         Caption         =   "Consulta PV"
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
         Left            =   7695
         TabIndex        =   42
         Top             =   7800
         Width           =   1215
      End
      Begin VB.ComboBox Benef 
         Height          =   315
         ItemData        =   "OrdemProducao2.ctx":1B52
         Left            =   600
         List            =   "OrdemProducao2.ctx":1B5F
         Style           =   2  'Dropdown List
         TabIndex        =   176
         Top             =   750
         Width           =   1260
      End
      Begin VB.CheckBox ProduzLogo 
         Height          =   210
         Left            =   4755
         TabIndex        =   173
         Top             =   3405
         Width           =   1125
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
         Height          =   375
         Left            =   11280
         TabIndex        =   169
         Top             =   7785
         Width           =   2670
      End
      Begin VB.CommandButton BotaoImprimirPrevia 
         Caption         =   "Prévia de Necessidades"
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
         Left            =   8985
         TabIndex        =   36
         Top             =   7800
         Width           =   2235
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
         Left            =   2205
         TabIndex        =   35
         Top             =   7815
         Width           =   975
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
         Left            =   1035
         TabIndex        =   38
         Top             =   7815
         Width           =   1125
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
         Left            =   -15
         TabIndex        =   37
         Top             =   7815
         Width           =   1005
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
         Left            =   3225
         TabIndex        =   39
         Top             =   7815
         Width           =   1680
      End
      Begin VB.ComboBox ComboFilialPedido 
         Height          =   315
         ItemData        =   "OrdemProducao2.ctx":1B84
         Left            =   2175
         List            =   "OrdemProducao2.ctx":1B86
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   1590
         Width           =   1875
      End
      Begin VB.ComboBox Situacao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OrdemProducao2.ctx":1B88
         Left            =   4920
         List            =   "OrdemProducao2.ctx":1B8A
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   1140
         Width           =   1830
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   5850
         MaxLength       =   50
         TabIndex        =   79
         Top             =   1950
         Width           =   4000
      End
      Begin VB.TextBox UnidadeMed 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   78
         Top             =   435
         Width           =   600
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
         Left            =   4935
         TabIndex        =   40
         Top             =   7815
         Width           =   1680
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3225
         TabIndex        =   77
         Text            =   "Cliente"
         Top             =   2100
         Width           =   1260
      End
      Begin VB.TextBox FilialCliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   240
         Left            =   4170
         TabIndex        =   76
         Text            =   "Filial do Cliente"
         Top             =   2400
         Width           =   1110
      End
      Begin VB.ComboBox Destinacao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OrdemProducao2.ctx":1B8C
         Left            =   6825
         List            =   "OrdemProducao2.ctx":1B8E
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1125
         Width           =   1830
      End
      Begin VB.ComboBox Versao 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "OrdemProducao2.ctx":1B90
         Left            =   6600
         List            =   "OrdemProducao2.ctx":1B92
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   2340
         Width           =   1095
      End
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
         Height          =   360
         Left            =   6645
         TabIndex        =   41
         Top             =   7815
         Width           =   1020
      End
      Begin MSMask.MaskEdBox Maquina 
         Height          =   255
         Left            =   6870
         TabIndex        =   73
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
         TabIndex        =   82
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
         TabIndex        =   83
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
         TabIndex        =   84
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
         TabIndex        =   85
         Top             =   1650
         Width           =   1600
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
         TabIndex        =   86
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
         TabIndex        =   87
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
         TabIndex        =   88
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
         TabIndex        =   89
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
         Height          =   7230
         Left            =   30
         TabIndex        =   34
         Top             =   420
         Width           =   16620
         _ExtentX        =   29316
         _ExtentY        =   12753
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
         TabIndex        =   92
         Top             =   120
         Width           =   2100
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
         Left            =   14190
         TabIndex        =   91
         Top             =   7905
         Width           =   1455
      End
      Begin VB.Label QuantDisponivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   15735
         TabIndex        =   90
         Top             =   7845
         Width           =   885
      End
   End
   Begin VB.CheckBox ImprimeAoGravar 
      Caption         =   "Imprimir ao Gravar"
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
      Left            =   12120
      TabIndex        =   27
      Top             =   45
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   14220
      ScaleHeight     =   435
      ScaleWidth      =   2610
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   0
      Width           =   2670
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2100
         Picture         =   "OrdemProducao2.ctx":1B94
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Fechar"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "OrdemProducao2.ctx":1D12
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Limpar"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1125
         Picture         =   "OrdemProducao2.ctx":2244
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Excluir"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   615
         Picture         =   "OrdemProducao2.ctx":23CE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Gravar"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   120
         Picture         =   "OrdemProducao2.ctx":2528
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprimir"
         Top             =   30
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8865
      Left            =   15
      TabIndex        =   28
      Top             =   240
      Width           =   16875
      _ExtentX        =   29766
      _ExtentY        =   15637
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
Attribute VB_Name = "OrdemProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTOrdemProducao
Attribute objCT.VB_VarHelpID = -1

Private Sub botaoProxNum2_Click()
    Call objCT.botaoProxNum2_Click
End Sub


Private Sub UserControl_Initialize()
    Set objCT = New CTOrdemProducao
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
     Call objCT.BotaoMaquinas_Click
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

Private Sub BotaoConsultaPV_Click()
     Call objCT.BotaoConsultaPV_Click
End Sub

Private Sub ClienteLabel_Click()
    Call objCT.ClienteLabel_Click
End Sub

Private Sub FornecedorLabel_Click()
    Call objCT.FornecedorLabel_Click
End Sub

Private Sub FilialTerc_Change()
    Call objCT.FilialTerc_Change
End Sub

Private Sub FornecedorTerc_Change()
    Call objCT.FornecedorTerc_Change
End Sub

Private Sub FornecedorTerc_Validate(Cancel As Boolean)
    Call objCT.FornecedorTerc_Validate(Cancel)
End Sub

Private Sub ClienteTerc_Change()
    Call objCT.ClienteTerc_Change
End Sub

Private Sub ClienteTerc_Validate(Cancel As Boolean)
    Call objCT.ClienteTerc_Validate(Cancel)
End Sub

Private Sub FilialTerc_Validate(Cancel As Boolean)
    Call objCT.FilialTerc_Validate(Cancel)
End Sub

Private Sub OptionTipoTerc_Click(Index As Integer)
    Call objCT.OptionTipoTerc_Click(Index)
End Sub

Private Sub EscaninhoTerc_Click()
    Call objCT.EscaninhoTerc_Click
End Sub

Private Sub IgnoraEst_Click()
     Call objCT.IgnoraEst_Click
End Sub
