VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DevolucaoChequeOcx 
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   9495
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4740
      Index           =   1
      Left            =   255
      TabIndex        =   11
      Top             =   855
      Width           =   8985
      Begin VB.Frame Frame5 
         Caption         =   "Titulo a Pagar (só para cheque vinculado a borderô de desconto)"
         Height          =   1110
         Left            =   90
         TabIndex        =   77
         Top             =   3510
         Width           =   8835
         Begin MSMask.MaskEdBox ValorCreditado 
            Height          =   300
            Left            =   3015
            TabIndex        =   6
            Top             =   675
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   6540
            TabIndex        =   7
            Top             =   255
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   3015
            TabIndex        =   5
            Top             =   255
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSComCtl2.UpDown UpDownVencimento 
            Height          =   300
            Left            =   7620
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   675
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   300
            Left            =   6540
            TabIndex        =   8
            Top             =   675
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelValorCreditado 
            AutoSize        =   -1  'True
            Caption         =   "Valor Creditado:"
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
            Index           =   10
            Left            =   1575
            TabIndex        =   80
            Top             =   735
            Width           =   1380
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   " Filial:"
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
            Left            =   5940
            TabIndex        =   82
            Top             =   300
            Width           =   525
         End
         Begin VB.Label LabelDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
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
            Left            =   5400
            TabIndex        =   79
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label LabelFornecedor 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1920
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   78
            Top             =   300
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados do Cheque"
         Height          =   2115
         Left            =   90
         TabIndex        =   58
         Top             =   1245
         Width           =   8835
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
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
            Left            =   5685
            TabIndex        =   76
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   2370
            TabIndex        =   75
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Banco 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3060
            TabIndex        =   74
            Top             =   195
            Width           =   870
         End
         Begin VB.Label Agencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6525
            TabIndex        =   73
            Top             =   195
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bom Para:"
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
            Left            =   2100
            TabIndex        =   72
            Top             =   990
            Width           =   885
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   5730
            TabIndex        =   71
            Top             =   630
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Index           =   5
            Left            =   2415
            TabIndex        =   70
            Top             =   615
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Index           =   7
            Left            =   5940
            TabIndex        =   69
            Top             =   990
            Width           =   510
         End
         Begin VB.Label ContaCorrente 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3060
            TabIndex        =   68
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label Numero 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6525
            TabIndex        =   67
            Top             =   570
            Width           =   1395
         End
         Begin VB.Label DataBomPara 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3060
            TabIndex        =   66
            Top             =   945
            Width           =   1170
         End
         Begin VB.Label ValorChq 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6540
            TabIndex        =   65
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Borderô:"
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
            Index           =   8
            Left            =   1545
            TabIndex        =   64
            Top             =   1365
            Width           =   1440
         End
         Begin VB.Label TipoBordero 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3060
            TabIndex        =   63
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Depósito:"
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
            Index           =   6
            Left            =   1725
            TabIndex        =   62
            Top             =   1755
            Width           =   1290
         End
         Begin VB.Label DataDeposito 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3060
            TabIndex        =   61
            Top             =   1695
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Borderô:"
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
            Index           =   4
            Left            =   5715
            TabIndex        =   60
            Top             =   1365
            Width           =   735
         End
         Begin VB.Label NumBordero 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6525
            TabIndex        =   59
            Top             =   1320
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificador da Devolução"
         Height          =   1065
         Left            =   60
         TabIndex        =   55
         Top             =   45
         Width           =   8835
         Begin VB.CommandButton BotaoCheque 
            Caption         =   "Selecione o cheque"
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
            Left            =   3120
            TabIndex        =   2
            Top             =   645
            Width           =   2340
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   4230
            Picture         =   "DevolucaoCheque.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "Numeração Automática"
            Top             =   255
            Width           =   300
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   3135
            TabIndex        =   0
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   6465
            TabIndex        =   1
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   7605
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox SeqChq 
            Height          =   300
            Left            =   3135
            TabIndex        =   4
            Top             =   660
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelSeqChq 
            AutoSize        =   -1  'True
            Caption         =   "Cheque:"
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
            Left            =   2325
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   81
            Top             =   705
            Visible         =   0   'False
            Width           =   720
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
            Index           =   9
            Left            =   5910
            TabIndex        =   57
            Top             =   300
            Width           =   480
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
            Left            =   2385
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   56
            Top             =   300
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4530
      Index           =   2
      Left            =   255
      TabIndex        =   12
      Top             =   915
      Visible         =   0   'False
      Width           =   8985
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4800
         TabIndex        =   89
         Tag             =   "1"
         Top             =   1560
         Width           =   870
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
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
         Left            =   3480
         TabIndex        =   25
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   20
         Top             =   3465
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1875
            TabIndex        =   24
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1875
            TabIndex        =   23
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Left            =   1215
            TabIndex        =   22
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            Left            =   330
            TabIndex        =   21
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6345
         TabIndex        =   19
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   18
         Top             =   2175
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   17
         Top             =   2565
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
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
         Left            =   6270
         TabIndex        =   16
         Top             =   390
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6270
         TabIndex        =   15
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   870
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7710
         TabIndex        =   13
         Top             =   30
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4800
         TabIndex        =   26
         Top             =   1920
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   27
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   28
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   29
         Top             =   1830
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   30
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   32
         Top             =   555
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5565
         TabIndex        =   33
         Top             =   150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3780
         TabIndex        =   34
         Top             =   165
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   15
         TabIndex        =   35
         Top             =   1215
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6345
         TabIndex        =   36
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6345
         TabIndex        =   37
         Top             =   1530
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   54
         Top             =   210
         Width           =   450
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   53
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label CTBLabel8 
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
         Height          =   195
         Left            =   45
         TabIndex        =   52
         Top             =   600
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   51
         Top             =   3090
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   50
         Top             =   3075
         Width           =   1155
      End
      Begin VB.Label CTBLabelTotais 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Totais:"
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
         Height          =   225
         Left            =   1800
         TabIndex        =   49
         Top             =   3090
         Width           =   615
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   48
         Top             =   1290
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   6345
         TabIndex        =   47
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   46
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         TabIndex        =   45
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   1995
         TabIndex        =   44
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   43
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5025
         TabIndex        =   42
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4245
         TabIndex        =   41
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   765
         TabIndex        =   40
         Top             =   195
         Width           =   1530
      End
      Begin VB.Label CTBLabel21 
         Caption         =   "Origem:"
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
         Height          =   255
         Left            =   45
         TabIndex        =   39
         Top             =   210
         Width           =   720
      End
      Begin VB.Label CTBLabel1 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6300
         TabIndex        =   38
         Top             =   660
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7275
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "DevolucaoCheque.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DevolucaoCheque.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "DevolucaoCheque.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DevolucaoCheque.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5145
      Left            =   75
      TabIndex        =   10
      Top             =   495
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9075
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
Attribute VB_Name = "DevolucaoChequeOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

'Variáveis Globáis
Dim iFornecedorAlterado  As Integer
Public objContabil As New ClassContabil
Public iAlterado As Integer
Public objGrid As AdmGrid

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoDevCheque As AdmEvento
Attribute objEventoDevCheque.VB_VarHelpID = -1
Private WithEvents objEventoCheque As AdmEvento
Attribute objEventoCheque.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Public iFrameAtual As Integer

Private Const VALOR_CHQ = "Valor_Cheque"

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()
'Inicialização da Tela de Devolução de Cheque

Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Visible = False
    
    End If
    
    Set objEventoLote = New AdmEvento
    Set objEventoDoc = New AdmEvento
    Set objEventoDevCheque = New AdmEvento
    Set objEventoCheque = New AdmEvento
    Set objEventoFornecedor = New AdmEvento

    iFrameAtual = 1

    Data.PromptInclude = False

    'Move a Data atual para o campo de data na tela
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")

    Data.PromptInclude = True

    'Chama o Metodo da Classe Contabil
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid, objEventoLote, objEventoDoc, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then gError 111330

    'Define que não Houve Alteração
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 111330

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158954)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objDevCheque As ClassDevCheque) As Long
'Trata os parametros

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Trata_Parametros

    'Verifica se o Obj esta Instanciado
    If Not (objDevCheque Is Nothing) Then

        'Função que lê os Cheques devolvidos no Banco de Dados e Preenche a Tela
        lErro = Traz_DevCheque_Tela(objDevCheque)
        If lErro <> SUCESSO Then gError 111332

        If lErro <> SUCESSO Then

                'Limpa a Tela
                Call Limpa_Tela(Me)

                'Mantém o Código do Cheque na tela
                Codigo.Text = objDevCheque.lCodigo


                Data.PromptInclude = False

                'Move o a Data atual para o campo data da tela
                Data.Text = gdtDataAtual

                Data.PromptInclude = True

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 111332
            'Erros Tratados Dentro da Função Chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158955)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoCheque_Click()
    Call LabelSeqChq_Click
End Sub

'***************************************************************
'*
'* Browser 's Referentes a Tela de devolução de Cheques
'*
'***************************************************************

Private Sub LabelCodigo_Click()

Dim objDevCheque As New ClassDevCheque
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o código está preenchido
    If Len(Trim(Codigo.Text)) > 0 Then

        objDevCheque.lCodigo = StrParaLong(Codigo.Text)

    End If

    'Chama o Browser Devolucao de Cheque
    Call Chama_Tela("DevolucaoChequeLista", colSelecao, objDevCheque, objEventoDevCheque, "NumIntCheque IN (SELECT NumIntCheque FROM ChequePre)")

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158956)

    End Select

    Exit Sub

End Sub

Private Sub objEventoDevCheque_evSelecao(obj1 As Object)

Dim objDevCheque As ClassDevCheque
Dim lErro As Long
Dim lCodigoMsgErro As Long

On Error GoTo Erro_objEventoDevCheque_evSelecao

    Set objDevCheque = obj1

    'Função que lê os Cheques devolvidos no Banco de Dados e Preenche a Tela(Tabela DevolucaoCheque)
    lErro = Traz_DevCheque_Tela(objDevCheque)
    If lErro <> SUCESSO Then gError 111333

    'Cheque não Encontrado no Banco de Dados
    If lErro = 111331 Then gError 111334

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoDevCheque_evSelecao:

    Select Case gErr

        Case 111333

        Case 111334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_DEVOLDIDO_INEXISTENTE", gErr, objDevCheque.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158957)

    End Select

    Exit Sub

End Sub

Private Sub LabelSeqChq_Click()
'Browser que Traz os Dados do Cheque para a Tela( Cheque devolvido )
'somente de cheques relacionados ao bordero e não relacionados a devolução

Dim objCheque As New ClassChequePre
Dim sSelecao As String
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSeqChq_Click

    'Verifica se o código está preenchido
    If Len(Trim(LabelSeqChq.Caption)) > 0 Then

        objCheque.lSequencialBack = StrParaLong(Codigo.Text)

    End If

    'Chama o Browser Devolucao de Cheque
    Call Chama_Tela("ChequeBordDevLista", colSelecao, objCheque, objEventoCheque)

    Exit Sub

Erro_LabelSeqChq_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158958)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCheque_evSelecao(obj1 As Object)
'Função que Traz os Cheque para a Tela e dados do Bordêro de desconto se o Mesmo fizer parte

Dim objCheque As ClassChequePre
Dim lErro As Long

On Error GoTo Erro_objEventoCheque_evSelecao

    Set objCheque = obj1
    
    SeqChq.Text = objCheque.lSequencialBack
    
    lErro = Traz_Cheque_Tela(objCheque)
    If lErro <> SUCESSO Then gError 112043
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCheque_evSelecao:

    Select Case gErr
        
        Case 112043

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158959)

    End Select

    Exit Sub

End Sub

Function Traz_Cheque_Tela(objCheque As ClassChequePre) As Long

Dim lErro As Long
Dim objChqBord As New ClassChequeBordero

On Error GoTo Erro_Traz_Cheque_Tela

    'Função que lê os Cheques
    lErro = CF("ChequePre_Le2", objCheque)
    If lErro <> SUCESSO And lErro <> 109964 Then gError 111336

    'Cheque não Encontrado no Banco de Dados
    If lErro = 109964 Then gError 111337
        
    'Verififica se o Cheque não esta relaionado a um Bordrô( Não foi depositado ) e não pode ter sido devolvido( fazer parte da tabela de cheques devolvidos)
    lErro = CF("ChequeRel_BordDevolvido", objCheque)
    If lErro <> SUCESSO And lErro <> 109871 Then gError 111339

    'Se o cheque não estiver em terceiros
    If objCheque.iLocalizacao <> CHEQUEPRE_LOCALIZACAO_EM_TERCEIROS Then
        'se nao há correspondencia entre as tabelas da seleção acima
        If lErro = 109871 Then gError 111340
    End If
    
    Call Inicializa_Cheque
    
    'Função que Traz para a Tela os Dados Relacionados a Borderô de Desconto
    lErro = Traz_Dados_BorderoDesc_Tela(objChqBord, objCheque)
    If lErro <> SUCESSO Then gError 111341
    
    Call Preenche_Dados_Cheque(objCheque)
    
    Traz_Cheque_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Cheque_Tela:
    
    Traz_Cheque_Tela = gErr
    
    Select Case gErr

        Case 111337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_INEXISTENTE", gErr, objCheque.lSequencialBack)

        Case 111336, 111339

        Case 111340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_VINCULADO_BORDERO", gErr, objCheque.lNumIntCheque)

        Case 111341

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158960)

    End Select
    
End Function

Private Sub Inicializa_Cheque()
    
    'Limpa os Caption's Relacionados ao Cheque que será devolvido
    Banco.Caption = ""
    Agencia.Caption = ""
    ContaCorrente.Caption = ""
    Numero.Caption = ""
    DataBomPara.Caption = ""
    ValorChq.Caption = ""
    TipoBordero.Caption = ""
    NumBordero.Caption = ""
    DataDeposito.Caption = ""
    
    Fornecedor.Text = ""
    Filial.Text = ""
    ValorCreditado.Text = ""
    
    DataVencimento.PromptInclude = False
    DataVencimento.Text = ""
    DataVencimento.PromptInclude = True
    
End Sub

Private Function Preenche_Dados_Cheque(objCheque As ClassChequePre) As Long
    
Dim lErro As Long

On Error GoTo Erro_Preenche_Dados_Cheque

    Banco.Caption = objCheque.iBanco
    ContaCorrente.Caption = objCheque.sContaCorrente
    
    DataBomPara.Caption = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
    Agencia.Caption = objCheque.sAgencia
    Numero.Caption = objCheque.lNumero
    ValorChq.Caption = Format(objCheque.dValor, "standard")
    
    If objCheque.lNumBordero <> 0 Then
        NumBordero.Caption = objCheque.lNumBordero
    End If
    
    If objCheque.iTipoBordero = TIPO_BORDERO_CHEQUEPRE Then TipoBordero.Caption = TIPO_BORDERO_CHEQUEPRE_TEXTO
    If objCheque.iTipoBordero = TIPO_BORDERO_DESCONTO Then TipoBordero.Caption = TIPO_BORDERO_DESCONTO_TEXTO
    
    'leitura para datadeposito borderoschequespre
    lErro = CF("BorderosChequesPre_Le", objCheque)
    If lErro <> SUCESSO And lErro <> 109970 Then gError 109966
    
'    If lErro = 109970 Then gError 109971
      
    DataDeposito.Caption = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
  
    Preenche_Dados_Cheque = SUCESSO
    
    Exit Function

Erro_Preenche_Dados_Cheque:
    
    Preenche_Dados_Cheque = gErr
    
    Select Case gErr
        
        Case 109966
        
        Case 109971
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158961)

    End Select

    Exit Function

End Function

Private Sub LabelFornecedor_Click()
'Função que Traz o Browse de Fornecedor para a Tela

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

On Error GoTo Erro_LabelFornecedor_Click

    'Verifica se fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then
        objFornecedor = Fornecedor.Text
    End If
    
    'Chama o Browse de Fornecedor
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

Erro_LabelFornecedor_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158962)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)
'Traz os Fornecedores

Dim objFornecedor As ClassFornecedor
Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1

    'Move os Dados do fornecedor para a tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Chama a Função que valida o nome
    Call Fornecedor_Validate(bCancel)

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158963)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objDevCheque As New ClassDevCheque

On Error GoTo Erro_Tela_Extrai

    sTabela = "DevolucaoCheque"

    'Move os dados da tela de Cheque devolvidos para a memória
    lErro = Move_Tela_Memoria(objDevCheque)
    If lErro <> SUCESSO Then gError 111342

    'Seleção
    colCampoValor.Add "Codigo", objDevCheque.lCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", objDevCheque.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Data", objDevCheque.dtData, 0, "Data"
    colCampoValor.Add "Fornecedor", objDevCheque.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objDevCheque.iFilial, 0, "Filial" 'Filial do Fornecedor
    colCampoValor.Add "DataVencimento", objDevCheque.dtDataVencimento, 0, "DataVencimento"
    colCampoValor.Add "ValorCredito", objDevCheque.dValorCredito, 0, "ValorCredito"
    colCampoValor.Add "NumIntChqBord", objDevCheque.lNumIntChqBord, 0, "NumIntChqBord"
    colCampoValor.Add "NumIntBaixasParcRecCanc", objDevCheque.lNumIntBaixasParcRecCanc, 0, "NumIntBaixasParcRecCanc"

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 111342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158964)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objDevCheque As New ClassDevCheque
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da colecao de campos-valores para o objDevCheque
    objDevCheque.lCodigo = colCampoValor.Item("Codigo").vValor
    objDevCheque.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor

    If objDevCheque.lCodigo <> 0 Then

        'Se o Sequencial do Cheque nao for nulo Traz o Cheque para a tela
        lErro = Traz_DevCheque_Tela(objDevCheque)
        If lErro <> SUCESSO Then gError 111343

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 111343

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158965)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
'Botão que Gera o Codigo Automático

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_CODDEVCHEQUE", "DevolucaoCheque", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 111344
    
    'Joga o Codigo ma Tela
    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True
    
    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 111344

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158966)

    End Select

    Exit Sub

End Sub

'**************************************
'***** Tratamentos dos Controles ******
'**************************************

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
'Valida o Codigo

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se o Codigo esta Preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        'Função que Critica se o Valor Colocado como Codigo é Long
        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 111345

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 111345

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158967)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se a data esta Preenchida
    If Len(Trim(Data.Text)) = 0 Then

        'Função que critica se a data é válida
        lErro = Data_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 111346

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 111346

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158968)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()
'Função que serve para decrementar a Data

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 111347

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 111347
            'Erro Tratado Dentro da Função Chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158969)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()
'Função que serve para imcrementar a Data

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 111348

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 111348

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158970)

    End Select

    Exit Sub

End Sub

Private Sub SeqChq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SeqChq_GotFocus()

    Call MaskEdBox_TrataGotFocus(SeqChq)

End Sub

Private Sub SeqChq_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCheque As New ClassChequePre
Dim objDevCheque As New ClassDevCheque

On Error GoTo Erro_SeqChq_Validate

    'Verifica se o Seq do Cheque esta Preenchido
    If Len(Trim(SeqChq.Text)) <> 0 Then

        'Função que Critica se o Valor Colocado como Codigo é Long
        lErro = Long_Critica(SeqChq.Text)
        If lErro <> SUCESSO Then gError 111349
        
        objCheque.lSequencialBack = StrParaLong(SeqChq.Text)
        objCheque.iFilialEmpresaLoja = giFilialEmpresa
        objCheque.iFilialEmpresa = giFilialEmpresa
        
        objDevCheque.lCodigo = StrParaLong(Codigo.Text)
                
        lErro = Traz_Cheque_Tela(objCheque)
        If lErro <> SUCESSO Then
            objDevCheque.lCodigo = StrParaLong(Codigo.Text)
            lErro = CF("DevCheque_Le", objDevCheque)
            If lErro <> SUCESSO Then gError 112048
        End If
    End If

    Exit Sub

Erro_SeqChq_Validate:

    Cancel = True

    Select Case gErr

        Case 111349, 112048, 112050
        
        Case 112049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_INEXISTENTE", gErr, objCheque.lSequencialBack)

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158971)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
'Função que valida os dados do Fornecedor

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se o Fornecedor foi Alterado
    If iFornecedorAlterado = REGISTRO_ALTERADO Then

        'Verifica se o Fornecedor está preenchido
        If Len(Trim(Fornecedor.Text)) <> 0 Then
            'Função que Lê o Fornecedor Pelo Nome/Codigo/CPF/CGC , iCodFilial Traz o Codigo da Filial Matriz
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 111350
            'obs se não Encontrar o fornecedor pelo nome erro 6663
            'obs se não Encontrar o fornecedor pelo Codigo erro 6664
            'obs se não Encontrar o fornecedor pelo Codigo erro 6675
            'obs se não Encontrar o fornecedor pelo Codigo erro 6660
            
            Fornecedor.Text = objFornecedor.sNomeReduzido

            'Função que a Coleção de nomes e Codidos das Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 111351

            'Função que Preenche a Combo de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Função que Seleciona a Filial na Combo de Filiais
            Call CF("Filial_Seleciona", Filial, iCodFilial)

         Else

            'Se o Fornecedor não estiver preenchido limpa a Combo de Filial
            Filial.Clear

        End If

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 111350, 111351

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158972)

    End Select

    Exit Sub


End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.ListIndex >= 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 111366

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica se o Fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 111367

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 111368

        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = Fornecedor.Text

            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 111369

            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            Error 111370

        End If

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 111371

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 111366, 111368, 111369 'Tratados nas Rotinas chamadas

        Case 111370
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 111367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 111371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158973)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorCreditado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorCreditado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorCreditado_Validate

    'Verifica se Valor Creditado esta Preenchido
    If Len(Trim(ValorCreditado.Text)) <> 0 Then

        'Função que Critica se o Valor Colocado é Positivo
        lErro = Valor_Positivo_Critica(ValorCreditado.Text)
        If lErro <> SUCESSO Then gError 111352


        'Formata o Valor para 2 casas decimais arredondando se necessário
        ValorCreditado.Text = Format(ValorCreditado.Text, "Fixed")

    End If

    Exit Sub

Erro_ValorCreditado_Validate:

    Cancel = True

    Select Case gErr

        Case 111352

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158974)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimento_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataVencimento)

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a data de vencimento esta Preenchida
    If Len(Trim(DataVencimento.ClipText)) <> 0 Then

        'Função que critica se a data é válida
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 111353

    End If

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True

    Select Case gErr

        Case 111353

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158975)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimento_DownClick()
'Função que serve para decrementar a DataVencimento

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_DownClick

    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 111354

    Exit Sub

Erro_UpDownDataVencimento_DownClick:

    Select Case gErr

        Case 111354
            'Erro Tratado Dentro da Função Chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158976)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimento_UpClick()
'Função que serve para imcrementar a DataVencimento

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_UpClick

    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 111355

    Exit Sub

Erro_UpDownDataVencimento_UpClick:

    Select Case gErr

        Case 111355

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158977)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Botão que inicialza o proceso de Gravação

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Função que Grava o Registro na Tabela
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 111356

    'Função que Limpa a Tela de devolução de cheques
    Call Limpa_Tela_DevolucaoCheque

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 111356

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158978)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Função que Grava o Registro no BD

Dim lErro As Long
Dim objDevCheque As New ClassDevCheque

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 111357
    If Len(Trim(Data.ClipText)) = 0 Then gError 111358
    If Len(Trim(SeqChq.Text)) = 0 Then gError 111359

    'Chama a Função que move da tela para a memória
    lErro = Move_Tela_Memoria(objDevCheque)
    If lErro <> SUCESSO Then gError 111360

    'Função que se a data contábil é diferente da data passada como parâmtro se for Erro
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 111361

    'Função que Grava cheque Devolvido no Banco de dados
    lErro = CF("DevolucaoCheque_Grava", objDevCheque, objContabil)
    If lErro <> SUCESSO Then gError 111362
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
    
    Select Case gErr

        Case 111357
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 111358
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 111359
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", gErr)

        Case 111361, 111362

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158979)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Botão que começa o Processo de Exclusão

Dim lErro As Long
Dim objDevCheque As New ClassDevCheque

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o Codigo esta preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 111363

    'Move o codigo da filial da Empresa para o Obj
    objDevCheque.iFilialEmpresa = giFilialEmpresa

    'Move o Codigo para dentro do obj
    objDevCheque.lCodigo = StrParaLong(Codigo.Text)

    'Chama a Função que Exclui op Cheque no Banco de Dados
    lErro = CF("DevolucaoCheque_Exclui", objDevCheque, objContabil)
    If lErro <> SUCESSO Then gError 111364

    'Função que Limpa a Tela de devolução de cheques
    Call Limpa_Tela_DevolucaoCheque

    iAlterado = 0

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 111363, 111364

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158980)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Chama a Função que Testa e Salva
    Call Teste_Salva(Me, iAlterado)

    'Chama a Função que Limpa a Tela de Devolução de Cheques
    Call Limpa_Tela_DevolucaoCheque

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158981)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_DevolucaoCheque()
'Função que Limpa a Tela

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_DevolucaoCheque

    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    Data.PromptInclude = False

    'Move a Data atual para o campo de data na tela
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")

    Data.PromptInclude = True

    'Limpa os Caption's Relacionados ao Cheque que será devolvido
    Banco.Caption = ""
    Agencia.Caption = ""
    ContaCorrente.Caption = ""
    Numero.Caption = ""
    DataBomPara.Caption = ""
    ValorChq.Caption = ""
    TipoBordero.Caption = ""
    NumBordero.Caption = ""
    DataDeposito.Caption = ""
    
    Filial.Clear
    
    'Limpar o GridContabil
    Call objContabil.Contabil_Limpa_GridContabil
    
    'Zera iAlterado
    iAlterado = 0
    
    Exit Sub

Erro_Limpa_Tela_DevolucaoCheque:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158982)

    End Select

    Exit Sub

End Sub


Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoDevCheque = Nothing
    Set objEventoCheque = Nothing
    Set objEventoFornecedor = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid = Nothing
    Set objContabil = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Function Traz_DevCheque_Tela(objDevCheque As ClassDevCheque) As Long
'Função que Traz os Dados do Cheque para a Tela

Dim lErro As Long
Dim objChqBord As New ClassChequeBordero
Dim objChequePre As New ClassChequePre

On Error GoTo Erro_Traz_DevCheque_Tela

    'Chama a Função que Limpa a Tela de Cheque
    Call Limpa_Tela_DevolucaoCheque

    'Função que Lê os Cheques os Cheques que foram devolvidos
    lErro = CF("DevCheque_Le", objDevCheque)
    If lErro <> SUCESSO And lErro <> 109875 Then gError 111367
    
    'se naum existe cheque
    If lErro = 109875 Then gError 109880
    
    If objDevCheque.lNumIntChqBord <> 0 Then
    
        objChqBord.iFilialEmpresa = giFilialEmpresa
        objChqBord.lNumIntDoc = objDevCheque.lNumIntChqBord
    
        'Função que lê o Cheque Vinculado a um borderô para descobriri qual é o numero Interno do Cheque
        lErro = CF("ChequeBordero_Le", objChqBord)
        If lErro <> SUCESSO And lErro <> 109879 Then gError 111368
        
        'se naum existe bordero
        If lErro = 109879 Then gError 109881
        
        objChequePre.lNumIntCheque = objChqBord.lNumIntCheque
        
    Else
        objChequePre.lNumIntCheque = objDevCheque.lNumIntCheque
        objChequePre.iFilialEmpresa = giFilialEmpresa
    End If
    
    'Função que lê os Cheques
    lErro = CF("ChequePre_Le", objChequePre)
    If lErro <> SUCESSO And lErro <> 17642 Then gError 111369
    
    objChequePre.iTipoBordero = objChqBord.iTipoBordero
    objChequePre.lNumBordero = objChqBord.lNumBordero
    
    'naum existe cheque
    If lErro = 17642 Then gError 109882
    
    'Preenche a Tela com os dados do cheque devolvido
    lErro = Preenche_Dados_Cheque(objChequePre)
    If lErro <> SUCESSO Then gError 112051
    
    If lErro = 109971 Then gError 112052
    
    SeqChq.Text = objChequePre.lSequencialBack
    
    Codigo.PromptInclude = False
    Codigo.Text = objDevCheque.lCodigo
    Codigo.PromptInclude = True
    
    Data.PromptInclude = False
    Data.Text = Format(objDevCheque.dtData, "dd/mm/yy")
    Data.PromptInclude = True
    
    If objDevCheque.lFornecedor > 0 Then
        Fornecedor.Text = objDevCheque.lFornecedor
        Call Fornecedor_Validate(False)
    End If
    
    If objDevCheque.iFilial > 0 Then
        Filial.Text = objDevCheque.iFilial
        Call Filial_Validate(False)
    End If
    
    If objDevCheque.dValorCredito <> 0 Then ValorCreditado.Text = objDevCheque.dValorCredito
        
    If objDevCheque.dtDataVencimento <> DATA_NULA Then
        DataVencimento.PromptInclude = False
        DataVencimento.Text = Format(objDevCheque.dtDataVencimento, "dd/mm/yy")
        DataVencimento.PromptInclude = True
    End If
    
    Call objContabil.Contabil_Traz_Doc_Tela(objDevCheque.lNumIntDoc)
    
    Traz_DevCheque_Tela = SUCESSO

    Exit Function

Erro_Traz_DevCheque_Tela:
    
    Traz_DevCheque_Tela = gErr
    
    Select Case gErr

        Case 111367, 111368, 111369, 112051
            
        Case 109880
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_DEVCHEQUE", gErr, objDevCheque.lCodigo)
        
        Case 109881
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_BORDERO", gErr, objChqBord.lNumBordero)
        
        Case 109882
            'Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_CHEQUE", gErr, objChequePre.lNumero)
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_DEV_EXCLUIDO", gErr)
        
        Case 112052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDEROSCHEQUESPRE_NAO_EXISTE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158983)

    End Select

    Exit Function

End Function

Function Traz_Dados_BorderoDesc_Tela(objChqBord As ClassChequeBordero, objChequePre As ClassChequePre) As Long
'Função que Traz os Dados do bordero de desconto para a tela

Dim lErro As Long
Dim objDevCheque As New ClassDevCheque
Dim objChequePre1 As New ClassChequePre
Dim objBorderoDesc As New ClassBorderoDescChq
Dim dValorTotal As Double
Dim objCobrador As New ClassCobrador

On Error GoTo Erro_Traz_Dados_BorderoDesc_Tela

    'se for bordero de desconto
    If objChequePre.iTipoBordero = TIPO_BORDERO_DESCONTO Then
        
        objBorderoDesc.lNumBordero = objChequePre.lNumBordero
        objBorderoDesc.iFilialEmpresa = objChequePre.iFilialEmpresa
        
        lErro = CF("BorderoDesc_Chq_Le", objBorderoDesc)
        If lErro <> SUCESSO And lErro <> 109886 Then gError 109892
            
        'se naum existe bordero de desconto
        If lErro = 109886 Then gError 109893
        
        objCobrador.iCodigo = objBorderoDesc.iCobrador
        
        'TABELA DE COBRADORES PARA TRAZER O FORNECEDOR
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 109894
        
        If lErro = 19250 Then gError 112000
        
        'o cobrador naum está relacionado a um fornecedor --> erro.
        If objCobrador.lFornecedor = 0 Then gError 112001
        
        Fornecedor.Text = objCobrador.lFornecedor
        Call Fornecedor_Validate(False)
        
        Filial.Text = objCobrador.iFilial
        Call Filial_Validate(False)
        
        lErro = CF("ChequePre_Le_BorderoDesc", objBorderoDesc)
        If lErro <> SUCESSO And lErro <> 109890 Then gError 109894
        
        'não existe cheque relacionado ao bordero
        If lErro = 109890 Then gError 109895
        
        'soma os valores de todos os cheques associados ao bordero
        For Each objChequePre1 In objBorderoDesc.colchequepre
            dValorTotal = dValorTotal + objChequePre1.dValor
        Next
        
        'valor creditado
        ValorCreditado.Text = Format((objChequePre.dValor * objBorderoDesc.dValorCredito) / dValorTotal, "standard")
        
    End If
        
    Traz_Dados_BorderoDesc_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_BorderoDesc_Tela:
    
    Traz_Dados_BorderoDesc_Tela = gErr
    
    Select Case gErr

        Case 109892, 109894
            
        Case 109893
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_BORDERO", gErr, objChqBord.lNumBordero)
        
        Case 109895
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_CHEQUE_BORDERO", gErr, objChequePre.lNumBordero)
        
        Case 112000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrador.iCodigo)
        
        Case 112001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ASSOCIADO_FORNECEDOR", gErr, objCobrador.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158984)

    End Select
    
    Exit Function
    
End Function

Function Move_Tela_Memoria(objDevCheque As ClassDevCheque) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    objDevCheque.iFilialEmpresa = giFilialEmpresa
    objDevCheque.lCodigo = StrParaLong(Codigo.Text)
    
    If Len(Trim(Data.ClipText)) > 0 Then
        objDevCheque.dtData = Data.Text
    Else
        objDevCheque.dtData = DATA_NULA
    End If
        
    objDevCheque.lSeqChq = StrParaLong(SeqChq.Text)
    
    If Len(Trim(Fornecedor.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = Fornecedor.Text
            
        'busca o códigodo fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 109896
        
        'naum existe fornecedor
        If lErro = 6681 Then gError 109897
        
        objDevCheque.lFornecedor = objFornecedor.lCodigo
        
    End If
    
    If Filial.ListIndex <> -1 Then objDevCheque.iFilial = Codigo_Extrai(Filial.Text)
        
    If Len(Trim(DataVencimento.ClipText)) > 0 Then
        objDevCheque.dtDataVencimento = DataVencimento.Text
    Else
        objDevCheque.dtDataVencimento = DATA_NULA
    End If
        
    objDevCheque.dValorCredito = StrParaLong(ValorCreditado.Text)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:
    
    Move_Tela_Memoria = gErr
    
    Select Case gErr

        Case 109896
            
        Case 109897
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158985)

    End Select
    

End Function

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If
    
    'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
    If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

    'If iFrameAtual = TAB_Contabilizacao Then Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_CONTABILIZACAO

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is SeqChq Then
            Call LabelSeqChq_Click
        End If
    
    End If
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 112046

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 112047

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 112046, 112047
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158986)

    End Select

    Exit Function

End Function

'inicio contabilidade

Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Private Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Private Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)

End Sub

Private Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Private Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Private Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Private Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Private Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Private Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Private Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Private Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Private Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Private Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Private Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Private Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Private Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Private Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Private Sub CTBAglutina_Click()

    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Private Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Private Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Private Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Private Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Private Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Private Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Private Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Private Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Private Sub CTBBotaoImprimir_Click()

    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Private Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick

End Sub

Private Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Private Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click

End Sub

Private Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click

End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case VALOR_CHQ
            objMnemonicoValor.colValor.Add StrParaDbl(ValorChq.Caption)

        Case Else
            Error 39678
            
    End Select

    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err
    
    Select Case Err
        
        Case 39678
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155393)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Devolução de Cheque"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "DevolucaoCheque"

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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******


Private Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Private Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub


