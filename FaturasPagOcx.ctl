VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FaturasPagOcx 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   KeyPreview      =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9585
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4695
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   780
      Width           =   9060
      Begin VB.Frame SSFrame1 
         Caption         =   "Notas Fiscais"
         Height          =   3285
         Left            =   60
         TabIndex        =   35
         Top             =   1395
         Width           =   8985
         Begin MSMask.MaskEdBox ValorNFBruto 
            Height          =   225
            Left            =   6600
            TabIndex        =   100
            Top             =   555
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FilialNF 
            Height          =   225
            Left            =   6525
            TabIndex        =   47
            Top             =   1155
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FilialFornecedor 
            Height          =   225
            Left            =   5475
            TabIndex        =   46
            Top             =   1860
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.CommandButton BotaoConsultarNFiscal 
            Caption         =   "Consultar Nota Fiscal"
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
            Left            =   7560
            TabIndex        =   11
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton BotaoAtualizarNFs 
            Caption         =   "Atualizar Lista de NFs"
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
            Left            =   6165
            TabIndex        =   10
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton NFRegistrar 
            Caption         =   "Cadastrar Nota Fiscal..."
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
            Left            =   4755
            TabIndex        =   9
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox Selecionada 
            Caption         =   "Check1"
            Height          =   225
            Left            =   3195
            TabIndex        =   16
            Top             =   360
            Width           =   1035
         End
         Begin MSMask.MaskEdBox ValorNF 
            Height          =   225
            Left            =   4620
            TabIndex        =   15
            Top             =   960
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataEmissaoNF 
            Height          =   225
            Left            =   5370
            TabIndex        =   14
            Top             =   645
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumNotaFiscal 
            Height          =   225
            Left            =   5085
            TabIndex        =   13
            Top             =   1425
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridNF 
            Height          =   2220
            Left            =   90
            TabIndex        =   8
            Top             =   225
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   3916
            _Version        =   393216
            Rows            =   50
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Valor (R$):"
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
            TabIndex        =   36
            Top             =   2970
            Width           =   930
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "No. de NFs:"
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
            Left            =   2400
            TabIndex        =   37
            Top             =   2970
            Width           =   1035
         End
         Begin VB.Label ValorTotalNFSelecionadas 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1050
            TabIndex        =   38
            Top             =   2970
            Width           =   1215
         End
         Begin VB.Label NumNFSelecionadas 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3450
            TabIndex        =   39
            Top             =   2970
            Width           =   1215
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Cabeçalho"
         Height          =   1395
         Left            =   60
         TabIndex        =   34
         Top             =   -15
         Width           =   8970
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   6675
            TabIndex        =   1
            Top             =   210
            Width           =   1815
         End
         Begin MSMask.MaskEdBox NumTitulo 
            Height          =   300
            Left            =   1620
            TabIndex        =   2
            Top             =   585
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "999999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1620
            TabIndex        =   0
            Top             =   210
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2670
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   960
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1620
            TabIndex        =   4
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotal 
            Height          =   300
            Left            =   6675
            TabIndex        =   3
            Top             =   585
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorIRRF 
            Height          =   300
            Left            =   4095
            TabIndex        =   6
            Top             =   975
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Diferenca 
            Height          =   300
            Left            =   6660
            TabIndex        =   7
            Top             =   975
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Diferença:"
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
            Left            =   5670
            TabIndex        =   97
            Top             =   1020
            Width           =   900
         End
         Begin VB.Label Label16 
            Caption         =   "Retenções:"
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
            Left            =   3075
            TabIndex        =   48
            Top             =   1020
            Width           =   990
         End
         Begin VB.Label Label6 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   6090
            TabIndex        =   40
            Top             =   645
            Width           =   510
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Left            =   750
            TabIndex        =   41
            Top             =   1020
            Width           =   765
         End
         Begin VB.Label Label12 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   6060
            TabIndex        =   42
            Top             =   270
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
            Left            =   780
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   43
            Top             =   660
            Width           =   720
         End
         Begin VB.Label FornecedorLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   44
            Top             =   270
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4680
      Index           =   2
      Left            =   135
      TabIndex        =   17
      Top             =   795
      Visible         =   0   'False
      Width           =   9090
      Begin VB.Frame SSFrame3 
         Caption         =   "Parcelas"
         Height          =   4170
         Left            =   135
         TabIndex        =   33
         Top             =   -30
         Width           =   8910
         Begin VB.ComboBox MotivoDiferenca 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4350
            TabIndex        =   98
            Top             =   3180
            Width           =   2235
         End
         Begin MSMask.MaskEdBox ValorOriginal 
            Height          =   225
            Left            =   2925
            TabIndex        =   99
            Top             =   3120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   14
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.ComboBox ComboCobrador 
            Height          =   315
            Left            =   2940
            TabIndex        =   95
            Top             =   2400
            Width           =   2295
         End
         Begin VB.ComboBox ComboPortador 
            Height          =   315
            Left            =   15
            TabIndex        =   94
            Top             =   2430
            Width           =   2445
         End
         Begin MSMask.MaskEdBox CodigodeBarras 
            Height          =   315
            Left            =   360
            TabIndex        =   93
            Top             =   1710
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   57
            Mask            =   "#####.#####.#####.######.#####.######.#.#################"
            PromptChar      =   " "
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
            Left            =   4335
            TabIndex        =   19
            Top             =   300
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   2355
            TabIndex        =   18
            Top             =   255
            Width           =   1815
         End
         Begin VB.ComboBox TipoCobranca 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            TabIndex        =   24
            Top             =   930
            Width           =   1815
         End
         Begin VB.CheckBox Suspenso 
            Caption         =   "Check1"
            Height          =   225
            Left            =   6360
            TabIndex        =   25
            Top             =   960
            Width           =   900
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   225
            Left            =   1980
            TabIndex        =   22
            Top             =   750
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   3150
            TabIndex        =   23
            Top             =   795
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   14
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   780
            TabIndex        =   21
            Top             =   780
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
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   3435
            Left            =   135
            TabIndex        =   26
            Top             =   690
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   6059
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label CondPagtoLabel 
            Caption         =   "Condição de Pagamento:"
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
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   45
            Top             =   285
            Width           =   2175
         End
      End
      Begin MSMask.MaskEdBox CodBarrasAvulso 
         Height          =   315
         Left            =   1740
         TabIndex        =   20
         Top             =   4260
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   57
         Mask            =   "#####.#####.#####.######.#####.######.#.#################"
         PromptChar      =   " "
      End
      Begin VB.Label CodigoBarras 
         AutoSize        =   -1  'True
         Caption         =   "Código de Barras:"
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
         Left            =   150
         TabIndex        =   96
         Top             =   4305
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4605
      Index           =   3
      Left            =   165
      TabIndex        =   49
      Top             =   855
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4680
         TabIndex        =   92
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
         Left            =   3450
         TabIndex        =   63
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   58
         Top             =   3450
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   62
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   61
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
            Left            =   1125
            TabIndex        =   60
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
            Left            =   240
            TabIndex        =   59
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   57
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   56
         Top             =   2175
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   55
         Top             =   2565
         Width           =   870
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
         Height          =   270
         Left            =   7770
         TabIndex        =   53
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   930
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
         Height          =   270
         Left            =   6330
         TabIndex        =   51
         Top             =   90
         Width           =   1245
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
         Left            =   6330
         TabIndex        =   50
         Top             =   405
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   54
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   69
         Top             =   525
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
         Left            =   5580
         TabIndex        =   70
         Top             =   135
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
         Left            =   3825
         TabIndex        =   71
         Top             =   120
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
         Left            =   0
         TabIndex        =   72
         Top             =   1185
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
         Left            =   6330
         TabIndex        =   73
         Top             =   1515
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
         Left            =   6330
         TabIndex        =   74
         Top             =   1515
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
         TabIndex        =   91
         Top             =   165
         Width           =   450
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
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
         TabIndex        =   90
         Top             =   165
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
         TabIndex        =   89
         Top             =   555
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   88
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   87
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBLabelTotais 
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
         TabIndex        =   86
         Top             =   3045
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
         TabIndex        =   85
         Top             =   1275
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
         Top             =   945
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
         TabIndex        =   81
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   80
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   79
         Top             =   570
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
         Left            =   4230
         TabIndex        =   78
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   77
         Top             =   120
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
         TabIndex        =   76
         Top             =   165
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
         Left            =   6360
         TabIndex        =   75
         Top             =   720
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7260
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FaturasPagOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FaturasPagOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FaturasPagOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FaturasPagOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5175
      Left            =   90
      TabIndex        =   32
      Top             =   420
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagamento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "FaturasPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTFaturasPag
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTFaturasPag
    Set objCT.objUserControl = Me
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoAtualizarNFs_Click()
     Call objCT.BotaoAtualizarNFs_Click
End Sub

Private Sub BotaoConsultarNFiscal_Click()
     Call objCT.BotaoConsultarNFiscal_Click
End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)
     Call objCT.CondicaoPagamento_Validate(Cancel)
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub FilialNF_Change()
     Call objCT.FilialNF_Change
End Sub

Private Sub FornecedorLabel_Click()
     Call objCT.FornecedorLabel_Click
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub NumTitulo_GotFocus()
     Call objCT.NumTitulo_GotFocus
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Private Sub CondPagtoLabel_Click()
     Call objCT.CondPagtoLabel_Click
End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long
     Trata_Parametros = objCT.Trata_Parametros(objTituloPagar)
End Function

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub NumTitulo_Validate(Cancel As Boolean)
     Call objCT.NumTitulo_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub UpDownEmissao_Change()
     Call objCT.UpDownEmissao_Change
End Sub

Private Sub ValorTotal_Change()
     Call objCT.ValorTotal_Change
End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)
     Call objCT.ValorTotal_Validate(Cancel)
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub CondicaoPagamento_Click()
     Call objCT.CondicaoPagamento_Click
End Sub

Private Sub Selecionada_Click()
     Call objCT.Selecionada_Click
End Sub

Private Sub NFRegistrar_Click()
     Call objCT.NFRegistrar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub GridNF_Click()
     Call objCT.GridNF_Click
End Sub

Private Sub GridNF_EnterCell()
     Call objCT.GridNF_EnterCell
End Sub

Private Sub GridNF_GotFocus()
     Call objCT.GridNF_GotFocus
End Sub

Private Sub GridNF_KeyPress(KeyAscii As Integer)
     Call objCT.GridNF_KeyPress(KeyAscii)
End Sub

Private Sub GridNF_LeaveCell()
     Call objCT.GridNF_LeaveCell
End Sub

Private Sub GridNF_Validate(Cancel As Boolean)
     Call objCT.GridNF_Validate(Cancel)
End Sub

Private Sub GridNF_RowColChange()
     Call objCT.GridNF_RowColChange
End Sub

Private Sub GridNF_Scroll()
     Call objCT.GridNF_Scroll
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

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub NumTitulo_Change()
     Call objCT.NumTitulo_Change
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

Private Sub DataVencimentoReal_Change()
     Call objCT.DataVencimentoReal_Change
End Sub

Private Sub DataVencimentoReal_GotFocus()
     Call objCT.DataVencimentoReal_GotFocus
End Sub

Private Sub DataVencimentoReal_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimentoReal_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimentoReal_Validate(Cancel As Boolean)
     Call objCT.DataVencimentoReal_Validate(Cancel)
End Sub

Private Sub TipoCobranca_Change()
     Call objCT.TipoCobranca_Change
End Sub

Private Sub TipoCobranca_Click()
     Call objCT.TipoCobranca_Click
End Sub

Private Sub TipoCobranca_GotFocus()
     Call objCT.TipoCobranca_GotFocus
End Sub

Private Sub TipoCobranca_KeyPress(KeyAscii As Integer)
     Call objCT.TipoCobranca_KeyPress(KeyAscii)
End Sub

Private Sub TipoCobranca_Validate(Cancel As Boolean)
     Call objCT.TipoCobranca_Validate(Cancel)
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

Private Sub Suspenso_Click()
     Call objCT.Suspenso_Click
End Sub

Private Sub Suspenso_GotFocus()
     Call objCT.Suspenso_GotFocus
End Sub

Private Sub Suspenso_KeyPress(KeyAscii As Integer)
     Call objCT.Suspenso_KeyPress(KeyAscii)
End Sub

Private Sub Suspenso_Validate(Cancel As Boolean)
     Call objCT.Suspenso_Validate(Cancel)
End Sub

Private Sub NumNotaFiscal_Change()
     Call objCT.NumNotaFiscal_Change
End Sub

Private Sub NumNotaFiscal_GotFocus()
     Call objCT.NumNotaFiscal_GotFocus
End Sub

Private Sub NumNotaFiscal_KeyPress(KeyAscii As Integer)
     Call objCT.NumNotaFiscal_KeyPress(KeyAscii)
End Sub

Private Sub NumNotaFiscal_Validate(Cancel As Boolean)
     Call objCT.NumNotaFiscal_Validate(Cancel)
End Sub

Private Sub DataEmissaoNF_Change()
     Call objCT.DataEmissaoNF_Change
End Sub

Private Sub DataEmissaoNF_GotFocus()
     Call objCT.DataEmissaoNF_GotFocus
End Sub

Private Sub DataEmissaoNF_KeyPress(KeyAscii As Integer)
     Call objCT.DataEmissaoNF_KeyPress(KeyAscii)
End Sub

Private Sub DataEmissaoNF_Validate(Cancel As Boolean)
     Call objCT.DataEmissaoNF_Validate(Cancel)
End Sub

Private Sub Selecionada_GotFocus()
     Call objCT.Selecionada_GotFocus
End Sub

Private Sub Selecionada_KeyPress(KeyAscii As Integer)
     Call objCT.Selecionada_KeyPress(KeyAscii)
End Sub

Private Sub Selecionada_Validate(Cancel As Boolean)
     Call objCT.Selecionada_Validate(Cancel)
End Sub

Private Sub ValorNF_Change()
     Call objCT.ValorNF_Change
End Sub

Private Sub ValorNF_GotFocus()
     Call objCT.ValorNF_GotFocus
End Sub

Private Sub ValorNF_KeyPress(KeyAscii As Integer)
     Call objCT.ValorNF_KeyPress(KeyAscii)
End Sub

Private Sub ValorNF_Validate(Cancel As Boolean)
     Call objCT.ValorNF_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub
Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub
Private Sub ValorTotalNFSelecionadas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotalNFSelecionadas, Source, X, Y)
End Sub
Private Sub ValorTotalNFSelecionadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotalNFSelecionadas, Button, Shift, X, Y)
End Sub
Private Sub NumNFSelecionadas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumNFSelecionadas, Source, X, Y)
End Sub
Private Sub NumNFSelecionadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumNFSelecionadas, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub
Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub
Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub
Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub
Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub
Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub
Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub
Private Sub Opcao_BeforeClick(Cancel As Integer)
     Call objCT.Opcao_BeforeClick(Cancel)
End Sub

Private Sub Fornecedor_Preenche()
     Call objCT.Fornecedor_Preenche
End Sub

Private Sub CTBGerencial_Click()
     Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
     Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
     Call objCT.CTBGerencial_Validate(Cancel)
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

Public Sub CTBBotaoModeloPadrao_Click()
    Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Public Sub CTBModelo_Click()
    Call objCT.CTBModelo_Click
End Sub

Public Sub CTBGridContabil_Click()
    Call objCT.CTBGridContabil_Click
End Sub

Public Sub CTBGridContabil_EnterCell()
    Call objCT.CTBGridContabil_EnterCell
End Sub

Public Sub CTBGridContabil_GotFocus()
    Call objCT.CTBGridContabil_GotFocus
End Sub

Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
    Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Public Sub CTBGridContabil_LeaveCell()
    Call objCT.CTBGridContabil_LeaveCell
End Sub

Public Sub CTBGridContabil_Validate(Cancel As Boolean)
    Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Public Sub CTBGridContabil_RowColChange()
    Call objCT.CTBGridContabil_RowColChange
End Sub

Public Sub CTBGridContabil_Scroll()
    Call objCT.CTBGridContabil_Scroll
End Sub

Public Sub CTBConta_Change()
    Call objCT.CTBConta_Change
End Sub

Public Sub CTBConta_GotFocus()
    Call objCT.CTBConta_GotFocus
End Sub

Public Sub CTBConta_KeyPress(KeyAscii As Integer)
    Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Public Sub CTBConta_Validate(Cancel As Boolean)
    Call objCT.CTBConta_Validate(Cancel)
End Sub

Public Sub CTBCcl_Change()
    Call objCT.CTBCcl_Change
End Sub

Public Sub CTBCcl_GotFocus()
    Call objCT.CTBCcl_GotFocus
End Sub

Public Sub CTBCcl_KeyPress(KeyAscii As Integer)
    Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Public Sub CTBCcl_Validate(Cancel As Boolean)
    Call objCT.CTBCcl_Validate(Cancel)
End Sub

Public Sub CTBCredito_Change()
    Call objCT.CTBCredito_Change
End Sub

Public Sub CTBCredito_GotFocus()
    Call objCT.CTBCredito_GotFocus
End Sub

Public Sub CTBCredito_KeyPress(KeyAscii As Integer)
    Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Public Sub CTBCredito_Validate(Cancel As Boolean)
    Call objCT.CTBCredito_Validate(Cancel)
End Sub

Public Sub CTBDebito_Change()
    Call objCT.CTBDebito_Change
End Sub

Public Sub CTBDebito_GotFocus()
    Call objCT.CTBDebito_GotFocus
End Sub

Public Sub CTBDebito_KeyPress(KeyAscii As Integer)
    Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Public Sub CTBDebito_Validate(Cancel As Boolean)
    Call objCT.CTBDebito_Validate(Cancel)
End Sub

Public Sub CTBSeqContraPartida_Change()
    Call objCT.CTBSeqContraPartida_Change
End Sub

Public Sub CTBSeqContraPartida_GotFocus()
    Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
    Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
    Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Public Sub CTBHistorico_Change()
    Call objCT.CTBHistorico_Change
End Sub

Public Sub CTBHistorico_GotFocus()
    Call objCT.CTBHistorico_GotFocus
End Sub

Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)
    Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Public Sub CTBHistorico_Validate(Cancel As Boolean)
    Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Public Sub CTBLancAutomatico_Click()
    Call objCT.CTBLancAutomatico_Click
End Sub

Public Sub CTBAglutina_Click()
    Call objCT.CTBAglutina_Click
End Sub

Public Sub CTBAglutina_GotFocus()
    Call objCT.CTBAglutina_GotFocus
End Sub

Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)
    Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Public Sub CTBAglutina_Validate(Cancel As Boolean)
    Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
    Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
    Call objCT.CTBTvwContas_Expand(Node)
End Sub

Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Public Sub CTBListHistoricos_DblClick()
    Call objCT.CTBListHistoricos_DblClick
End Sub

Public Sub CTBBotaoLimparGrid_Click()
    Call objCT.CTBBotaoLimparGrid_Click
End Sub

Public Sub CTBLote_Change()
    Call objCT.CTBLote_Change
End Sub

Public Sub CTBLote_GotFocus()
    Call objCT.CTBLote_GotFocus
End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)
    Call objCT.CTBLote_Validate(Cancel)
End Sub

Public Sub CTBDataContabil_Change()
    Call objCT.CTBDataContabil_Change
End Sub

Public Sub CTBDataContabil_GotFocus()
    Call objCT.CTBDataContabil_GotFocus
End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)
    Call objCT.CTBDataContabil_Validate(Cancel)
End Sub



Private Sub CodigodeBarras_Change()
     Call objCT.CodigodeBarras_Change
End Sub

Private Sub CodigodeBarras_GotFocus()
    Call objCT.CodigodeBarras_GotFocus
End Sub

Private Sub CodigodeBarras_KeyPress(KeyAscii As Integer)
    Call objCT.CodigodeBarras_KeyPress(KeyAscii)
End Sub

Private Sub CodigodeBarras_Validate(Cancel As Boolean)
    Call objCT.CodigodeBarras_Validate(Cancel)
End Sub

Private Sub CodBarrasAvulso_Validate(Cancel As Boolean)
     Call objCT.CodBarrasAvulso_Validate(Cancel)
End Sub

Private Sub CodBarrasAvulso_GotFocus()
     Call objCT.CodBarrasAvulso_GotFocus
End Sub

Private Sub CodBarrasAvulso_Change()
    Call objCT.CodBarrasAvulso_Change
End Sub


Private Sub ComboCobrador_Click()
     Call objCT.ComboCobrador_Click
End Sub

Private Sub ComboCobrador_Change()
     Call objCT.ComboCobrador_Change
End Sub

Private Sub ComboCobrador_GotFocus()
    Call objCT.ComboCobrador_GotFocus
End Sub

Private Sub ComboCobrador_KeyPress(KeyAscii As Integer)
    Call objCT.ComboCobrador_KeyPress(KeyAscii)
End Sub

Private Sub ComboCobrador_Validate(Cancel As Boolean)
    Call objCT.ComboCobrador_Validate(Cancel)
End Sub

Private Sub ComboPortador_Click()
     Call objCT.ComboPortador_Click
End Sub

Private Sub ComboPortador_Change()
     Call objCT.ComboPortador_Change
End Sub

Private Sub ComboPortador_GotFocus()
    Call objCT.ComboPortador_GotFocus
End Sub

Private Sub ComboPortador_KeyPress(KeyAscii As Integer)
    Call objCT.ComboPortador_KeyPress(KeyAscii)
End Sub

Private Sub ComboPortador_Validate(Cancel As Boolean)
    Call objCT.ComboPortador_Validate(Cancel)
End Sub

Public Sub ValorOriginal_Change()
    Call objCT.ValorOriginal_Change
End Sub

Public Sub ValorOriginal_GotFocus()
    Call objCT.ValorOriginal_GotFocus
End Sub

Public Sub ValorOriginal_KeyPress(KeyAscii As Integer)
    Call objCT.ValorOriginal_KeyPress(KeyAscii)
End Sub

Public Sub ValorOriginal_Validate(Cancel As Boolean)
    Call objCT.ValorOriginal_Validate(Cancel)
End Sub

Public Sub MotivoDiferenca_Click()
    Call objCT.MotivoDiferenca_Click
End Sub

Public Sub MotivoDiferenca_GotFocus()
    Call objCT.MotivoDiferenca_GotFocus
End Sub

Public Sub MotivoDiferenca_KeyPress(KeyAscii As Integer)
    Call objCT.MotivoDiferenca_KeyPress(KeyAscii)
End Sub

Public Sub MotivoDiferenca_Validate(Cancel As Boolean)
    Call objCT.MotivoDiferenca_Validate(Cancel)
End Sub

Private Sub Diferenca_Validate(Cancel As Boolean)
    Call objCT.Diferenca_Validate(Cancel)
End Sub
