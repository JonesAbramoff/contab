VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      Caption         =   "Frame2"
      Height          =   4680
      Index           =   2
      Left            =   135
      TabIndex        =   13
      Top             =   795
      Visible         =   0   'False
      Width           =   9090
      Begin VB.Frame SSFrame3 
         Caption         =   "Parcelas"
         Height          =   4275
         Left            =   135
         TabIndex        =   29
         Top             =   210
         Width           =   8910
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
            TabIndex        =   15
            Top             =   300
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   2355
            TabIndex        =   14
            Top             =   255
            Width           =   1815
         End
         Begin VB.ComboBox TipoCobranca 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            TabIndex        =   19
            Top             =   930
            Width           =   1815
         End
         Begin VB.CheckBox Suspenso 
            Caption         =   "Check1"
            Height          =   225
            Left            =   6360
            TabIndex        =   20
            Top             =   960
            Width           =   900
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   225
            Left            =   1980
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   16
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
            TabIndex        =   21
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
            TabIndex        =   41
            Top             =   285
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4695
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   780
      Width           =   9060
      Begin VB.Frame SSFrame1 
         Caption         =   "Notas Fiscais"
         Height          =   2595
         Left            =   360
         TabIndex        =   31
         Top             =   2040
         Width           =   8175
         Begin MSMask.MaskEdBox FilialNF 
            Height          =   225
            Left            =   180
            TabIndex        =   44
            Top             =   225
            Width           =   560
            _ExtentX        =   979
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FilialFornecedor 
            Height          =   225
            Left            =   315
            TabIndex        =   43
            Top             =   525
            Width           =   780
            _ExtentX        =   1376
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
            Height          =   420
            Left            =   5910
            TabIndex        =   42
            Top             =   2040
            Width           =   2175
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
            Height          =   420
            Left            =   5910
            TabIndex        =   12
            Top             =   1552
            Width           =   2175
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
            Height          =   420
            Left            =   5910
            TabIndex        =   11
            Top             =   1065
            Width           =   2175
         End
         Begin VB.CheckBox Selecionada 
            Caption         =   "Check1"
            Height          =   225
            Left            =   3195
            TabIndex        =   9
            Top             =   360
            Width           =   1035
         End
         Begin MSMask.MaskEdBox ValorNF 
            Height          =   225
            Left            =   2250
            TabIndex        =   8
            Top             =   360
            Width           =   1005
            _ExtentX        =   1773
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
            Left            =   1170
            TabIndex        =   7
            Top             =   375
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
            Left            =   795
            TabIndex        =   6
            Top             =   195
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridNF 
            Height          =   2220
            Left            =   90
            TabIndex        =   10
            Top             =   225
            Width           =   5670
            _ExtentX        =   10001
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
            Left            =   5910
            TabIndex        =   32
            Top             =   270
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
            Left            =   5805
            TabIndex        =   33
            Top             =   645
            Width           =   1035
         End
         Begin VB.Label ValorTotalNFSelecionadas 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6870
            TabIndex        =   34
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label NumNFSelecionadas 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6870
            TabIndex        =   35
            Top             =   645
            Width           =   1215
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Cabeçalho"
         Height          =   1845
         Left            =   375
         TabIndex        =   30
         Top             =   105
         Width           =   8160
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5505
            TabIndex        =   2
            Top             =   330
            Width           =   1815
         End
         Begin MSMask.MaskEdBox NumTitulo 
            Height          =   300
            Left            =   1620
            TabIndex        =   3
            Top             =   840
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
            TabIndex        =   1
            Top             =   330
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2670
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1335
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1620
            TabIndex        =   5
            Top             =   1335
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
            Left            =   5505
            TabIndex        =   4
            Top             =   855
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
            Left            =   5505
            TabIndex        =   45
            Top             =   1350
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            Caption         =   "IR:"
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
            Left            =   5160
            TabIndex        =   46
            Top             =   1395
            Width           =   300
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
            Left            =   4920
            TabIndex        =   36
            Top             =   915
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
            TabIndex        =   37
            Top             =   1395
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
            Left            =   4890
            TabIndex        =   38
            Top             =   390
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
            TabIndex        =   39
            Top             =   915
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
            TabIndex        =   40
            Top             =   390
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4605
      Index           =   3
      Left            =   165
      TabIndex        =   47
      Top             =   855
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4680
         TabIndex        =   90
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
         TabIndex        =   61
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   56
         Top             =   3450
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   60
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   55
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   54
         Top             =   2175
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   53
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
         TabIndex        =   51
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
         Top             =   405
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   52
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   70
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
         TabIndex        =   71
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
         TabIndex        =   72
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
         Top             =   555
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   86
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   78
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   77
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
         TabIndex        =   76
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
         Top             =   720
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7260
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FaturasPagOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FaturasPagOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FaturasPagOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FaturasPagOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5175
      Left            =   90
      TabIndex        =   28
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

'Property Variables:
Dim m_Caption As String
Event Unload()

'inicio contabilidade
Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
Private Const NUMERO1 As String = "Numero_Nota_Fiscal"
Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALORTOTAL1 As String = "Valor_da_Nota"
Private Const VALOR_IR As String = "Valor_IRRF"
Private Const CONTA_DESP_ESTOQUE As String = "Conta_Desp_Estoque"
Private Const CONTA_DESP_EST_FORN As String = "Conta_Desp_Est_Forn"

'fim contabilidade

Public iAlterado As Integer
Private iFrameAtual As Integer
Private iFornecedorAlterado As Integer
Private iValorTotalAlterado As Integer
Private iEmissaoAlterada As Integer
Private iValorIRRFAlterado As Integer
Private sOldFornecedor As String
Private sOldFilial As String
Private sOldNumTitulo As String
Private dtOldDataEmissao As Date

Dim objGridNFiscais As AdmGrid
Dim iGrid_FilialNF_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Emissao_Col As Integer
Dim iGrid_ValorNF_Col As Integer
Dim iGrid_Selecionada_Col As Integer
Dim iGrid_FilialFornecedor_Col As Integer

Dim objGridParcelas As AdmGrid
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_VenctoReal_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_Cobranca_Col As Integer
Dim iGrid_Suspenso_Col As Integer

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1

'obs.: nao remover deste fonte
Private Const NUM_LINHAS_GRID_NF = 7 'numero minimo de linhas visiveis no grid de nfs

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Cobranca = 2
Private Const TAB_Contabilizacao = 3

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFornecedorAlterado = 0
    iEmissaoAlterada = 0
    iValorIRRFAlterado = 0
    
    Set objEventoFornecedor = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoCondPagto = New AdmEvento

    'Carrega a List da Combo de Tipos de Cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then Error 18597

    'Carrega a List da Combo de Condições de Pagamento
    lErro = Carrega_CondicaoPagamento()
    If lErro <> SUCESSO Then Error 18600

    'Inicializa o GridParcelas
    Set objGridParcelas = New AdmGrid
    lErro = Inicializa_GridParcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 18601

    'Inicializa o GridNF
    Set objGridNFiscais = New AdmGrid
    lErro = Inicializa_GridNF(objGridNFiscais)
    If lErro <> SUCESSO Then Error 18602
    
    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 184125
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 18597, 18600, 18601, 18602, 184125

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160143)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colNFPag As New ColNFsPag

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosPag"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas, colNFPag)
    If lErro <> SUCESSO Then Error 18782

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Fornecedor", objTituloPagar.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objTituloPagar.iFilial, 0, "Filial"
    colCampoValor.Add "NumTitulo", objTituloPagar.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objTituloPagar.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorTotal", objTituloPagar.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "NumParcelas", objTituloPagar.iNumParcelas, 0, "NumParcelas"
    colCampoValor.Add "NumIntDoc", objTituloPagar.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "CondicaoPagto", objTituloPagar.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "ValorIRRF", objTituloPagar.dValorIRRF, 0, "ValorIRRF"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "SiglaDocumento", OP_IGUAL, TIPODOC_FATURA_A_PAGAR
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 18782

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160144)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_Tela_Preenche

    objTituloPagar.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objTituloPagar.lNumIntDoc <> 0 Then

        'Carrega objTituloPagar com os dados passados em colCampoValor
        objTituloPagar.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objTituloPagar.iFilial = colCampoValor.Item("Filial").vValor
        objTituloPagar.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objTituloPagar.iNumParcelas = colCampoValor.Item("NumParcelas").vValor
        objTituloPagar.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objTituloPagar.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objTituloPagar.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        objTituloPagar.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        
        'Traz os dados da Fatura, passada em objTituloPagar, para a tela
        lErro = Traz_FaturaPagar_Tela(objTituloPagar)
        If lErro <> SUCESSO Then Error 18786

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 18786

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160145)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAtualizarNFs_Click()

    'força a atualizacao das lista de nfs no grid
    Call Atualiza_NotasFiscais(1)
        
End Sub

Private Sub BotaoConsultarNFiscal_Click()

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag
Dim sTela As String
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsultarNFiscal_Click

    'Verifica se uma linha do Grid foi selecionada
    If GridNF.Row <= 0 Then gError 79387
    
    'Critica se a linha selecionada está preenchida
    If Len(Trim(GridNF.TextMatrix(GridNF.Row, iGrid_Numero_Col))) = 0 Then gError 79388
    
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79389

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 79390

    objNFsPag.lFornecedor = objFornecedor.lCodigo
    objNFsPag.iFilial = Codigo_Extrai(Filial.Text)
    objNFsPag.lNumNotaFiscal = StrParaLong(GridNF.TextMatrix(GridNF.Row, iGrid_Numero_Col))
    objNFsPag.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(GridNF.TextMatrix(GridNF.Row, iGrid_Emissao_Col))) > 0 Then
        objNFsPag.dtDataEmissao = CDate(GridNF.TextMatrix(GridNF.Row, iGrid_Emissao_Col))
    Else
        objNFsPag.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (apenas na tabela de NFiscais em aberto)
    lErro = CF("NFPag_Le_Numero", objNFsPag)
    If lErro <> SUCESSO And lErro <> 18338 Then gError 79391

    'Le o Nome da Tela que originou este Título
    lErro = CF("Titulo_Le_DocumentoOriginal", objNFsPag.lNumIntDoc, CPR_NF_PAGAR, objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 58942 Then gError 79392
    
    If lErro = SUCESSO Then
        'Chama a Tela Estoque
        Call Chama_Tela(sTela, objNFiscal)
    Else
        Call Chama_Tela("NFPag_Consulta", objNFsPag)
    End If

    Exit Sub

Erro_BotaoConsultarNFiscal_Click:

    Select Case gErr

        Case 79389, 79391, 79392 'Tratado na Rotina chamada

        Case 79387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 79388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 79390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160146)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CondicaoPagamento_Validate

    'Verifica se a Condicao Pagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicao Pagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18660

    'Se não encontra, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 18661
        If lErro <> SUCESSO Then Error 18663

        'Testa se pode ser usada em Contas a Pagar
        If objCondicaoPagto.iEmPagamento = 0 Then Error 18784

        'Coloca na tela
        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida

        Call Recalcula_Cobranca
        
    End If

    'Não encontrou e é STRING
    If lErro = 6731 Then Error 18664

    Exit Sub

Erro_CondicaoPagamento_Validate:

    Cancel = True
    
    Select Case Err

       Case 18660, 18661, 18662

       Case 18663
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 18664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagamento.Text)
        
        Case 18784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", Err, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160147)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_GotFocus()
Dim iEmissaoAux As Integer
    
    iEmissaoAux = iEmissaoAlterada
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    iEmissaoAlterada = iEmissaoAux

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim sFornecedor As String
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18624

    'Não encontrou, mas extrai o código
    If lErro = 6730 Then

        'Verifica se Fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 18626

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 18625
        
        If lErro <> SUCESSO Then Error 18627

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        Call Atualiza_NotasFiscais
    
    End If

    'Não encontrou e é STRING
    If lErro = 6731 Then Error 18628

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case Err
       
       Case 18624, 18625

       Case 18626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
       
       Case 18627
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 18628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160148)

    End Select

    Exit Sub

End Sub

Private Sub FilialNF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o Fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text
    
    'Chama a Tela de com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)
    
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

End Sub

Private Sub NumTitulo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1

    'Coloca a descrição do Fornecedor na tela
    Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Call Fornecedor_Validate(Cancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colNFPagar As New ColNFsPag
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click
    
    'Verifica se o Fornecedor e a Filial foram informados
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 18603
    If Len(Trim(Filial.Text)) = 0 Then Error 18604

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas, colNFPagar)
    If lErro <> SUCESSO Then Error 18605

    'Adiciona o código do Fornecedor e o código da Filial
    colSelecao.Add objTituloPagar.lFornecedor
    colSelecao.Add objTituloPagar.iFilial

    'Chama a Tela com a Lista de Faturas
    Call Chama_Tela("FaturasPagLista", colSelecao, objTituloPagar, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 18603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 18604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 18605

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160149)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloPagar = obj1

    'Traz para a Tela os dados da Fatura passada em objTituloPagar
    lErro = Traz_FaturaPagar_Tela(objTituloPagar)
    If lErro <> SUCESSO Then Error 18604

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 18604

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160150)

    End Select

    Exit Sub

End Sub

Private Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As Collection

    'Verifica se a condicao está preenchida
    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
    End If

    'Chama a Tela com a Lista das Condicões de Pagamento disponíveis
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Coloca na Tela a condicao de Pagamento retornada
    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida

    'Preenchi o Grid de Parcela baseado com os dados presentes na tela
    lErro = GridParcelas_Preenche(objCondicaoPagto)
    If lErro <> SUCESSO Then Error 18605

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case Err

        Case 18605

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160151)

     End Select

     Exit Sub

End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloPagar Is Nothing) Then

        'Lê o Título
        lErro = CF("TituloPagar_Le", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18372 Then Error 18606
        
        If lErro <> SUCESSO Then Error 18384
            
        'Verifica se é Fatura
        If objTituloPagar.sSiglaDocumento <> TIPODOC_FATURA_A_PAGAR Then Error 18607

        'Traz os dados para a tela
        lErro = Traz_FaturaPagar_Tela(objTituloPagar)
        If lErro <> SUCESSO Then Error 18608

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 18606, 18608

        Case 18607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_FATURAPAGAR", Err, objTituloPagar.lNumTitulo)

        Case 18384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURAPAG_NAO_CADASTRADA", Err, objTituloPagar.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160152)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Traz_FaturaPagar_Tela(objTituloPagar As ClassTituloPagar) As Long
'Coloca na Tela os dados da Fatura passada em objTituloPagar

Dim lErro As Long
Dim colParcelasPag As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim colNFPagAberta As New ColNFsPag
Dim colNFPagVinculada As New ColNFsPag
Dim objNFPag As ClassNFsPag
Dim iLinha As Integer
Dim iIndice As Integer, bCancel As Boolean
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Traz_FaturaPagar_Tela

    Call Limpa_Tela_FaturasPag
    
    'Preenche o Cabeçalho
    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(bCancel)

    Filial.Text = objTituloPagar.iFilial
    Call Filial_Validate(bCancel)

    NumTitulo.Text = objTituloPagar.lNumTitulo
    
    Call DateParaMasked(DataEmissao, objTituloPagar.dtDataEmissao)

    ValorTotal.Text = Format(objTituloPagar.dValorTotal, "Standard")
    ValorIRRF.Text = objTituloPagar.dValorIRRF
          
    If objTituloPagar.iCondicaoPagto <> 0 Then
    
        CondicaoPagamento.Text = CStr(objTituloPagar.iCondicaoPagto)
        Call CondicaoPagamento_Validate(bCancel)
    
    Else
    
        CondicaoPagamento.Text = ""
    
    End If
    
    'Lê as Parcelas a Pagar vinculadas ao Título
    lErro = CF("ParcelasPagar_Le", objTituloPagar, colParcelasPag)
    If lErro <> SUCESSO Then Error 18608

    'Verifica se o número de parcelas no BD é superior ao máximo
    If colParcelasPag.Count > NUM_MAXIMO_PARCELAS Then Error 18787
    
    Call Grid_Limpa(objGridParcelas)
    
    iLinha = 0

    'Preenche as linhas do Grid Parcelas com os dados de cada Parcela
    For Each objParcelaPagar In colParcelasPag

        iLinha = iLinha + 1

        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objParcelaPagar.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_VenctoReal_Col) = Format(objParcelaPagar.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objParcelaPagar.dValor, "Standard")

        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = objParcelaPagar.iTipoCobranca Then
                GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next

        If objParcelaPagar.iStatus = STATUS_SUSPENSO Then
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "1"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "0"
        End If
    
    Next

    'Coloca o Numero de Linhas existentes do Grid de Parcelas
    objGridParcelas.iLinhasExistentes = iLinha
    Call Grid_Refresh_Checkbox(objGridParcelas)

    'Atualiza as Notas Fiscais no Grid de Notas Fiscais
    Call Atualiza_NotasFiscais

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objTituloPagar.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 184126

    iAlterado = 0
    iFornecedorAlterado = 0
    iValorTotalAlterado = 0
    iEmissaoAlterada = 0
    iValorIRRFAlterado = 0
    
    Traz_FaturaPagar_Tela = SUCESSO

    Exit Function

Erro_Traz_FaturaPagar_Tela:

    Select Case Err

        Case 18606, 184126

        Case 18787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", Err, colParcelasPag.Count, NUM_MAXIMO_PARCELAS)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160153)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = 1

    Call Fornecedor_Preenche

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 26023

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then Error 26024

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

            'CODIGO ESPECÍFICO
            'Verifica se chave de TipoFornecedor está preenchida
            If objFornecedor.iTipo > 0 Then

                'le os dados do Tipo de fornecedor
                lErro = TipoFornecedor_Dados(objFornecedor, objTipoFornecedor)
                If lErro <> SUCESSO Then Error 26025

            End If
            
            'Se Cond Pagto de Fornecedor está preenchida
            If objFornecedor.iCondicaoPagto <> 0 Then
              
                objCondicaoPagto.iCodigo = objFornecedor.iCondicaoPagto
                
                'Lê a Cond Pagto
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then Error 26026
                If lErro = 19205 Then Error 26028
                
                If objCondicaoPagto.iEmPagamento = 1 Then
                    'Coloca na Tela
                    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
                    Call CondicaoPagamento_Validate(bCancel)
                End If
            'Se Cond Pagto de TipoFornecedor está preenchida
            ElseIf objTipoFornecedor.iCondicaoPagto <> 0 Then
                
                objCondicaoPagto.iCodigo = objTipoFornecedor.iCondicaoPagto
                
                'Lê a Cond Pagto
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then Error 26027
                If lErro = 19205 Then Error 26029
                
                If objCondicaoPagto.iEmPagamento = 1 Then
                    'Coloca na Tela
                    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
                    Call CondicaoPagamento_Validate(bCancel)
                End If
                
            End If

            'Atualiza as Notas Fiscais no Grid de Notas Fiscais
            Call Atualiza_NotasFiscais
            
            'FIM CODIGO ESPECÍFICO
            
        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then
            
            'Limpa Combo de Filial
            Filial.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:
    
    Cancel = True

    Select Case Err

        Case 26023, 26024, 26025, 26026, 26027
        
        Case 26028, 26029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160154)

    End Select

    Exit Sub

End Sub

Private Function Atualiza_NotasFiscais(Optional iForcada As Integer = 0) As Long
'Atualiza no Grid as Notas Fiscais a Pagar de acordo com os dados passados por parâmetro

Dim lErro As Long
Dim lErro1 As Integer
Dim colNfsPagAberta As New ColNFsPag
Dim colNFsPagVinculada As New ColNFsPag
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim sFornecedor As String
Dim sFilial As String
Dim sNumTitulo As String
Dim dtDataEmissao As Date

On Error GoTo Erro_Atualiza_NotasFiscais

    'Recolhe os dados para fazera atualizacao das Notas Fiscais
    sFornecedor = Fornecedor.Text
    sFilial = Filial.Text
    sNumTitulo = NumTitulo.Text
    dtDataEmissao = MaskedParaDate(DataEmissao)

    'Verifica se os dados da tela continuam os mesmos
    If iForcada = 0 And sFornecedor = sOldFornecedor And sFilial = sOldFilial And sNumTitulo = sOldNumTitulo And dtDataEmissao = dtOldDataEmissao Then Exit Function

    'Verifica se todos estão preenchidos
    If Len(Trim(sFornecedor)) = 0 Or Len(Trim(sFilial)) = 0 Or Len(Trim(sNumTitulo)) = 0 Then Exit Function

    objFornecedor.sNomeReduzido = sFornecedor
    
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO Then Error 18648

    'Carrega o fornecedor e filial em objFilialFornecedor
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    objFilialFornecedor.iCodFilial = Codigo_Extrai(sFilial)

    'Guarda os dados do Título
    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = objFilialFornecedor.iCodFilial
    objTituloPagar.lNumTitulo = CLng(sNumTitulo)
    objTituloPagar.dtDataEmissao = dtDataEmissao
    objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR
    objTituloPagar.iFilialEmpresa = giFilialEmpresa

    'Lê o Título
    lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18551 Then Error 18640
    
    'se é uma fatura já cadastrada
    If lErro = SUCESSO Then
    
        'Lê as Notas Fiscais Vinculadas a Fatura a Pagar
        lErro = CF("NfsPag_Le_FaturaPagar", objTituloPagar, colNFsPagVinculada)
        If lErro <> SUCESSO And lErro <> 26020 Then Error 18642
    
    Else
    
        lErro1 = CF("NfsPag_Le_FilialFornecedor_Desvinculadas", objFilialFornecedor, colNfsPagAberta)
        If lErro1 <> SUCESSO Then Error 18639

    End If
    
    'Limpa o Grid
    Call Grid_Limpa(objGridNFiscais)
    ValorTotalNFSelecionadas.Caption = ""
    NumNFSelecionadas.Caption = ""
        
    'joga p/o grid as nfs desvinculadas
    Call Carrega_GridNF(colNfsPagAberta, colNFsPagVinculada)

    'Atualiza os dados
    sOldFornecedor = objFornecedor.sNomeReduzido
    sOldFilial = sFilial
    sOldNumTitulo = sNumTitulo
    dtOldDataEmissao = dtDataEmissao

    Atualiza_NotasFiscais = SUCESSO

    Exit Function

Erro_Atualiza_NotasFiscais:

    Atualiza_NotasFiscais = Err

    Select Case Err

        Case 18639, 18640, 18642, 18648

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160155)

    End Select

    Exit Function

End Function

Private Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    'Verifica se alguma Filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub

    'Atualiza as Notas Fiscais no Grid de Notas Fiscais
    Call Atualiza_NotasFiscais
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_Filial_Click:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160156)
        
    End Select
    
    Exit Sub

End Sub

Private Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumTitulo.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumTitulo.ClipText)
    If lErro <> SUCESSO Then Error 18647

    'Atualiza as Notas Fiscais no Grid de Notas Fiscais
    Call Atualiza_NotasFiscais

    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True


    Select Case Err

        Case 18647

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160157)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_FATURAS_PAGAR_ID
                
            Case TAB_Cobranca
                Parent.HelpContextID = IDH_FATURAS_PAGAR_COBRANCA
                        
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_CONTABILIZACAO
        
        End Select
        
    End If

End Sub

Private Sub UpDownEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iEmissaoAlterada = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorTotalAlterado = 1
    
End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(ValorTotal.Text)
    If lErro <> SUCESSO Then Error 18651

    'Põe o valor formatado na tela
    ValorTotal.Text = Format(ValorTotal.Text, "Standard")

    If iValorTotalAlterado = 1 Then
    
        Call Recalcula_Cobranca
        iValorTotalAlterado = 0
        
    End If
    
    Exit Sub

Erro_ValorTotal_Validate:

    Cancel = True


    Select Case Err

        Case 18651

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160158)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 18652

    If iEmissaoAlterada = 1 Then
        
        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        Call Atualiza_NotasFiscais
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case Err

        Case 18652

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160159)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 18654

    If iEmissaoAlterada = 1 Then
        
        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        Call Atualiza_NotasFiscais
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 18654

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160160)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 18655

    If iEmissaoAlterada = 1 Then
        
        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        Call Atualiza_NotasFiscais
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 18655

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160161)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagamento_Click()

Dim lErro As Long

On Error GoTo Erro_CondicaoPagamento_Click

    iAlterado = REGISTRO_ALTERADO

    Call Recalcula_Cobranca
    
    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160162)

      End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValor As Double
Dim dtDataEmissao As Date
Dim dtDataVenctoReal As Date
Dim iIndice As Integer, dValorIRRF As Double

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o Grid Parcelas
    Call Grid_Limpa(objGridParcelas)

    dValorIRRF = StrParaDbl(ValorIRRF)
    
    'Valor a Pagar
    If Len(Trim(ValorTotal)) > 0 Then dValor = CDbl(ValorTotal) - dValorIRRF

    'Se Valor a Pagar for positivo
    If dValor > 0 Then

        objCondicaoPagto.dValorTotal = dValor
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then Error 18417

        'Número de Parcelas
        objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        dtDataEmissao = CDate(DataEmissao.Text)
        objCondicaoPagto.dtDataEmissao = dtDataEmissao
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then Error 18441

        objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas
        
        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
            
            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then Error 18443

            'Coloca Data de Vencimento Real no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")
            
            lErro = CF("Fatura_Preenche_Cobranca_Cust", objGridParcelas, TipoCobranca, iIndice, iGrid_Cobranca_Col, Fornecedor.Text)
            If lErro <> SUCESSO Then Error 18443

        Next

    End If

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = Err

    Select Case Err

        Case 18417, 18441, 18443

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160163)

    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 18613

        If objGridInt.objGrid Is GridParcelas Then

            'Verifica qual e coluna atual
            Select Case objGridInt.objGrid.Col

                'faz a Critica da Data de Vencimento e Gera a Data de Vencimento Real
                Case iGrid_Vencimento_Col
                    lErro = Saida_Celula_Vencimento(objGridInt)
                    If lErro <> SUCESSO Then Error 18613

                'faz o tratamento relacionado ao valor da Parcela
                Case iGrid_ValorParcela_Col
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then Error 18614

                'faz o tratamento relacionado ao Tipo de Cobrança
                Case iGrid_Cobranca_Col
                    lErro = Saida_Celula_Cobranca(objGridInt)
                    If lErro <> SUCESSO Then Error 18615

            End Select
       
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 18617

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 18613, 18614, 18615, 18616

        Case 18617
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a critica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = ValorParcela

    'Verifica se o valor parcela foi digitda
    If Len(ValorParcela.ClipText) > 0 Then
        
        'Verifica se o Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then Error 18666
        
        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then CobrancaAutomatica.Value = vbUnchecked
        
        'No caso de ser uma linha nova incrementa-se o número de linhas existente
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            For iIndice = 0 To TipoCobranca.ListCount - 1
                If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                    Exit For
                End If
            Next
        
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 18667

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 18666, 18667
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160164)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vencimento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim dtDataVenctoReal As Date
Dim sDataVencimento As String

On Error GoTo Erro_Saida_Celula_Vencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 18613

        dtDataVencimento = CDate(DataVencimento.Text)

        'Se Data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            
            dtDataEmissao = CDate(DataEmissao.Text)
            If dtDataVencimento < DataEmissao Then Error 18614
        
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
        
        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then Error 18615

        'Coloca a  Data de Vencimento Real na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        'Incrementa o número de linhas existentes caso esteja sendo digitado em uma lina nao existente
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            For iIndice = 0 To TipoCobranca.ListCount - 1
                If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                    Exit For
                End If
            Next
        
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 18616

    Saida_Celula_Vencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_Vencimento:

    Saida_Celula_Vencimento = Err

    Select Case Err

        Case 18613, 18615, 18616
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, DataVencimento.Text, GridParcelas.Row, DataEmissao.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160165)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobranca(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Cobranca

    Set objGridInt.objControle = TipoCobranca

    'Verifica se o Tipo de Cobrança foi preenchido
    If Len(Trim(TipoCobranca.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoCobranca.Text <> TipoCobranca.List(TipoCobranca.ListIndex) Then

            'Tenta Selecionar na combo o Tipo de Cobrança
            lErro = Combo_Seleciona(TipoCobranca, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18665

            'Não foi encontrado, mas retornou código
            If lErro = 6730 Then Error 18788
            
            'Não foi encontrado a STRING
            If lErro = 6731 Then Error 18789

        End If
    
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 18668

    Saida_Celula_Cobranca = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobranca:

    Saida_Celula_Cobranca = Err

    Select Case Err

        Case 18665
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", Err, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18668
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18788
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_CADASTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 18789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", Err, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160166)

    End Select

    Exit Function

End Function

Private Sub Selecionada_Click()

Dim dValorTotalNFSel As Double
Dim dValorNF As Double
Dim iNumSelecionadas As Integer
Dim iSelecionada As Integer

    'Verifica se o click foi dado em uma linha não preenchida do Grid
    If GridNF.Row > objGridNFiscais.iLinhasExistentes Then Exit Sub
    
    'Recolhe os valores da tela
    iSelecionada = CInt(GridNF.TextMatrix(GridNF.Row, iGrid_Selecionada_Col))
    dValorNF = CDbl(GridNF.TextMatrix(GridNF.Row, iGrid_ValorNF_Col))
    
    'Recolhe os valores da tela
    iNumSelecionadas = CInt(NumNFSelecionadas.Caption)
    dValorTotalNFSel = CDbl(ValorTotalNFSelecionadas.Caption)
    
    'atualiza os valores dos acumuladores
    dValorTotalNFSel = dValorTotalNFSel + IIf(iSelecionada = 1, dValorNF, -dValorNF)
    iNumSelecionadas = iNumSelecionadas + IIf(iSelecionada = 1, 1, -1)

    'Atualiza os totais na Tela
    ValorTotalNFSelecionadas.Caption = Format(dValorTotalNFSel, "Standard")
    NumNFSelecionadas.Caption = CStr(iNumSelecionadas)

    Exit Sub

End Sub

Private Sub NFRegistrar_Click()
    
    'Chama a tela de Notas Fiscais
    Call Chama_Tela("NFPag")
    
End Sub

Private Sub BotaoGravar_Click()
'Dispara a gravação de uma Fatura

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 18669

    'Limpa os campos da tela
    Call Limpa_Tela_FaturasPag

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 18669

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160167)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValorParcelas As Double
Dim iIndice As Integer
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelaPagar As New colParcelaPagar
Dim colNFPag As New ColNFsPag
Dim dtDataVencimento As Date, dtDataEmissao As Date
Dim dValorTotal As Double
Dim dValorIRRF As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 18670

    'Verifica se a Filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 18671

    'Verifica se o Número do Título foi preenchido
    If Len(Trim(NumTitulo.ClipText)) = 0 Then gError 18673

    'Verifica se o Valor Total foi preenchido
    If Len(Trim(ValorTotal.ClipText)) = 0 Then gError 18684
    
    'Verifica se exitem parcelas no Grid para gravar
    If objGridParcelas.iLinhasExistentes = 0 Then gError 18686
    
    dtDataEmissao = MaskedParaDate(DataEmissao)
    dValorParcelas = 0
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        'Verifica se a Data de Vencimento foi informada
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 18687
    
        dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        
        'Verifica se a Data de Vencimento é maior ou igual à Data de Emissao
        If dtDataEmissao <> DATA_NULA And dtDataVencimento < dtDataEmissao Then gError 18688
        
        'Verifica se as parcelas estão ordenadas pela Data de Vencimento
        If iIndice > 1 Then
            If dtDataVencimento < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 18689
        End If
        
        'Verifica se o Valor das Parcelas foi informado
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 18690
             
        'Verifica se Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        If lErro <> SUCESSO Then gError 18691

        'Acumula Valor Parcela em dSomaParcelas
        dValorParcelas = dValorParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        
    Next
    
    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
    
    dValorTotal = CDbl(ValorTotal.Text)
    
    'Verifica se ValorTotal = soma das Parcelas
    If Format(dValorTotal - dValorIRRF, "0.00") <> Format(dValorParcelas, "0.00") Then gError 18692
    
    'Verifica se Valor Total = soma Notas Fiscais selecionadas
    If StrParaDbl(ValorTotal.Text) <> StrParaDbl(ValorTotalNFSelecionadas.Caption) Then gError 18693
    
    '??? Jones 09/05/01    If StrParaInt(NumNFSelecionadas.Caption) > NUM_MAXIMO_NF_VINCULADA_FATURA Then gError 19392
    
    'Move dados da Tela para objTituloPagar, colParcelaPagar e colNFPag
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelaPagar, colNFPag)
    If lErro <> SUCESSO Then gError 18694
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    If StrParaDate(DataEmissao.Text) <> DATA_NULA Then
        lErro = objContabil.Contabil_Testa_Data(StrParaDate(DataEmissao.Text))
        If lErro <> SUCESSO Then gError 182905
    End If
    
    'Grava no BD
    lErro = CF("FaturaPagar_Grava", objTituloPagar, colParcelaPagar, colNFPag, objContabil)
    If lErro <> SUCESSO Then gError 18695
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 18670
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 18671
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
                
        Case 18673
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
        
        Case 18684
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", gErr)
        
        Case 18686
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", gErr)

        Case 18687
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", gErr, iIndice)
            
        Case 18688
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", gErr, dtDataVencimento, DataEmissao.Text, iIndice)
        
        Case 18689
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_NAO_ORDENADA", gErr)
        
        Case 18690
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_NAO_INFORMADO", gErr, iIndice)
            
        Case 18691, 18694, 18695, 182905
        
        Case 18692
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_INVALIDA", gErr, dValorParcelas, dValorTotal)
        
        Case 18693
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_NFS_SELECIONADAS_INVALIDA", gErr, ValorTotalNFSelecionadas.Caption, ValorTotal.Text)
    
        Case 19392
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_MAX_NFS_SELEC_EXCEDIDO", gErr, NUM_MAXIMO_NF_VINCULADA_FATURA)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160168)
            
    End Select
        
    Exit Function
        
End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelaPagar As New colParcelaPagar
Dim colNFPagar As New ColNFsPag
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 18752
    
    'Verifica se Filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Error 18753
    
    'Verifica se o número do Título foi preenchido
    If Len(Trim(NumTitulo.Text)) = 0 Then Error 18754

    'Guarda os dados da tela em objTituloPagar, colParcelaPagar, colNFPagar
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelaPagar, colNFPagar)
    If lErro <> SUCESSO Then Error 18755
    
    'Tenta ler Fatura como Baixada
    lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then Error 18758
    
    'Se encontrou --> Erro
    If lErro = SUCESSO Then Error 18759
    
    'Verifica se a Fatura está cadastrada
    lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18551 Then Error 18760
        
    'Se não encontrou --> Erro
    If lErro <> SUCESSO Then Error 18761

    'Pede a confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FATURAPAGAR", objTituloPagar.lNumTitulo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui a Fatura do BD (incluindo dados contabeis (contabilidade))
    lErro = CF("FaturaPagar_Exclui", objTituloPagar, objContabil)
    If lErro <> SUCESSO Then Error 18762
    
    'Limpa a tela
    Call Limpa_Tela_FaturasPag
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 18752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
            
        Case 18753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
        
        Case 18754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)
    
        Case 18755, 18758, 18760, 18762
        
        Case 18759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURAPAG_BAIXADA_EXCLUSAO", Err, objTituloPagar.lNumTitulo)
            
        Case 18761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURAPAG_NAO_CADASTRADA2", Err, objTituloPagar.lNumTitulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160169)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 18751
    
    'Lima a Tela
    Call Limpa_Tela_FaturasPag
       
    Exit Sub

Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 18751
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160170)
    
    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iEmissaoAlterada = 1
    
End Sub

Private Function Carrega_TipoCobranca() As Long
'Carrega na combobox os Tipos de Cobrança

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o nome e o codigo de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 18598

    'Carrega na combo de Tipos de Cobrança os Tipos que estão em colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        TipoCobranca.AddItem CInt(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = Err

    Select Case Err

        Case 18598

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160171)

    End Select

    Exit Function

End Function

Private Function Carrega_CondicaoPagamento() As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê as Condicões de Pagamento utilizadas em Contas a Pagar
    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
    If lErro <> SUCESSO Then Error 18599

    'Carrega na Combo as Condições de Pagamento retornadas em colCod_DescReduzida
    For Each objCodDescricao In colCod_DescReduzida

        CondicaoPagamento.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = Err

    Select Case Err

        Case 18599

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160172)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridParcelas(objGridInt As AdmGrid) As Long
'Faz as Inicializações no Grid de Parcelas a Pagar

    'Indica o Form do Grid
    Set objGridInt.objForm = Me

    'Indica os nomes das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Suspenso")

    'Indica os campos relacionados a cada coluna
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoCobranca.Name)
    objGridInt.colCampo.Add (Suspenso.Name)

    'Inicializa os valores das colunas
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_Col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_ValorParcela_Col = 3
    iGrid_Cobranca_Col = 4
    iGrid_Suspenso_Col = 5

    'Indica o Grid à que se referem os dados
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do Grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do Grid
    objGridInt.iLinhasVisiveis = 8

    'Determina a largura da coluna 0
    GridParcelas.ColWidth(0) = 900

    'Indica a largura do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama a rotina que faz as demais inicializações
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridParcelas = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridNF(objGridInt As AdmGrid) As Long
'Faz as Inicializações no Grid de Notas Fiscais Fatura

    'Indica o Form do Grid
    Set objGridInt.objForm = Me

    'Indica os nomes das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Filial ")
    objGridInt.colColuna.Add ("Filial Forn")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Selecionada")
    
    'Indica os campos relacionados a cada coluna
    objGridInt.colCampo.Add (FilialNF.Name)
    objGridInt.colCampo.Add (FilialFornecedor.Name)
    objGridInt.colCampo.Add (NumNotaFiscal.Name)
    objGridInt.colCampo.Add (DataEmissaoNF.Name)
    objGridInt.colCampo.Add (ValorNF.Name)
    objGridInt.colCampo.Add (Selecionada.Name)
    
    
    'Inicializa os valores das colunas
    iGrid_FilialNF_Col = 1
    iGrid_FilialFornecedor_Col = 2
    iGrid_Numero_Col = 3
    iGrid_Emissao_Col = 4
    iGrid_ValorNF_Col = 5
    iGrid_Selecionada_Col = 6

    'Indica o Grid ao qual faz referencia
    objGridInt.objGrid = GridNF

    'Linhas visíveis do Grid
    objGridInt.iLinhasVisiveis = NUM_LINHAS_GRID_NF

    'Todas as linhas do Grid
    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1

    'Indica a largura da coluna 0
    GridNF.ColWidth(0) = 300

    'Indica a largura automática do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR

    'Chama a rotina que faz as demais inicializações
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridNF = SUCESSO

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoCondPagto = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
    Set objGridParcelas = Nothing
    Set objGridNFiscais = Nothing
    Set objContabil = Nothing

End Sub

Private Sub GridNF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNFiscais, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscais, iAlterado)
    End If

End Sub

Private Sub GridNF_EnterCell()
    Call Grid_Entrada_Celula(objGridNFiscais, iAlterado)
End Sub

Private Sub GridNF_GotFocus()
    Call Grid_Recebe_Foco(objGridNFiscais)
End Sub

Private Sub GridNF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNFiscais, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscais, iAlterado)
    End If

End Sub

Private Sub GridNF_LeaveCell()
        Call Saida_Celula(objGridNFiscais)
End Sub

Private Sub GridNF_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridNFiscais)
End Sub

Private Sub GridNF_RowColChange()
    Call Grid_RowColChange(objGridNFiscais)
End Sub

Private Sub GridNF_Scroll()
    Call Grid_Scroll(objGridNFiscais)
End Sub

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_EnterCell()
    
    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    
End Sub

Private Sub GridParcelas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridParcelas)
    
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_LeaveCell()
    Call Saida_Celula(objGridParcelas)
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridParcelas)
End Sub

Private Sub GridParcelas_RowColChange()
    Call Grid_RowColChange(objGridParcelas)
End Sub

Private Sub GridParcelas_Scroll()
    Call Grid_Scroll(objGridParcelas)
End Sub

Private Sub NumTitulo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Limpa_Tela_FaturasPag()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa os campos comuns nas telas
    Call Limpa_Tela(Me)

    'Limpa os Grids
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridNFiscais)

    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    'Limpa os campos não limpos em Limpa_Tela
    Filial.Clear
    CondicaoPagamento.Text = ""
    ValorTotalNFSelecionadas.Caption = ""
    NumNFSelecionadas.Caption = ""
        
    sOldFornecedor = ""
    
    iAlterado = 0
    
End Sub

Private Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataVencimento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DataVencimentoReal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataVencimentoReal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Private Sub DataVencimentoReal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Private Sub DataVencimentoReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimentoReal
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TipoCobranca_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoCobranca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoCobranca_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Private Sub TipoCobranca_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Private Sub TipoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoCobranca
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorParcela_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorParcela_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Suspenso_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Suspenso_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Private Sub Suspenso_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Private Sub Suspenso_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Suspenso
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Carrega_GridNF(colNFPag As ColNFsPag, colNFsPagVinculada As ColNFsPag)
'Coloca no Grid de Notas Fiscais as Notas passadas em colNFPag

Dim iLinha As Integer
Dim iIndice As Integer
Dim dValor As Double
Dim iNumNfsSel As Integer
Dim dValorNFsSel As Double
  
    'Recoloca o Numero de Linhas caso seja maior que o Numero de Linhas Visiveis na Tela que é 6
    If (colNFPag.Count + colNFsPagVinculada.Count + 1) > NUM_LINHAS_GRID_NF Then
        GridNF.Rows = colNFPag.Count + colNFsPagVinculada.Count + 1
    Else
        GridNF.Rows = NUM_LINHAS_GRID_NF + 1
    End If
       
    'Renicializa
    Call Grid_Inicializa(objGridNFiscais)
    
    iIndice = 0

    'Prenche os GriNF com as Nfs de colNFsPagVinculada
    For iLinha = 1 To colNFsPagVinculada.Count

        iIndice = iIndice + 1

        GridNF.TextMatrix(iLinha, iGrid_FilialNF_Col) = colNFsPagVinculada(iIndice).iFilialEmpresa
        GridNF.TextMatrix(iLinha, iGrid_Numero_Col) = colNFsPagVinculada(iIndice).lNumNotaFiscal
        GridNF.TextMatrix(iLinha, iGrid_Emissao_Col) = IIf(colNFsPagVinculada(iIndice).dtDataEmissao <> DATA_NULA, Format(colNFsPagVinculada(iIndice).dtDataEmissao, "dd/mm/yyyy"), "")
        GridNF.TextMatrix(iLinha, iGrid_FilialFornecedor_Col) = colNFsPagVinculada(iIndice).iFilial
        
        'Calcula o Valor da Nota Fiscal
        dValor = colNFsPagVinculada(iIndice).dValorTotal - colNFsPagVinculada(iIndice).dValorIRRF
        GridNF.TextMatrix(iLinha, iGrid_ValorNF_Col) = Format(dValor, "Standard")
        If colNFsPagVinculada(iIndice).lNumIntTitPag <> 0 Then
            GridNF.TextMatrix(iLinha, iGrid_Selecionada_Col) = "1"
            iNumNfsSel = iNumNfsSel + 1
            dValorNFsSel = dValorNFsSel + dValor
        End If

    Next

    iIndice = 0

    'Prenche os GriNF com as Nfs de colNFPag
    For iLinha = (colNFsPagVinculada.Count + 1) To colNFPag.Count + colNFsPagVinculada.Count

        iIndice = iIndice + 1

        GridNF.TextMatrix(iLinha, iGrid_FilialNF_Col) = colNFPag(iIndice).iFilialEmpresa
        GridNF.TextMatrix(iLinha, iGrid_Numero_Col) = colNFPag(iIndice).lNumNotaFiscal
        GridNF.TextMatrix(iLinha, iGrid_Emissao_Col) = IIf(colNFPag(iIndice).dtDataEmissao <> DATA_NULA, Format(colNFPag(iIndice).dtDataEmissao, "dd/mm/yyyy"), "")
        GridNF.TextMatrix(iLinha, iGrid_FilialFornecedor_Col) = colNFPag(iIndice).iFilial

        'Calcula o Valor da Nota Fiscal
        dValor = colNFPag(iIndice).dValorTotal - colNFPag(iIndice).dValorIRRF
        GridNF.TextMatrix(iLinha, iGrid_ValorNF_Col) = Format(dValor, "Standard")
        If colNFPag(iIndice).lNumIntTitPag <> 0 Then
            GridNF.TextMatrix(iLinha, iGrid_Selecionada_Col) = "1"
            iNumNfsSel = iNumNfsSel + 1
            dValorNFsSel = dValorNFsSel + dValor
        End If

    Next
   
    objGridNFiscais.iLinhasExistentes = colNFPag.Count + colNFsPagVinculada.Count

    'Inicializa os contadores e acumuladores dos valores
    'das Notas Fiscais
    NumNFSelecionadas.Caption = CStr(iNumNfsSel)
    ValorTotalNFSelecionadas.Caption = Format(dValorNFsSel, "Standard")
    
    'Faz um Refresh nas CheckBox's do Grid de Notas Fiscais
    Call Grid_Refresh_Checkbox(objGridNFiscais)

 End Sub

Private Sub NumNotaFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumNotaFiscal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Private Sub NumNotaFiscal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Private Sub NumNotaFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DataEmissaoNF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEmissaoNF_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Private Sub DataEmissaoNF_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Private Sub DataEmissaoNF_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridParcelas.objControle = DataEmissaoNF
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Selecionada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Private Sub Selecionada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Private Sub Selecionada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscais.objControle = Selecionada
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorNF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorNF_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Private Sub ValorNF_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Private Sub ValorNF_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorNF
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Function Move_Tela_Memoria(objTituloPagar As ClassTituloPagar, colParcelaPagar As colParcelaPagar, colNFPag As ColNFsPag) As Long
'Recolhe os dados da Tela e passa para objTituloPagar, colParcelaPagar, colNFPag

Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer, sData As String
Dim lErro As Long
Dim lNumNotaFiscal As Long
Dim dtDataEmissao As Date

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 18747
        
        If lErro <> SUCESSO Then Error 26019

        objTituloPagar.lFornecedor = objFornecedor.lCodigo
    Else
        objTituloPagar.lFornecedor = 0
    End If
    
    'Verifica se a Filial foi preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    Else
        objTituloPagar.iFilial = 0
    End If
    
    'Se estiver preenchido preenche com o NumTitulo da tela
    If Len(Trim(NumTitulo.ClipText)) > 0 Then objTituloPagar.lNumTitulo = CLng(NumTitulo.ClipText)
    
    'Preenche objTituloPagar com Data Emissão. Se não foi digitada preenche com DATA_NULA.
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    If Len(Trim(ValorTotal.ClipText)) > 0 Then objTituloPagar.dValorTotal = CDbl(ValorTotal.Text)
    If Len(Trim(ValorIRRF.ClipText)) > 0 Then objTituloPagar.dValorIRRF = CDbl(ValorIRRF.Text)
    
    'Carrega os dados restantes da tela
    objTituloPagar.iNumParcelas = objGridParcelas.iLinhasExistentes
    objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR
    objTituloPagar.iFilialEmpresa = giFilialEmpresa
    objTituloPagar.iStatus = STATUS_LANCADO
    objTituloPagar.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    
    'Move os dados do Grid para colParcelaPagar
    lErro = Move_GridParcelas_Memoria(colParcelaPagar)
    If lErro <> SUCESSO Then Error 18748
       
    'Loop de adição das linhas selecionadas de GridNF à colNFPag
    For iIndice = 1 To objGridNFiscais.iLinhasExistentes
    
        'Se Selecionada, adiciona Nota Fiscal a colNFPag
        If GridNF.TextMatrix(iIndice, iGrid_Selecionada_Col) = "1" Then
            lNumNotaFiscal = CLng(GridNF.TextMatrix(iIndice, iGrid_Numero_Col))
            sData = GridNF.TextMatrix(iIndice, iGrid_Emissao_Col)
            If Len(Trim(sData)) > 0 Then
                dtDataEmissao = CDate(sData)
            Else
                dtDataEmissao = DATA_NULA
            End If
            
            colNFPag.Add 0, StrParaInt(GridNF.TextMatrix(iIndice, iGrid_FilialNF_Col)), objTituloPagar.lFornecedor, StrParaInt(GridNF.TextMatrix(iIndice, iGrid_FilialFornecedor_Col)), lNumNotaFiscal, dtDataEmissao, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        End If
    Next

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err
    
        Case 18747, 18748
        
        Case 26019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160173)
        
    End Select
        
    Exit Function

End Function

Private Function Move_GridParcelas_Memoria(colParcelas As colParcelaPagar)
'Move para colParcelas os dados existentes no GridParcelas

Dim iIndice As Integer
Dim objParcelaPag As ClassParcelaPagar
 
    'Loop de preenchimento de colParcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        Set objParcelaPag = New ClassParcelaPagar
        
        'Preenchimento de objParcelaPag com linha do GridParcelas
        objParcelaPag.iNumParcela = iIndice
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) > 0 Then objParcelaPag.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))) > 0 Then objParcelaPag.dtDataVencimentoReal = CDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) > 0 Then objParcelaPag.dValor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))) = 0 Then
            objParcelaPag.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
        Else
            objParcelaPag.iTipoCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
        End If
                
        If GridParcelas.TextMatrix(iIndice, iGrid_Suspenso_Col) = "1" Then
            objParcelaPag.iStatus = STATUS_SUSPENSO
        Else
            objParcelaPag.iStatus = STATUS_ABERTO
        End If
        
        'Adição de objParcelaPag a colParcelas
        With objParcelaPag
            Call colParcelas.Add(0, 0, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, 0, .dValor, 0, 1, .iTipoCobranca, 0, "", "")
        End With
    Next
        
    Move_GridParcelas_Memoria = SUCESSO
    
    Exit Function
                
End Function

Private Function TipoFornecedor_Dados(objFornecedor As ClassFornecedor, objTipoFornecedor As ClassTipoFornecedor) As Long
'Lê os dados de TipoFornecedor ligado a objFornecedor.
    
Dim lErro As Long

On Error GoTo Erro_TipoFornecedor_Dados

    objTipoFornecedor.iCodigo = objFornecedor.iTipo
    
    'Lê o TipoFornecedor a partir do código
    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then Error 26031
    If lErro <> SUCESSO Then Error 26030
   
    TipoFornecedor_Dados = SUCESSO

    Exit Function

Erro_TipoFornecedor_Dados:

    TipoFornecedor_Dados = Err

    Select Case Err

        Case 26030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO", Err, objTipoFornecedor.iCodigo)
        
        Case 26031

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160174)

    End Select

    Exit Function

End Function

Private Sub Recalcula_Cobranca()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Recalcula_Cobranca

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
    
        'Passa o código da Condição para objCondicaoPagto
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
    
        'Lê Condição a partir do código
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 18657
        If lErro = 19205 Then Error 18659
    
        'Preenche o GridParcelas
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then Error 18658

    End If
    
    Exit Sub
     
Erro_Recalcula_Cobranca:

    Select Case Err
          
        Case 18657, 18658

        Case 18659
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160175)
     
    End Select
     
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FATURAS_PAGAR_ID
    Set Form_Load_Ocx = Me
    Caption = "Faturas a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FaturasPag"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is CondicaoPagamento Then
            Call CondPagtoLabel_Click
        End If
    
    End If
    
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
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134046

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134046

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160176)

    End Select
    
    Exit Sub

End Sub

Public Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIRRF_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIRRF.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
        If lErro <> SUCESSO Then Error 18405
    
        'Põe o valor formatado na tela
        ValorIRRF.Text = Format(ValorIRRF.Text, "Fixed")

    End If
    
    If iValorIRRFAlterado <> 0 Then
    
        Call Recalcula_Cobranca
        iValorIRRFAlterado = 0
        
    End If
    
    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True

    Select Case Err

        Case 18405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156142)

    End Select

    Exit Sub

End Sub

'inicio contabilidade

Public Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Public Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Public Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Public Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub


Public Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Public Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Public Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Public Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Public Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Public Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Public Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Public Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Public Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Public Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Public Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Public Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Public Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Public Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Public Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Public Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Public Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Public Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Public Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Public Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Public Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Public Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Public Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Public Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Public Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Public Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Public Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Public Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Public Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Public Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Public Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Public Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Public Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Public Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Public Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Public Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Public Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Public Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Public Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Public Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor, objTipoFornecedor As New ClassTipoFornecedor
Dim objFilial As New ClassFilialFornecedor, sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case VALORTOTAL1
            
            If Len(Trim(ValorTotal.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case FORNECEDOR_COD
            
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 39567
                
                objMnemonicoValor.colValor.Add objFornecedor.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case FORNECEDOR_NOME
        
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 39609
            
                objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case FILIAL_COD
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                objMnemonicoValor.colValor.Add objFilial.iCodFilial
            
            Else
                
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Case FILIAL_NOME_RED
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39568
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39569
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 41600
                    
                    objMnemonicoValor.colValor.Add sContaTela
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39570
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case NUMERO1
            
            If Len(Trim(NumTitulo.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NumTitulo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case EMISSAO1
            If Len(Trim(DataEmissao.ClipText)) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case VALOR_IR
            If Len(Trim(ValorIRRF.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIRRF.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CONTA_DESP_ESTOQUE
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 39567
                
                objTipoFornecedor.iCodigo = objFornecedor.iTipo
                lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
                If lErro <> SUCESSO Then gError 41599
                
                If objTipoFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTipoFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 41968
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case CONTA_DESP_EST_FORN
        
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 133564
                
                If objFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 133565
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case Else
            gError 36229
            
    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 36229
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 39567, 39568, 39569, 39570, 39609, 41599, 41968, 133564, 133565
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156160)

    End Select

    Exit Function

End Function

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


