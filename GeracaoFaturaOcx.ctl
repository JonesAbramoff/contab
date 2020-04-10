VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl GeracaoFaturaOcx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4665
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   855
      Width           =   9075
      Begin VB.Frame Frame4 
         Caption         =   "Cliente da Fatura"
         Height          =   900
         Left            =   900
         TabIndex        =   123
         Top             =   150
         Width           =   7260
         Begin VB.ComboBox FilialCliente 
            Height          =   315
            Left            =   4575
            TabIndex        =   1
            Top             =   390
            Width           =   1848
         End
         Begin MSMask.MaskEdBox CodCliente 
            Height          =   300
            Left            =   1635
            TabIndex        =   0
            Top             =   375
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LblFilialCli 
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
            Height          =   255
            Left            =   4050
            TabIndex        =   125
            Top             =   435
            Width           =   495
         End
         Begin VB.Label LblCliente 
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
            Left            =   885
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   124
            Top             =   435
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seleciona as Notas Fiscais"
         Height          =   3300
         Left            =   885
         TabIndex        =   68
         Top             =   1215
         Width           =   7275
         Begin VB.CheckBox optTodasFiliais 
            Caption         =   "Trazer dados de todas as filiais"
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
            Left            =   1650
            TabIndex        =   2
            Top             =   375
            Width           =   3285
         End
         Begin VB.Frame Frame6 
            Caption         =   "Número"
            Height          =   1350
            Left            =   885
            TabIndex        =   75
            Top             =   675
            Width           =   5520
            Begin VB.ComboBox Serie 
               Height          =   315
               Left            =   795
               TabIndex        =   3
               Top             =   360
               Width           =   765
            End
            Begin MSMask.MaskEdBox NFiscalInicial 
               Height          =   300
               Left            =   795
               TabIndex        =   4
               Top             =   840
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NFiscalFinal 
               Height          =   300
               Left            =   3450
               TabIndex        =   5
               Top             =   840
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   3030
               TabIndex        =   78
               Top             =   885
               Width           =   360
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               Left            =   420
               TabIndex        =   79
               Top             =   885
               Width           =   315
            End
            Begin VB.Label LabelSerie 
               AutoSize        =   -1  'True
               Caption         =   "Série:"
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
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   80
               Top             =   390
               Width           =   510
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data Emissão"
            Height          =   945
            Left            =   885
            TabIndex        =   73
            Top             =   2130
            Width           =   5505
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   780
               TabIndex        =   6
               Top             =   390
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1935
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   390
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   3450
               TabIndex        =   8
               Top             =   390
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   4605
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   390
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               Left            =   420
               TabIndex        =   81
               Top             =   443
               Width           =   315
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   3030
               TabIndex        =   82
               Top             =   450
               Width           =   360
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4665
      Index           =   2
      Left            =   180
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox GeraFatura 
         Height          =   225
         Left            =   450
         TabIndex        =   16
         Top             =   3270
         Width           =   1050
      End
      Begin VB.TextBox Serie1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Text            =   "Serie"
         Top             =   3120
         Width           =   480
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         TabIndex        =   19
         Text            =   "Cliente"
         Top             =   3465
         Width           =   1305
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   20
         Text            =   "Filial"
         Top             =   3450
         Width           =   705
      End
      Begin VB.TextBox DataEmissao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   21
         Text            =   "Emissão"
         Top             =   3300
         Width           =   1185
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   885
         Picture         =   "GeracaoFaturaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3900
         Width           =   1800
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2850
         Picture         =   "GeracaoFaturaOcx.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3900
         Width           =   1800
      End
      Begin VB.CommandButton BotaoNFiscal 
         Caption         =   "Ver Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   6030
         Picture         =   "GeracaoFaturaOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3900
         Width           =   1800
      End
      Begin VB.TextBox SiglaDoc 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6660
         TabIndex        =   22
         Text            =   "Tipo"
         Top             =   3285
         Width           =   1095
      End
      Begin VB.TextBox ValorTotal 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5265
         TabIndex        =   23
         Text            =   "Valor Total"
         Top             =   3465
         Width           =   1035
      End
      Begin VB.TextBox Numero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1950
         TabIndex        =   18
         Text            =   "Numero"
         Top             =   3120
         Width           =   840
      End
      Begin MSFlexGridLib.MSFlexGrid GridNFiscal 
         Height          =   2805
         Left            =   225
         TabIndex        =   24
         Top             =   345
         Width           =   7770
         _ExtentX        =   13705
         _ExtentY        =   4948
         _Version        =   393216
         Rows            =   11
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label LabelTotalNotaSel 
         AutoSize        =   -1  'True
         Caption         =   "Total das Notas Selecionadas:"
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
         Left            =   2520
         TabIndex        =   83
         Top             =   3240
         Width           =   2625
      End
      Begin VB.Label TotalNotasSel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6465
         TabIndex        =   84
         Top             =   3210
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4665
      Index           =   5
      Left            =   180
      TabIndex        =   47
      Top             =   840
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4560
         TabIndex        =   122
         Tag             =   "1"
         Top             =   2520
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
         Left            =   6330
         TabIndex        =   53
         Top             =   345
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
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   55
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
         Height          =   270
         Left            =   7770
         TabIndex        =   52
         Top             =   30
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4920
         TabIndex        =   60
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   62
         Top             =   2040
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   61
         Top             =   1650
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6420
         TabIndex        =   65
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   74
         Top             =   3510
         Width           =   5895
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
            Height          =   195
            Left            =   240
            TabIndex        =   93
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
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
            Height          =   195
            Left            =   1125
            TabIndex        =   94
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   95
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   96
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
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
         Left            =   3510
         TabIndex        =   56
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   63
         Top             =   1335
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
         TabIndex        =   59
         Top             =   1365
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
         TabIndex        =   58
         Top             =   1305
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
         TabIndex        =   57
         Top             =   1350
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
         TabIndex        =   72
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
         TabIndex        =   50
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
         Height          =   285
         Left            =   5580
         TabIndex        =   49
         Top             =   150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   285
         Left            =   3795
         TabIndex        =   48
         Top             =   150
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   60
         TabIndex        =   64
         Top             =   1170
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
         Left            =   6360
         TabIndex        =   67
         Top             =   1560
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
         Height          =   3015
         Left            =   6360
         TabIndex        =   66
         Top             =   1560
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5318
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
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
         Left            =   6330
         TabIndex        =   54
         Top             =   660
         Width           =   690
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
         Height          =   255
         Left            =   45
         TabIndex        =   97
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   98
         Top             =   150
         Width           =   1530
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
         TabIndex        =   99
         Top             =   585
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   100
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   101
         Top             =   555
         Width           =   1185
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
         TabIndex        =   102
         Top             =   585
         Width           =   870
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
         TabIndex        =   103
         Top             =   975
         Width           =   1140
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
         TabIndex        =   104
         Top             =   1335
         Visible         =   0   'False
         Width           =   2475
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
         TabIndex        =   105
         Top             =   1335
         Width           =   2340
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
         TabIndex        =   106
         Top             =   1335
         Visible         =   0   'False
         Width           =   2490
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
         Height          =   225
         Left            =   1800
         TabIndex        =   107
         Top             =   3075
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   108
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   109
         Top             =   3060
         Width           =   1155
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
         TabIndex        =   110
         Top             =   585
         Width           =   480
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
         Height          =   195
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   111
         Top             =   165
         Width           =   1035
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
         Height          =   195
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   112
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4665
      Index           =   3
      Left            =   180
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   9075
      Begin VB.Frame Frame5 
         Caption         =   "Retenções"
         Height          =   615
         Left            =   4440
         TabIndex        =   113
         Top             =   0
         Width           =   3795
         Begin VB.Frame SSFrame6 
            Caption         =   "ISS"
            Height          =   915
            Left            =   255
            TabIndex        =   114
            Top             =   1365
            Visible         =   0   'False
            Width           =   3885
            Begin VB.CheckBox ISSRetido 
               Caption         =   "Retido"
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
               Left            =   2475
               TabIndex        =   115
               Top             =   435
               Width           =   900
            End
            Begin VB.Label Label30 
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
               Height          =   210
               Left            =   450
               TabIndex        =   117
               Top             =   450
               Width           =   510
            End
            Begin VB.Label ISSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1065
               TabIndex        =   116
               Top             =   390
               Width           =   1305
            End
         End
         Begin MSMask.MaskEdBox ValorIRRF 
            Height          =   300
            Left            =   570
            TabIndex        =   118
            Top             =   225
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox INSSValor 
            Height          =   300
            Left            =   2460
            TabIndex        =   120
            Top             =   225
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "INSS:"
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
            Left            =   1920
            TabIndex        =   121
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "I.R.:"
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
            TabIndex        =   119
            Top             =   285
            Width           =   390
         End
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   4080
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoFatura 
         Height          =   300
         Left            =   3000
         TabIndex        =   29
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   3915
         Left            =   645
         TabIndex        =   70
         Top             =   600
         Width           =   7590
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   3270
            TabIndex        =   30
            Top             =   255
            Width           =   1815
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   255
            Left            =   2805
            TabIndex        =   32
            Top             =   1125
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   255
            Left            =   4860
            TabIndex        =   33
            Top             =   1110
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   255
            Left            =   975
            TabIndex        =   31
            Top             =   1140
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2820
            Left            =   270
            TabIndex        =   34
            Top             =   735
            Width           =   6555
            _ExtentX        =   11562
            _ExtentY        =   4974
            _Version        =   393216
            Rows            =   50
            Cols            =   4
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
            Left            =   1050
            TabIndex        =   85
            Top             =   315
            Width           =   2175
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Data de Emissão da Fatura:"
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
         Left            =   615
         TabIndex        =   86
         Top             =   240
         Width           =   2370
      End
   End
   Begin VB.CommandButton BotaoFaturasGeradas 
      Caption         =   "Faturas Geradas"
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
      Left            =   6142
      TabIndex        =   12
      Top             =   90
      Width           =   1800
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   8070
      Picture         =   "GeracaoFaturaOcx.ctx":2CCA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Fechar"
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton BotaoGerar 
      Caption         =   "Gerar Fatura"
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
      Left            =   4215
      TabIndex        =   11
      Top             =   90
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4665
      Index           =   4
      Left            =   180
      TabIndex        =   35
      Top             =   840
      Visible         =   0   'False
      Width           =   9075
      Begin VB.ListBox Vendedores 
         Height          =   3960
         Left            =   6960
         TabIndex        =   46
         Top             =   645
         Width           =   2010
      End
      Begin MSComCtl2.UpDown UpDownParcela 
         Height          =   300
         Left            =   3240
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   150
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   2115
         Left            =   105
         TabIndex        =   76
         Top             =   2475
         Width           =   6690
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   315
            Left            =   780
            TabIndex        =   41
            Top             =   555
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   4740
            TabIndex        =   44
            Top             =   600
            Width           =   1275
            _ExtentX        =   2249
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   3435
            TabIndex        =   43
            Top             =   600
            Width           =   1275
            _ExtentX        =   2249
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   225
            Left            =   2370
            TabIndex        =   42
            Top             =   570
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
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
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1320
            Left            =   180
            TabIndex        =   45
            Top             =   315
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   2328
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label LabelTotaisComissoes 
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
            Height          =   225
            Left            =   1455
            TabIndex        =   87
            Top             =   1740
            Width           =   705
         End
         Begin VB.Label TotalValorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3405
            TabIndex        =   88
            Top             =   1710
            Width           =   1155
         End
         Begin VB.Label TotalPercentualComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2325
            TabIndex        =   89
            Top             =   1710
            Width           =   945
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Descontos"
         Height          =   1815
         Left            =   90
         TabIndex        =   77
         Top             =   555
         Width           =   6690
         Begin VB.ComboBox TipoDesconto 
            Height          =   315
            Left            =   675
            TabIndex        =   36
            Top             =   450
            Width           =   2490
         End
         Begin MSMask.MaskEdBox Percentual1 
            Height          =   225
            Left            =   5355
            TabIndex        =   39
            Top             =   495
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   225
            Left            =   4290
            TabIndex        =   38
            Top             =   495
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox Data 
            Height          =   285
            Left            =   3180
            TabIndex        =   37
            Tag             =   "1"
            Top             =   435
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDescontos 
            Height          =   1320
            Left            =   180
            TabIndex        =   40
            Top             =   300
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   2328
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Label Parcela 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   90
         Top             =   165
         Width           =   330
      End
      Begin VB.Label Label21 
         Caption         =   "Parcela:"
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
         Left            =   2145
         TabIndex        =   91
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Vendedores"
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
         Left            =   6945
         TabIndex        =   92
         Top             =   420
         Width           =   1065
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5145
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   9075
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notas Fiscais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Descontos/Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "GeracaoFaturaOcx"
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

'Mnemônicos
Private Const CLIENTE1 As String = "Cliente"
Private Const FILIAL1 As String = "Filial"
Private Const FILIAL_COD As String = "Codigo_Filial"
Private Const Data1 As String = "Data"
Private Const VALORFATURA As String = "ValorFatura"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const VALOR_IRRF As String = "Valor_IRRF"


'Fim da Contabilidade

Public iAlterado As Integer
Dim iTabPrincipalAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFrameAtual As Integer
Dim colcolDesconto As colcolDesconto
Dim colcolComissao As colcolComissao

Dim objGridNFiscal As AdmGrid
Dim objGridParcelas As AdmGrid
Dim objGridDesconto As AdmGrid
Dim objGridComissoes As AdmGrid

Dim iGrid_Parcela_Col As Integer
Dim iGrid_Vencimento_col As Integer
Dim iGrid_VenctoReal_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_PercentualComissao_Col As Integer
Dim iGrid_ValorBase_Col As Integer
Dim iGrid_ValorComissao_Col As Integer
Dim iGrid_TipoDesconto_Col As Integer
Dim iGrid_DataDesconto_Col As Integer
Dim iGrid_ValorDesconto_Col As Integer
Dim iGrid_PercentualDesconto_Col As Integer
Dim iGrid_GeraFatura_Col As Integer
Dim iGrid_Serie_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_SiglaDoc_Col As Integer
Dim iGrid_ValorTotal_Col As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

'Ordenacao
Const NFISCAL_ORDEM_DATA = "DataEmissao, Serie, NumNotaFiscal"

'Constantes públicas dos tabs
Private Const TAB_Selecao = 1
Private Const TAB_NFs = 2
Private Const TAB_Cobranca = 3
Private Const TAB_Descontos = 4
'Private Const TAB_Tributacao = 5
Private Const TAB_Contabilizacao = 5

Public Sub Form_Load()

Dim lErro As Long
Dim lNumInt As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
    
        FilialCliente.Visible = False
        LblFilialCli.Visible = False
    
    End If
    
    iFrameAtual = 1
    
    If gobjFAT.iGeraFatTodasFiliais = MARCADO Then
        optTodasFiliais.Value = vbChecked
    Else
        optTodasFiliais.Value = vbUnchecked
    End If

    'Inicializa as variáveis globais da tela
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoSerie = New AdmEvento

    Set colcolComissao = New colcolComissao
    Set colcolDesconto = New colcolDesconto

    'Carrega a combo de Séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then Error 31182

    'Carrega a combo de Condições de Pagamento
    lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 31184

    'Carrega a Lista de Vendedores
    lErro = Carrega_Vendedores()
    If lErro <> SUCESSO Then Error 31196

    'Carrega a combo de Tipo de Desconto
    lErro = Carrega_TipoDesconto()
    If lErro <> SUCESSO Then Error 31222

    'Faz as inicializações do Grid de Notas Fiscais
    Set objGridNFiscal = New AdmGrid

    lErro = Inicializa_Grid_NFiscal(objGridNFiscal)
    If lErro <> SUCESSO Then Error 31186

    'Faz as inicializações no Grid de Parcelas
    Set objGridParcelas = New AdmGrid

    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 31193

    'Faz as inicializações no Grid de Descontos
    Set objGridDesconto = New AdmGrid

    lErro = Inicializa_Grid_Descontos(objGridDesconto)
    If lErro <> SUCESSO Then Error 31194

    'Faz as inicializações no Grid de Comissões
    Set objGridComissoes = New AdmGrid

    lErro = Inicializa_Grid_Comissoes(objGridComissoes)
    If lErro <> SUCESSO Then Error 31195

    'Coloca como "default" gdtDataAtual p/data de emissão da fatura
    EmissaoFatura.PromptInclude = False
    EmissaoFatura.Text = Format(gdtDataAtual, "dd/mm/yy")
    EmissaoFatura.PromptInclude = True

    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO)
    If lErro <> SUCESSO Then Error 39760

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 31182, 31184, 31186, 31193, 31194, 31195, 31196, 31222, 39760

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160785)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then Error 31183

    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = Err

    Select Case Err

        Case 31183

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160786)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_CondicaoPagamento() As Long
''Carrega a combo de Condições de Pagamento com  as Condições lidas do BD
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then Error 31185
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Adiciona novo ítem na List da Combo CondicaoPagamento
'        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo
'
'    Next
'
'    Carrega_CondicaoPagamento = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CondicaoPagamento:
'
'    Carrega_CondicaoPagamento = Err
'
'    Select Case Err
'
'        Case 31185
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160787)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_TipoDesconto() As Long
'Carrega na combo os Tiops de descontos

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    'Lê o código e a descrição de todos os Tipos de Desconto
    lErro = CF("Cod_Nomes_Le", "TiposDeDesconto", "Codigo", "DescReduzida", 50, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 31223

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na List da Combo TipoDesconto
        TipoDesconto.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TipoDesconto.ItemData(TipoDesconto.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = Err

    Select Case Err

        Case 31223

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160788)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_NFiscal(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Gera Fatura")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Cliente")
    If giTipoVersao <> VERSAO_LIGHT Then objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Valor")

    'Campos de edição do grid
    objGridInt.colCampo.Add (GeraFatura.Name)
    objGridInt.colCampo.Add (Serie1.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    If giTipoVersao <> VERSAO_LIGHT Then
        objGridInt.colCampo.Add (Filial.Name)
    Else
        Filial.left = -20000
    End If
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (SiglaDoc.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)

    'Inicializa as variáveis que guardarão as colunas do Grid
    iGrid_GeraFatura_Col = 1
    iGrid_Serie_Col = 2
    iGrid_Numero_Col = 3
    iGrid_Cliente_Col = 4
    
    If giTipoVersao <> VERSAO_LIGHT Then
    
        iGrid_Filial_Col = 5
        iGrid_DataEmissao_Col = 6
        iGrid_SiglaDoc_Col = 7
        iGrid_ValorTotal_Col = 8

    Else
    
        iGrid_DataEmissao_Col = 5
        iGrid_SiglaDoc_Col = 6
        iGrid_ValorTotal_Col = 7
    
    End If
    
    objGridInt.objGrid = GridNFiscal

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9

    GridNFiscal.ColWidth(0) = 300

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Faz a inicializações mais internas
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalNotasSel.top = GridNFiscal.top + GridNFiscal.Height
    TotalNotasSel.left = GridNFiscal.left
    For iIndice = 0 To iGrid_ValorTotal_Col - 1
        TotalNotasSel.left = TotalNotasSel.left + GridNFiscal.ColWidth(iIndice) + GridNFiscal.GridLineWidth
    Next

    TotalNotasSel.Width = GridNFiscal.ColWidth(iGrid_ValorTotal_Col)

    LabelTotalNotaSel.top = TotalNotasSel.top + (TotalNotasSel.Height - LabelTotalNotaSel.Height) / 2
    LabelTotalNotaSel.left = TotalNotasSel.left - LabelTotalNotaSel.Width

    Inicializa_Grid_NFiscal = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Data de Vencimento")
    objGridInt.colColuna.Add ("Data de Vencimento Real")
    objGridInt.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)

    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_Valor_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 705

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Descontos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Descontos

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Percentual")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (TipoDesconto.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (ValorDesconto.Name)
    objGridInt.colCampo.Add (Percentual1.Name)

    'colunas do Grid
    iGrid_TipoDesconto_Col = 1
    iGrid_DataDesconto_Col = 2
    iGrid_ValorDesconto_Col = 3
    iGrid_PercentualDesconto_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridDescontos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_DESCONTOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridDescontos.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Descontos = SUCESSO

    Exit Function

End Function


Private Function Inicializa_Grid_Comissoes(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Comissoes

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Vendedor.Name)
    objGridInt.colCampo.Add (PercentualComissao.Name)
    objGridInt.colCampo.Add (ValorBase.Name)
    objGridInt.colCampo.Add (ValorComissao.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridComissoes

    'Colunas do Grid
    iGrid_Vendedor_Col = 1
    iGrid_PercentualComissao_Col = 2
    iGrid_ValorBase_Col = 3
    iGrid_ValorComissao_Col = 4

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_COMISSOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridComissoes.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalPercentualComissao.top = GridComissoes.top + GridComissoes.Height
    TotalPercentualComissao.left = GridComissoes.left
    For iIndice = 0 To 1
        TotalPercentualComissao.left = TotalPercentualComissao.left + GridComissoes.ColWidth(iIndice) + GridComissoes.GridLineWidth + 20
    Next

    TotalPercentualComissao.Width = GridComissoes.ColWidth(2)

    TotalValorComissao.top = TotalPercentualComissao.top
    TotalValorComissao.Width = GridComissoes.ColWidth(4)
    For iIndice = 0 To iGrid_ValorComissao_Col - 1
        TotalValorComissao.left = TotalPercentualComissao.left + TotalPercentualComissao.Width + GridComissoes.ColWidth(iIndice) + GridComissoes.GridLineWidth + 20
    Next

    LabelTotaisComissoes.top = TotalPercentualComissao.top + (TotalPercentualComissao.Height - LabelTotaisComissoes.Height) / 2
    LabelTotaisComissoes.left = TotalPercentualComissao.left - LabelTotaisComissoes.Width

    Inicializa_Grid_Comissoes = SUCESSO

    Exit Function

End Function

Private Function Carrega_Vendedores() As Long

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Vendedores

    'Lê o Código e o Nome Reduzido dos Vendedores
    lErro = CF("Cod_Nomes_Le", "Vendedores", "Codigo", "NomeReduzido", STRING_VENDEDOR_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 31197

    For Each objCodNome In colCodigoNome

        'Adiciona o ítem na List de Vendedores
        Vendedores.AddItem objCodNome.iCodigo & SEPARADOR & objCodNome.sNome
        Vendedores.ItemData(Vendedores.NewIndex) = objCodNome.iCodigo

    Next

    Carrega_Vendedores = SUCESSO

    Exit Function

Erro_Carrega_Vendedores:

    Carrega_Vendedores = Err

    Select Case Err

        Case 31197

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160789)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodos_Click()
'Desseleciona todas as Notas Fiscais do Grid

Dim iIndice As Integer
Dim dValorNFsSelecionadas As Double
Dim lErro As Long

On Error GoTo Erro_BotaoDesmarcarTodos_Click

    'Verifica se existe alguma Nota Fiscal no Grid de Notas Fiscais
    If objGridNFiscal.iLinhasExistentes <= 0 Then Exit Sub

    'desseleciona todas as Notas Fiscais no Grid
    For iIndice = 1 To objGridNFiscal.iLinhasExistentes
        GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col) = 0
    Next

    'Faz o Refresh nas Checkboxes do Grid
    Call Grid_Refresh_Checkbox(objGridNFiscal)
        
    'Calcula o Valor das Notas selecionadas (Total)
    Call Calcula_Valor_Notas_Selecionadas(dValorNFsSelecionadas)
    
    'Calcula os totais
    TotalNotasSel.Caption = Format(dValorNFsSelecionadas, "Standard")
    
    'Calcula as Comissoes para as Notas Fiscais
    lErro = Calcula_Comissoes()
    If lErro <> SUCESSO Then Error 58448
    
    'Calcula os Descontos
    lErro = Calcula_Descontos()
    If lErro <> SUCESSO Then Error 58449
    
    'Para na Mudança do Tab Trazer as Novas Comissoes
    Parcela.Caption = ""

    Exit Sub

Erro_BotaoDesmarcarTodos_Click:

    Select Case Err

        Case 58448, 58449 'Tratados na rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160790)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFaturasGeradas_Click()

Dim objTituloReceber As New ClassTituloReceber
Dim colParcelas As New colParcelaReceber
Dim colSelecao As New Collection
Dim lErro As Long
Dim objcliente As New ClassCliente
Dim sSelecao As String
Dim iPreenchido As Integer

On Error GoTo Erro_BotaoFaturasGeradas_Click

    'Se Cliente estiver vazio, erro
    If Len(Trim(CodCliente.Text)) > 0 Then

        objcliente.sNomeReduzido = CodCliente.Text
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 87216

        'Se não encontrou o Cliente --> Erro
        If lErro <> SUCESSO Then gError 87217

    End If
    
    'Guarda o código no objTituloReceber
    objTituloReceber.lCliente = objcliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(FilialCliente.Text)
    objTituloReceber.sSiglaDocumento = TIPODOC_FATURA_A_RECEBER

    'Verifica se os obj(s) estão preenchidos antes de serem incluídos na coleção
    If objTituloReceber.lCliente <> 0 Then
        sSelecao = "Cliente = ?"
        iPreenchido = 1
        colSelecao.Add (objTituloReceber.lCliente)
    End If

    If objTituloReceber.iFilial <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Filial = ?"
        Else
            iPreenchido = 1
            sSelecao = "Filial = ?"
        End If
        colSelecao.Add (objTituloReceber.iFilial)
    End If

    If iPreenchido = 1 Then
        sSelecao = sSelecao & " AND SiglaDocumento = ?"
    Else
        iPreenchido = 1
        sSelecao = "SiglaDocumento = ?"
    End If
        colSelecao.Add (objTituloReceber.sSiglaDocumento)

    'Chama Tela TituloReceberLista
    Call Chama_Tela("TituloReceberLista", colSelecao, objTituloReceber, objEventoNumero, sSelecao)

    Exit Sub

Erro_BotaoFaturasGeradas_Click:

    Select Case gErr

        Case 87216

        Case 87217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160791)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()
'Seleciona todas as notas fiscais do Grid de Notas Fiscais

Dim iIndice As Integer
Dim dValorNFsSelecionadas As Double
Dim lErro As Long

On Error GoTo Erro_BotaoMarcarTodos_Click

    'Verifica se existe alguma Nota Fiscal no Grid
    If objGridNFiscal.iLinhasExistentes <= 0 Then Exit Sub

    'Seleciona e atulializa o valor das Notas Fiscais selecionadas
    For iIndice = 1 To objGridNFiscal.iLinhasExistentes
        GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col) = 1
    Next

    'Faz o Refresh nas CheckBoxes do Grid
    Call Grid_Refresh_Checkbox(objGridNFiscal)

    'Calcula o Valor das Notas selecionadas (Total)
    Call Calcula_Valor_Notas_Selecionadas(dValorNFsSelecionadas)
    
    'Calcula os totais
    TotalNotasSel.Caption = Format(dValorNFsSelecionadas, "Standard")
    
    'Calcula as Comissoes para as Notas Fiscais
    lErro = Calcula_Comissoes()
    If lErro <> SUCESSO Then Error 58445
    
    'Calcula os Descontos
    lErro = Calcula_Descontos()
    If lErro <> SUCESSO Then Error 58446
    
    'Para na Mudança do Tab Trazer as Novas Comissoes
    Parcela.Caption = ""

    Exit Sub

Erro_BotaoMarcarTodos_Click:

    Select Case Err

        Case 58445, 58446 'Tratados na rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160792)

    End Select

    Exit Sub
 
End Sub

Private Sub BotaoNFiscal_Click()

Dim lErro As Long
Dim iAchou As Integer
Dim objNFiscal As New ClassNFiscal
Dim iNotaFiscal As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_BotaoNFiscal_Click

    If objGridNFiscal.iLinhasExistentes <= 0 Then Exit Sub

    If GridNFiscal.Row <= 0 Then Error 58301
    
    objNFiscal.lNumNotaFiscal = GridNFiscal.TextMatrix(GridNFiscal.Row, iGrid_Numero_Col)
    objNFiscal.sSerie = GridNFiscal.TextMatrix(GridNFiscal.Row, iGrid_Serie_Col)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataEmissao = CDate(GridNFiscal.TextMatrix(GridNFiscal.Row, iGrid_DataEmissao_Col))
    
    'Lê o NumIntDoc da NFiscal
    lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
    If lErro <> SUCESSO And lErro <> 43676 Then Error 58404
    
    'Se não encontrou
    If lErro = 43676 Then Error 58405
    
'    If objNFiscal.iTipoNFiscal = DOCINFO_NFISV Then
'        'Chama a Tela de Notas Fiscais
'        Call Chama_Tela("NFiscal", objNFiscal)
'
'    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NFICF Then
'
'        Call Chama_Tela("ConhecimentoFrete", objNFiscal)
'
'    Else
'        'Chama a Tela de Notas Fiscais
'        Call Chama_Tela("NFiscalPedido", objNFiscal)
'    End If

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    'lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 58404

    Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
    
    Exit Sub

Erro_BotaoNFiscal_Click:

    Select Case Err

        Case 58301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
        
        Case 58404 'Tratado na Rotina chamada
        
        Case 58405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NUM_SERIE_NAO_CADASTRADA", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160793)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub DataDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub EmissaoFatura_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EmissaoFatura_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoFatura, iAlterado)

End Sub

Private Sub FilialCliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCliente_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialCliente.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialCliente.Text = FilialCliente.List(FilialCliente.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCliente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 31282

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(CodCliente.Text)) = 0 Then Error 31283

        sCliente = CodCliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 31284

        If lErro = 17660 Then Error 31285

        'Coloca na tela a Filial lida
        FilialCliente.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then Error 31286

    Exit Sub

Erro_FilialCliente_Validate:

    Cancel = True


    Select Case Err

        Case 31282, 31284

        Case 31283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 31285
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 31286
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160794)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub CodCliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub CodCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_CodCliente_Validate

    If iClienteAlterado = 1 Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(CodCliente.Text)) > 0 Then
            
            lErro = TP_Cliente_Le(CodCliente, objcliente, iCodFilial)
            If lErro <> SUCESSO Then Error 31198

            lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
            If lErro <> SUCESSO Then Error 31199

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

            'Seleciona filial na Combo Filial
             Call CF("Filial_Seleciona", FilialCliente, iCodFilial)

        'Se não estiver preenchido
        ElseIf Len(Trim(CodCliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            FilialCliente.Clear

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_CodCliente_Validate:

    Cancel = True


    Select Case Err

        Case 31198
            
        Case 31199

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160795)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then Error 31201

    'Verifica se a Data Inicial é maior que a Data Final no Intervalo dados
    If Len(Trim(DataAte.ClipText)) > 0 Then
        If CDate(DataDe.Text) > CDate(DataAte.Text) Then Error 31209
    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case Err

        Case 31201

        Case 31209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160796)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'eventos e grids
    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set colcolComissao = Nothing
    Set colcolDesconto = Nothing

    Set objEventoSerie = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing

    Set objGridNFiscal = Nothing
    Set objGridComissoes = Nothing
    Set objGridDesconto = Nothing
    Set objGridParcelas = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

End Sub

Private Sub GeraFatura_Click()
'trata checkbox no grid que seleciona as notas fiscais a serem faturadas

Dim iClick As Integer
Dim lErro As Long
Dim dValorNotasSelecionadas As Double

On Error GoTo Erro_GeraFatura_Click

    'Verifica se é alguma linha válida
    If GridNFiscal.Row > objGridNFiscal.iLinhasExistentes Then Exit Sub

    'Verifica se está selecionando ou desselecionando
    If Len(Trim(GridNFiscal.TextMatrix(GridNFiscal.Row, iGrid_GeraFatura_Col))) > 0 Then
        iClick = CInt(GridNFiscal.TextMatrix(GridNFiscal.Row, iGrid_GeraFatura_Col)) = 1
    End If

    'Se está selecionando
    If iClick = True Then
        Call Calcula_Valor_Notas_Selecionadas(dValorNotasSelecionadas)
        TotalNotasSel.Caption = Format(dValorNotasSelecionadas, "Standard")
    'Senão
    Else
        'Abate o valor da NF no Valor dsa Notas Fiscais Selecionadas
        Call Calcula_Valor_Notas_Selecionadas(dValorNotasSelecionadas)
        TotalNotasSel.Caption = Format(dValorNotasSelecionadas, "Standard")

    End If
        
    'Calcula as Comissoes para as Notas Selecionadas
    lErro = Calcula_Comissoes()
    If lErro <> SUCESSO Then Error 58406
    
    'Calcula os Descontos
    lErro = Calcula_Descontos()
    If lErro <> SUCESSO Then Error 58407
    
    'Limpa a Parcela para que o Grid de Comissoes seja Preenchido com a Comissao da Parcela 1
    Parcela.Caption = ""
    
    Exit Sub

Erro_GeraFatura_Click:

    Select Case Err
        
        Case 58406, 58407 'Tratados nas Rotinas Chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160797)

    End Select

    Exit Sub

End Sub

Private Sub ISSRetido_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelSerie_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As New ClassSerie
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSerie_Click

    'Seleciona serie para objSerie
    For iIndice = 0 To Serie.ListCount - 1
        If Serie.List(iIndice) = Serie.Text Then
            Serie.ListIndex = iIndice
            objSerie.sSerie = Serie.Text
            Exit For
        End If
    Next

    'Chama lista das Series
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

Erro_LabelSerie_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160798)

    End Select

    Exit Sub


End Sub

Private Sub NFiscalFinal_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(NFiscalFinal, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub NFiscalInicial_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(NFiscalInicial, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub


Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objSerie As ClassSerie
Dim iIndice As Integer

On Error GoTo Erro_objEventoSerie_evSelecao

    Set objSerie = obj1
    
    'Preenche a Série na Tela
    For iIndice = 0 To Serie.ListCount - 1
        If Serie.List(iIndice) = objSerie.sSerie Then
            Serie.ListIndex = iIndice
            Exit For
        End If
    Next

    Me.Show

    Exit Sub

Erro_objEventoSerie_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160799)

    End Select

    Exit Sub

End Sub

Private Sub LblCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = CodCliente.Text
    
    'Chama a tela que lista os Clientes
    Call Chama_Tela("Cliente_NFFaturarLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub NFiscalFinal_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscalInicial_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscalInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalInicial_Validate

    'Verifica se foi preenchida
    If Len(Trim(NFiscalInicial.ClipText)) = 0 Then Exit Sub

    'Verifica se o valor inicial é maior que o valor final para o intervalo dado
    If Len(Trim(NFiscalFinal.ClipText)) > 0 Then
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then Error 31210
    End If

    Exit Sub

Erro_NFiscalInicial_Validate:

    Cancel = True

    Select Case Err

        Case 31210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160800)

    End Select

    Exit Sub

End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalFinal_Validate

    'Verifica se está preenchido
    If Len(Trim(NFiscalFinal.ClipText)) = 0 Then Exit Sub

    'Verifica se o valor Final é maior que o valor inicial no Intervalo dado
    If Len(Trim(NFiscalInicial.ClipText)) > 0 Then
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then Error 31211
    End If

    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True

    Select Case Err

        Case 31211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160801)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente

    Set objcliente = obj1

    'Preenche o Codigo do Cliente e dispara o Validate
    CodCliente.Text = objcliente.sNomeReduzido
    Call CodCliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub optTodasFiliais_Click()
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    'Verifica se a Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub

    'Verifica se é uma Serie selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 31212

    If lErro = 12253 Then Error 31213

    Exit Sub

Erro_Serie_Validate:

    Cancel = True


    Select Case Err

        Case 31212

        Case 31213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160802)

    End Select

    Exit Sub

End Sub

Private Sub TotalNotasSel_Change()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim lErro As Long
Dim dValorNFsSelecionadas As Double

On Error GoTo Erro_TotalNotasSel_Change

    If CondicaoPagamento.ListIndex = -1 Then Exit Sub

    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
    
    'Le a Condicao de Pagamento no BD
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then Error 51123
    If lErro <> SUCESSO Then Error 51124

    'Limpa o Grid Parcelas
    Call Grid_Limpa(objGridParcelas)
    
    Call Calcula_Valor_Notas_Selecionadas(dValorNFsSelecionadas)
    
    'Testa se EmissaoFatura está preenchida
    If Len(Trim(EmissaoFatura.ClipText)) > 0 And dValorNFsSelecionadas > 0 Then

        'Preenche o GridParcelas
        lErro = GridParcelas_Preenche(objCondicaoPagto, dValorNFsSelecionadas)
        If lErro <> SUCESSO Then Error 51125

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_TotalNotasSel_Change:

    Select Case Err

        Case 51123, 51125

        Case 51124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160803)

      End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro Then Error 31202

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case Err

        Case 31202

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160804)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro Then Error 31203

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case Err

        Case 31203

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160805)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then Error 31204

    'Verifica se a Data Inicial é menor que a Data Final para o intervalo formado
    If Len(Trim(DataDe.ClipText)) > 0 Then
        If CDate(DataDe.Text) > CDate(DataAte.Text) Then Error 31208
    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case Err

        Case 31204

        Case 31208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160806)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro Then Error 31205

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case Err

        Case 31205

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160807)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro Then Error 31206

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case Err

        Case 31206

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160808)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim lNumInt As Long

On Error GoTo Erro_BotaoGerar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 46186

    'Limpa a Tela
    Call Limpa_Tela_GeracaoFatura

    'Coloca como "default" gdtDataAtual p/data de emissão da fatura
    EmissaoFatura.PromptInclude = False
    EmissaoFatura.Text = Format(gdtDataAtual, "dd/mm/yy")
    EmissaoFatura.PromptInclude = True
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGerar_Click:

    Select Case Err

        Case 46186

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160809)

    End Select

    Exit Sub

End Sub


Private Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

    Frame1(Opcao.SelectedItem.Index).Visible = True
    Frame1(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = Opcao.SelectedItem.Index

    'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
    If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

    'Trata a mudança de frame
    lErro = Trata_Mudanca_Frame()
    If lErro <> SUCESSO Then Error 31214

    Select Case iFrameAtual
    
        Case TAB_Selecao
            Parent.HelpContextID = IDH_GERACAO_FATURA_SELECAO
            
        Case TAB_NFs
            Parent.HelpContextID = IDH_GERACAO_FATURA_NF
            
        Case TAB_Cobranca
            Parent.HelpContextID = IDH_GERACAO_FATURA_COBRANCA
            
        Case TAB_Descontos
            Parent.HelpContextID = IDH_GERACAO_FATURA_DESCONTOS
            
''        Case TAB_Tributacao
''            Parent.HelpContextID = IDH_GERACAO_FATURA_TRIBUTACAO
            
        Case TAB_Contabilizacao
            Parent.HelpContextID = IDH_GERACAO_FATURA_CONTABILIZACAO
    
    End Select
    
    Exit Sub

Erro_Opcao_Click:

    Select Case Err

        Case 31214

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160810)

    End Select

    Exit Sub

End Sub


Private Sub ValorIRRF_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIRRF_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIRRF.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
    If lErro <> SUCESSO Then Error 31207

    'Põe o valor formatado na tela
    ValorIRRF.Text = Format(ValorIRRF.Text, "Fixed")

    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True


    Select Case Err

        Case 31207

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160811)

    End Select

    Exit Sub

End Sub

Private Function Trata_Mudanca_Frame() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Mudanca_Frame

    'Verifica qual é o frame ativo
    Select Case iFrameAtual

        'Tab Notas Fiscais
        Case 2
            If Len(Trim(CodCliente.Text)) = 0 Then Error 64135
            
            'Verifica se algum campo do frame principal foi alterado
            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
                lErro = Trata_TabNFs()
                If lErro <> SUCESSO Then Error 31215
            End If

        'Tab de Tributacao
        Case 5
            lErro = Trata_TabTributacao()
            If lErro <> SUCESSO Then Error 31275

        'Tab de Parcelas
        Case 4
            lErro = Trata_TabDescontosComissoes()
            If lErro <> SUCESSO Then Error 31224

    End Select

    Trata_Mudanca_Frame = SUCESSO

    Exit Function

Erro_Trata_Mudanca_Frame:

    Trata_Mudanca_Frame = Err

    Select Case Err

        Case 31215, 31275, 31224

        Case 64135
            'Limpa o Grid de Notas Fiscais
            Call Grid_Limpa(objGridNFiscal)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_INFORMADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160812)

    End Select

    Exit Function

End Function

Private Function Trata_TabNFs() As Long
'Faz o tratamento no tab de Notas Fiscais

Dim lErro As Long
Dim iIndice As Integer
Dim objNFiscalInfo As ClassNFiscalInfo
Dim objcliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice2 As Integer
Dim objGeracaoFatura As New ClassGeracaoFatura

On Error GoTo Erro_Trata_TabNFs
    
    'Carrega o objGeracaoFatura com os dados do tab principal
    Call Move_TabSelecao_Memoria(objGeracaoFatura)

    'Limpa o Grid de Notas Fiscais
    Call Grid_Limpa(objGridNFiscal)
    TotalNotasSel.Caption = ""

    'Verifica se o Cliente e a Filial estão preenchidos
    'If objGeracaoFatura.lCliente = 0 Or objGeracaoFatura.iFilialCli = 0 Then Error 31287
    If objGeracaoFatura.lCliente = 0 Or Codigo_Extrai(FilialCliente.Text) = 0 Then Error 31287

    objGeracaoFatura.sOrdenacao = NFISCAL_ORDEM_DATA

    'Obtém as Notas Fiscais com as características passadas em objGeracaoFatura
    lErro = CF("GeracaoFatura_ObterNFs", objGeracaoFatura)
    If lErro <> SUCESSO Then Error 31272

    iIndice = 0

    objGridNFiscal.objGrid.Rows = objGridNFiscal.iLinhasVisiveis + 1

    If objGeracaoFatura.colNFiscalInfo.Count >= objGridNFiscal.objGrid.Rows Then
        objGridNFiscal.objGrid.Rows = objGeracaoFatura.colNFiscalInfo.Count + 1
    End If

    Call Grid_Inicializa(objGridNFiscal)

    'Coloca na tela as Notas Fiscais lidas
    For Each objNFiscalInfo In objGeracaoFatura.colNFiscalInfo
        iIndice = iIndice + 1
        GridNFiscal.TextMatrix(iIndice, iGrid_Serie_Col) = objNFiscalInfo.sSerie
        GridNFiscal.TextMatrix(iIndice, iGrid_Numero_Col) = objNFiscalInfo.lNumero
        GridNFiscal.TextMatrix(iIndice, iGrid_Cliente_Col) = objNFiscalInfo.sClienteNomeReduzido
        GridNFiscal.TextMatrix(iIndice, iGrid_DataEmissao_Col) = objNFiscalInfo.dtEmissao
        GridNFiscal.TextMatrix(iIndice, iGrid_SiglaDoc_Col) = objNFiscalInfo.sSiglaDoc
        GridNFiscal.TextMatrix(iIndice, iGrid_ValorTotal_Col) = Format(objNFiscalInfo.dValorTotal, "Standard")
        If giTipoVersao <> VERSAO_LIGHT Then GridNFiscal.TextMatrix(iIndice, iGrid_Filial_Col) = objNFiscalInfo.iFilialCliente
    Next

    'Atualiza o número de linhas existentes do Grid
    objGridNFiscal.iLinhasExistentes = iIndice

    'zera a flag dos campos do tab principal
    iTabPrincipalAlterado = 0

    Trata_TabNFs = SUCESSO

    Exit Function

Erro_Trata_TabNFs:

    Trata_TabNFs = Err

    Select Case Err

        Case 31217, 31287, 31272
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160813)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Critica a Data de Vencimento e gera a Data de Vencto Real
        Case iGrid_Vencimento_col
            lErro = Saida_Celula_Vencimento(objGridInt)
            If lErro <> SUCESSO Then Error 31218

        'Faz a critica do valor da Parcela
        Case iGrid_Valor_Col
            lErro = Saida_Celula_Valor(objGridInt)
            If lErro <> SUCESSO Then Error 31219

    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = Err

    Select Case Err

        Case 31218, 31219

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160814)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vencimento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim dtDataVenctoReal As Date

On Error GoTo Erro_Saida_Celula_Vencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 31221

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(EmissaoFatura.ClipText)) > 0 Then
            dtDataEmissao = CDate(EmissaoFatura.Text)
            If dtDataVencimento < dtDataEmissao Then Error 31304
        End If

        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then Error 31305

        'Coloca data de Vencimento Real no Grid
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31306

    Saida_Celula_Vencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_Vencimento:

    Saida_Celula_Vencimento = Err

    Select Case Err

        Case 31221, 31305, 31306
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, DataVencimento.Text, EmissaoFatura.Text, GridParcelas.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160815)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then Error 31307

        Call Calcula_Comissoes_Parcela(GridParcelas.Row, CDbl(ValorParcela.Text))
        
        If Len(Trim(Parcela.Caption)) > 0 Then
            If GridParcelas.Row = CLng(Parcela.Caption) Then
                Call Traz_Parcela_Tela(CInt(Parcela.Caption))
            End If
        End If
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31308

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 31307, 31308
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160816)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridDescontos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDescontos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_TipoDesconto_Col
            'faz a crítica do tipo de desconto
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then Error 31309

        Case iGrid_DataDesconto_Col
            'faz a crítica da Data
            lErro = Saida_Celula_Data(objGridInt)
            If lErro <> SUCESSO Then Error 31310

        Case iGrid_ValorDesconto_Col
            'faz a crítica do Valor do desconto
            lErro = Saida_Celula_ValorDesconto(objGridInt)
            If lErro <> SUCESSO Then Error 31311

        Case iGrid_PercentualDesconto_Col
            'Faz a crítica do Percentual do desconto
            lErro = Saida_Celula_Percentual(objGridInt)
            If lErro <> SUCESSO Then Error 31312

    End Select

    Saida_Celula_GridDescontos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDescontos:

    Saida_Celula_GridDescontos = Err

    Select Case Err

        Case 31309, 31310, 31311, 31312

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160817)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComissoes(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComissoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vendedor_Col
            'faz a crítica do vendedor
            lErro = Saida_Celula_Vendedor(objGridInt)
            If lErro <> SUCESSO Then Error 31314

        Case iGrid_PercentualComissao_Col
            'faz a crítica do Percentual da comissao
            lErro = Saida_Celula_PercentualComissao(objGridInt)
            If lErro <> SUCESSO Then Error 31315

        Case iGrid_ValorBase_Col
            'faz a crítica do Valor Base
            lErro = Saida_Celula_ValorBase(objGridInt)
            If lErro <> SUCESSO Then Error 31316

        Case iGrid_ValorComissao_Col
            'faz a crítica do Valor da Comissao
            lErro = Saida_Celula_ValorComissao(objGridInt)
            If lErro <> SUCESSO Then Error 31317

    End Select

    Saida_Celula_GridComissoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComissoes:

    Saida_Celula_GridComissoes = Err

    Select Case Err

        Case 31314, 31315, 31316, 31317

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160818)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_TipoDesconto

    Set objGridInt.objControle = TipoDesconto

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoDesconto.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoDesconto.Text <> TipoDesconto.List(TipoDesconto.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona(TipoDesconto, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 31319

            'Não foi encontrado
            If lErro = 6730 Then Error 31320
            If lErro = 6731 Then Error 31321

        End If

        iCodigo = Codigo_Extrai(TipoDesconto.Text)

        If iCodigo = VALOR_FIXO Or iCodigo = VALOR_ANT_DIA Or iCodigo = VALOR_ANT_DIA_UTIL Then
            GridDescontos.TextMatrix(GridDescontos.Row, iGrid_PercentualDesconto_Col) = ""
        Else
            GridDescontos.TextMatrix(GridDescontos.Row, iGrid_ValorDesconto_Col) = ""
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridDescontos.Row - GridDescontos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31322

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = Err

    Select Case Err

        Case 31319
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", Err, TipoDesconto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31322

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160819)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_Data(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = Data

    'Verifica se a Data está preenchida
    If Len(Trim(Data.Text)) > 0 Then

        'Faz a Crítica da Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 31323

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31324

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = Err

    Select Case Err

        Case 31323, 31324
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160820)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorDesconto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ValorDesconto

    Set objGridInt.objControle = ValorDesconto

    'Verifica se o Valor do Desconto foi preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then Error 31325

        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31352

    Saida_Celula_ValorDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorDesconto:

    Saida_Celula_ValorDesconto = Err

    Select Case Err

        Case 31325
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31352

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160821)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percentual(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = Percentual1

    If Len(Trim(Percentual1.Text)) > 0 Then

        'Verifica se o Percentual foi preenchido
        lErro = Porcentagem_Critica(Percentual1.Text)
        If lErro <> SUCESSO Then Error 31327

        'Formata o Percentual
        Percentual1.Text = Format(Percentual1.Text, "Fixed")

        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31328

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = Err

    Select Case Err

        Case 31327, 31328
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160822)

    End Select

    Exit Function

End Function


Public Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = Vendedor

    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then

        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Grid(Vendedor, objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then Error 31329

        If lErro = 25018 Then Error 31330

        If lErro = 25020 Then Error 31331

        'Verifica se GridComissoes foi preenchido
        If objGridComissoes.iLinhasExistentes > 0 Then

            'Loop no GridComissoes
            For iIndice = 1 To objGridComissoes.iLinhasExistentes

                'Verifica se Vendedor comparece em outra linha
                If iIndice <> GridComissoes.Row Then If GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido Then Error 31332

            Next

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31333

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = Err

    Select Case Err

        Case 31329

        Case 31330 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = Vendedor.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 31331 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = CDbl(Vendedor.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 31332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err, objVendedor.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160823)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula PercentualComissoes do grid que está deixando de ser o corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_PercentualComissao

    Set objGridInt.objControle = PercentualComissao

    'Verifica se o percentual está preenchido
    If Len(PercentualComissao.ClipText) > 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(PercentualComissao.Text)
        If lErro <> SUCESSO Then Error 31334

        dPercentual = CDbl(PercentualComissao.Text)

        'Mostra na tela o percentual formatado
        PercentualComissao.Text = Format(dPercentual, "Fixed")

        'Verifica se valorbase correspondente esta preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))

           'Calcula o valorcomissao
           dValorComissao = dPercentual * dValorBase / 100

           'Coloca o valorcomissoes na tela
           GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31335

    'Chama SomaPercentual
    lErro = Soma_Percentual()
    If lErro <> SUCESSO Then Error 31336

    'Chama SomaValor
    lErro = Soma_Valor()
    If lErro <> SUCESSO Then Error 31337

    Saida_Celula_PercentualComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualComissao:

    Saida_Celula_PercentualComissao = Err

    Select Case Err

        Case 31334, 31335
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31336, 31337

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160824)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorBase(objGridInt As AdmGrid) As Long
'Faz a crítica da celula ValorBase do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lTamanho As Long

On Error GoTo Erro_Saida_Celula_ValorBase

    Set objGridInt.objControle = ValorBase

    'Verifica se valor base está preenchido
    If Len(ValorBase.ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(ValorBase.Text)
        If lErro <> SUCESSO Then Error 31338

        dValorBase = CDbl(ValorBase.Text)

        'Mostra na tela o ValorBase formatado
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col) = Format(dValorBase, "Fixed")

        'Verifica se percentual comissao está preenchido
        lTamanho = Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col)))
        If lTamanho > 0 Then

            dPercentual = CDbl(left(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col), lTamanho - 1))

            'Calcula o valor da comissao
            dValorComissao = dPercentual * dValorBase / 100

            'Mostra na tela o valor da comissao
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Fixed")

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31339

    'Chama SomaValor
    lErro = Soma_Valor()
    If lErro <> SUCESSO Then Error 31340

    Saida_Celula_ValorBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBase:

    Saida_Celula_ValorBase = Err

    Select Case Err

        Case 31338, 31339
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31340

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160825)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lTamanho As Long

On Error GoTo Erro_Saida_Celula_ValorComissao

    Set objGridInt.objControle = ValorComissao

    'Verifica se valor está preenchido
    If Len(ValorComissao.ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(ValorComissao.Text)
        If lErro <> SUCESSO Then Error 31341

        dValorComissao = CDbl(ValorComissao.Text)

        'Mostra na tela o Valor
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Fixed")

        'Verifica se valor base correspondente está preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))

            'Verifica se percentual comissao correspondente está preenchido
            lTamanho = Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col)))
            If lTamanho > 0 Then

                dPercentual = CDbl(left(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col), lTamanho - 1))

                If (dPercentual * dValorBase) <> dValorComissao Then
                    
                    dPercentual = (dValorComissao) / dValorBase

                    'Mostra Valor Base da comissao na tela
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col) = Format(dPercentual, "Percent")

                End If

            End If

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 31342

    'Chama SomaPercentual
    lErro = Soma_Percentual()
    If lErro <> SUCESSO Then Error 31343

    'Chama SomaValor
    lErro = Soma_Valor()
    If lErro <> SUCESSO Then Error 31344

    Saida_Celula_ValorComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorComissao:

    Saida_Celula_ValorComissao = Err

    Select Case Err

        Case 31341, 31342
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31343, 31344

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160826)

    End Select

    Exit Function

End Function

Private Function Soma_Percentual() As Long
'Faz a soma da coluna de Percentual de objGridComissoes

Dim iIndice As Integer
Dim dSomaPercentual As Double
Dim lTamanho As Long

    dSomaPercentual = 0

    'Loop no Grid
    For iIndice = 1 To objGridComissoes.iLinhasExistentes

        'Verifica se Percentual da Comissão está preenchido
        lTamanho = Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_PercentualComissao_Col)))
        If lTamanho > 0 Then

            'Acumula Percentual em dSomaPercentual
            dSomaPercentual = dSomaPercentual + CDbl(left(GridComissoes.TextMatrix(iIndice, iGrid_PercentualComissao_Col), lTamanho - 1))

        End If

    Next

    'Mostra na tela o Total Percentual
    TotalPercentualComissao.Caption = Format((dSomaPercentual / 100), "Percent")

    Soma_Percentual = SUCESSO

    Exit Function

End Function


Private Function Soma_Valor() As Long

Dim iIndice As Integer
Dim dSomaValor As Double

    dSomaValor = 0

    'Loop no GridComissao
    For iIndice = 1 To objGridComissoes.iLinhasExistentes

        'Verifica se Valor da Comissão está preenchido
        If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_ValorComissao_Col))) > 0 Then

            'Acumula Valor em dSomaValor
            dSomaValor = dSomaValor + CDbl(GridComissoes.TextMatrix(iIndice, iGrid_ValorComissao_Col))

        End If
    Next

    'Mostra na tela o Total Valor
    TotalValorComissao.Caption = Format(dSomaValor, "Standard")

    Soma_Valor = SUCESSO

    Exit Function

End Function

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_EnterCell()

    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Private Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

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

Private Sub GridDescontos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDesconto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDesconto, iAlterado)
    End If

End Sub

Private Sub GridDescontos_GotFocus()

    Call Grid_Recebe_Foco(objGridDesconto)

End Sub

Private Sub GridDescontos_EnterCell()

    Call Grid_Entrada_Celula(objGridDesconto, iAlterado)

End Sub

Private Sub GridDescontos_LeaveCell()

    Call Saida_Celula(objGridDesconto)

End Sub

Private Sub GridDescontos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDesconto)

End Sub

Private Sub GridDescontos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDesconto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDesconto, iAlterado)
    End If

End Sub

Private Sub GridDescontos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridDesconto)
    
End Sub

Private Sub GridDescontos_RowColChange()

    Call Grid_RowColChange(objGridDesconto)

End Sub

Private Sub GridDescontos_Scroll()

    Call Grid_Scroll(objGridDesconto)

End Sub

Private Sub GridComissoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Private Sub GridComissoes_GotFocus()

    Call Grid_Recebe_Foco(objGridComissoes)

End Sub

Private Sub GridComissoes_EnterCell()

    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)

End Sub

Private Sub GridComissoes_LeaveCell()

    Call Saida_Celula(objGridComissoes)

End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)

End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Private Sub GridComissoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoes)
    
End Sub

Private Sub GridComissoes_RowColChange()

    Call Grid_RowColChange(objGridComissoes)

End Sub

Private Sub GridComissoes_Scroll()

    Call Grid_Scroll(objGridComissoes)

End Sub

Private Sub GridNFiscal_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNFiscal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscal, iAlterado)
    End If

End Sub

Private Sub GridNFiscal_GotFocus()

    Call Grid_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub GridNFiscal_EnterCell()

    Call Grid_Entrada_Celula(objGridNFiscal, iAlterado)

End Sub

Private Sub GridNFiscal_LeaveCell()

    Call Saida_Celula(objGridNFiscal)

End Sub

Private Sub GridNFiscal_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridNFiscal)

'    Call GeraFatura_Click

End Sub

Private Sub GridNFiscal_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNFiscal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscal, iAlterado)
    End If

    Call GeraFatura_Click

End Sub

Private Sub GridNFiscal_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridNFiscal)
    
End Sub

Private Sub GridNFiscal_RowColChange()

    Call Grid_RowColChange(objGridNFiscal)

End Sub

Private Sub GridNFiscal_Scroll()

    Call Grid_Scroll(objGridNFiscal)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39762

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then Error 31345

            'Se for o GridDescontos
            Case GridDescontos.Name

                lErro = Saida_Celula_GridDescontos(objGridInt)
                If lErro <> SUCESSO Then Error 31346

            'se for o GridComissoes
            Case GridComissoes.Name

                lErro = Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then Error 31347

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 31348

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 31345, 31346, 31347, 31348, 39762

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160827)

    End Select

    Exit Function

End Function

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

Private Sub GeraFatura_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub GeraFatura_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub GeraFatura_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = GeraFatura
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Serie1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub Serie1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub Serie1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = Serie1
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = Numero
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub Filial_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = Filial
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub DataEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = DataEmissao
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub SiglaDoc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SiglaDoc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub SiglaDoc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub SiglaDoc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = SiglaDoc
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNFiscal)

End Sub

Private Sub ValorTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscal.objControle = ValorTotal
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

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

Private Sub TipoDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Private Sub TipoDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Private Sub TipoDesconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDesconto.objControle = TipoDesconto
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Private Sub Data_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDesconto.objControle = Data
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDesconto.objControle = ValorDesconto
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Percentual1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Percentual1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Private Sub Percentual1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Private Sub Percentual1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDesconto.objControle = Percentual1
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PercentualComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentualComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub PercentualComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub PercentualComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub ValorComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Vendedores_DblClick()

Dim lPosicaoSeparador As Long
Dim sVendedor As String
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Vendedores_DblClick

    'Se a célula do Grid for a de vendedores
    If GridComissoes.Col = iGrid_Vendedor_Col Then
        
        'Coloca no Grid o Vendedor selecionado
        lPosicaoSeparador = InStr(Vendedores.Text, SEPARADOR)
        sVendedor = Mid(Vendedores.Text, lPosicaoSeparador + 1)
        
        'Verifica se GridComissoes foi preenchido
        If objGridComissoes.iLinhasExistentes > 0 Then

            'Loop no GridComissoes
            For iIndice = 1 To objGridComissoes.iLinhasExistentes

                'Verifica se Vendedor comparece em outra linha
                If iIndice <> GridComissoes.Row Then If GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col) = sVendedor Then gError 64457

            Next

        End If

        
        'Coloca no Grid o Vendedor selecionado
        GridComissoes.TextMatrix(GridComissoes.Row, GridComissoes.Col) = sVendedor
        Vendedor.Text = sVendedor

    End If

    Exit Sub
    
Erro_Vendedores_DblClick:
    
    Select Case gErr

        Case 64457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", gErr, sVendedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160828)

    End Select

    Exit Sub
    
End Sub

Private Function Trata_TabDescontosComissoes() As Long
'Faz o tratamento necessário no caso de o Tab DescontosComissoes ser selecionado

Dim lErro As Long
Dim iIndice As Integer
Dim objDescontoPadrao As New ClassDescontoPadrao
Dim colDescontoPadrao As Collection
Dim colDesconto As colDesconto
Dim colComissao As colComissao
Dim objVendedor As New ClassVendedor
Dim objcliente As New ClassCliente
Dim objTipoCliente As New ClassTipoCliente
Dim objFilialCliente As New ClassFilialCliente
Dim iParcela As Integer
Dim dValor As Double
Dim dtData As Date
Dim dtDataReal As Date
Dim colComissaoVendedorPercentual As New Collection

On Error GoTo Erro_Trata_TabDescontosComissoes

    'Verifica se existe alguma parcela no Tab
    If Len(Trim(Parcela.Caption)) <= 0 Then

        'Traz a Unica Parcela
        If objGridParcelas.iLinhasExistentes > 0 Then iParcela = 1

        Call Traz_Parcela_Tela(iParcela)

    End If

    'Se não existirem Parcelas então limpa-se o Tab DescontosComissoes
    If objGridParcelas.iLinhasExistentes = 0 Then Limpa_Tab_DescontosComissoes

    Trata_TabDescontosComissoes = SUCESSO

    Exit Function

Erro_Trata_TabDescontosComissoes:

    Trata_TabDescontosComissoes = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160829)

    End Select

    Exit Function

End Function

Function NFiscal_Monta_ColecaoComissao(colComissaoTotal As Collection) As Long
'Monta uma Comissao de Vendedor e Percentual, com todos os Vendedores das Comissoes das Notais

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objVendedorPercentualBaixa As AdmCodigoValor
Dim dPercentualBaixaVendedor As Double
Dim iEncontrou As Integer

On Error GoTo Erro_NFiscal_Monta_ColecaoComissao
    
    'Le as Comissoes para cada NotaFiscal Selecionada
    For iIndice = 1 To objGridNFiscal.iLinhasExistentes
        
        If GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col) = S_MARCADO Then
            
            Set objNFiscal = New ClassNFiscal
            
            objNFiscal.sSerie = GridNFiscal.TextMatrix(iIndice, iGrid_Serie_Col)
            objNFiscal.lNumNotaFiscal = CLng(GridNFiscal.TextMatrix(iIndice, iGrid_Numero_Col))
            objNFiscal.iFilialEmpresa = giFilialEmpresa
            objNFiscal.dtDataEmissao = CDate(GridNFiscal.TextMatrix(iIndice, iGrid_DataEmissao_Col))
            
            'TEM QUE LER A PROPRIA NF
            lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
            If lErro <> SUCESSO And lErro <> 43676 Then gError 214964
        
            'Se não encontrou a Nota Fiscal --> Erro
            If lErro <> SUCESSO Then gError 214965
            
            'Lê todas as Comissões da Nota
            lErro = CF("NFiscal_Le_Comissoes", objNFiscal)
            If lErro <> SUCESSO And lErro <> 21386 Then gError 58408
            
            For iIndice2 = 1 To objNFiscal.ColComissoesNF.Count
            
                dPercentualBaixaVendedor = (objNFiscal.ColComissoesNF(iIndice2).dValor - objNFiscal.ColComissoesNF(iIndice2).dValorEmissao) / CDbl(TotalNotasSel)
          
                iEncontrou = 0
                                            
                'Verifica se o Vendedor já está na coleção
                For Each objVendedorPercentualBaixa In colComissaoTotal
                    
                    'Se estiver atualiza o Percentual
                    If objVendedorPercentualBaixa.iCodigo = objNFiscal.ColComissoesNF.Item(iIndice2).iCodVendedor Then
                            
                        objVendedorPercentualBaixa.dValor = objVendedorPercentualBaixa.dValor + dPercentualBaixaVendedor
                        
                        iEncontrou = 1
                        
                    End If
                    
                Next
                                
                'Se não estiver na coleção -- adiciona
                If iEncontrou = 0 Then
                    
                    Set objVendedorPercentualBaixa = New AdmCodigoValor
                    
                    objVendedorPercentualBaixa.iCodigo = objNFiscal.ColComissoesNF(iIndice2).iCodVendedor
                    objVendedorPercentualBaixa.dValor = dPercentualBaixaVendedor
                    
                    colComissaoTotal.Add objVendedorPercentualBaixa
                
                End If
            
            Next
                         
        End If
    
    Next
   
    'Arredonda o Percentual em 2 casas decimais
    For Each objVendedorPercentualBaixa In colComissaoTotal
        
        objVendedorPercentualBaixa.dValor = Round(objVendedorPercentualBaixa.dValor, 2)
        
    Next
                                 
    NFiscal_Monta_ColecaoComissao = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Monta_ColecaoComissao:
    
    NFiscal_Monta_ColecaoComissao = gErr
    
    Select Case gErr
        
        Case 58408, 214964 'Tratado na Rotina Chamada
        
        Case 214965
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160830)

    End Select

    Exit Function
        
End Function

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentes As Integer
Dim iLinhaAtual As Integer

    'Guarda a linha atual e o número de linhas existentes
    iLinhasExistentes = objGridParcelas.iLinhasExistentes
    iLinhaAtual = GridParcelas.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

    'Se a linha foi apagada
    If (objGridParcelas.iLinhasExistentes < iLinhasExistentes) Then
        'Retira-se das coleções os dados correspondentes a parcela que foi excluída
        If colcolDesconto.Count >= iLinhaAtual Then colcolDesconto.Remove (iLinhaAtual)
        If colcolComissao.Count >= iLinhaAtual Then colcolComissao.Remove (iLinhaAtual)
    End If

End Sub

Private Function Traz_Parcela_Tela(iParcela As Integer) As Long
'Traz para a Tela os dados da Parcela passada

Dim objDesconto As ClassDesconto
Dim colDescontos As colDesconto
Dim colComissao As colComissao
Dim objComissao As ClassComissao
Dim iLinha As Integer
Dim iIndice As Integer
Dim lSeparador As Long
Dim lErro As Long

On Error GoTo Erro_Traz_Parcela_Tela

    If iParcela = 0 Then Exit Function

    'Põe a Parcela na tela
    Parcela.Caption = iParcela
    
    'Se a coleção de Descontos estiver vazia --> Não faz nada
    If colcolDesconto.Count >= iParcela Then
    
        'Pega a coleção de descontos correspondente a Parcela
        Set colDescontos = colcolDesconto(iParcela).colDesconto
    
        iLinha = 0
    
        'Preenche o GridDesconto com os Descontos existentes em colDesconto
        For Each objDesconto In colDescontos
    
            iLinha = iLinha + 1
            'Coloca no grid a data do desconto
            If objDesconto.dtData <> DATA_NULA Then GridDescontos.TextMatrix(iLinha, iGrid_DataDesconto_Col) = Format(objDesconto.dtData, "dd/mm/yyyy")
            'Coloca no grid tipo de desconto
            For iIndice = 0 To TipoDesconto.ListCount - 1
                If TipoDesconto.ItemData(iIndice) = objDesconto.iCodigo Then
                    GridDescontos.TextMatrix(iLinha, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
                    Exit For
                End If
            Next
            
            'Coloca no grid o valor do Desconto
            If objDesconto.iCodigo = VALOR_FIXO Or objDesconto.iCodigo = VALOR_ANT_DIA Or objDesconto.iCodigo = VALOR_ANT_DIA_UTIL Then
                GridDescontos.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = Format(objDesconto.dValor, "Standard")
            ElseIf objDesconto.iCodigo = Percentual Or objDesconto.iCodigo = PERC_ANT_DIA Or objDesconto.iCodigo = PERC_ANT_DIA_UTIL Then
                GridDescontos.TextMatrix(iLinha, iGrid_PercentualDesconto_Col) = Format(objDesconto.dValor, "Percent")
            End If
    
        Next
    
        'Atribui o número de linhas existentes
        objGridDesconto.iLinhasExistentes = iLinha
    
    End If
    
    'Se a coleção de Comissões estiver vazia --> Não faz nada
    If colcolComissao.Count >= iParcela Then
        'Pega a coleção de Comissões correspondente a Parcela
        Set colComissao = colcolComissao(iParcela).colComissao
    
        iLinha = 0
    
        'Preenche o GridComissoes com as Comissões existentes em colComissao
        For Each objComissao In colComissao
    
            iLinha = iLinha + 1
            GridComissoes.TextMatrix(iLinha, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
            GridComissoes.TextMatrix(iLinha, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
            GridComissoes.TextMatrix(iLinha, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")
            For iIndice = 0 To Vendedores.ListCount - 1
                If Vendedores.ItemData(iIndice) = objComissao.iCodVendedor Then
                    lSeparador = InStr(Vendedores.List(iIndice), SEPARADOR)
                    GridComissoes.TextMatrix(iLinha, iGrid_Vendedor_Col) = Mid(Vendedores.List(iIndice), lSeparador + 1)
                    Exit For
                End If
            Next
        Next
    
        'Atualiza o número de linhas existentes
        objGridComissoes.iLinhasExistentes = iLinha
    
        'Faz a Soma das colunas Percentual e Valor do Grid de Comissões
        Call Soma_Percentual
        Call Soma_Valor
    
    End If
    
    Traz_Parcela_Tela = SUCESSO

    Exit Function

Erro_Traz_Parcela_Tela:

    Traz_Parcela_Tela = Err

    Select Case Err

        Case 39761

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160831)

    End Select

    Exit Function

End Function

Sub Limpa_Tab_Cobranca()

    CondicaoPagamento.ListIndex = -1
    CondicaoPagamento.Text = ""
    
    Call Grid_Limpa(objGridParcelas)
    
End Sub

Sub Limpa_Tab_DescontosComissoes()
'Limpa o TabDescontosComissoes

    'Limpa a parcela
    Parcela.Caption = " "

    'Limpa os Grids de Desconto e Comissões
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridDesconto)

    Exit Sub

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data de emissao em um dia
    lErro = Data_Up_Down_Click(EmissaoFatura, DIMINUI_DATA)
    If lErro Then Error 31356

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 31356

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160832)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a Data de Emissão em um dia
    lErro = Data_Up_Down_Click(EmissaoFatura, AUMENTA_DATA)
    If lErro Then Error 31357

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 31357

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160833)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoFatura_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoFatura_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(EmissaoFatura.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(EmissaoFatura.Text)
    If lErro <> SUCESSO Then Error 31358

    Exit Sub

Erro_EmissaoFatura_Validate:

    Cancel = True


    Select Case Err

        Case 31358

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160834)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagamento_Click()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dValorNFsSelecionadas As Double

On Error GoTo Erro_CondicaoPagamento_Click

    'Verifica se alguma Condição foi selecionada
    If CondicaoPagamento.ListIndex = -1 Then Exit Sub

    'Passa o código da Condição para objCondicaoPagto
    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then Error 31359

    'se a condição não está cadastrada ==> erro
    If lErro = 19205 Then Error 31361

    'Limpa o Grid Parcelas
    Call Grid_Limpa(objGridParcelas)

    Call Calcula_Valor_Notas_Selecionadas(dValorNFsSelecionadas)
    
    'Testa se EmissaoFatura está preenchida
    If Len(Trim(EmissaoFatura.ClipText)) > 0 And dValorNFsSelecionadas > 0 Then

        'Preenche o GridParcelas
        lErro = GridParcelas_Preenche(objCondicaoPagto, dValorNFsSelecionadas)
        If lErro <> SUCESSO Then Error 31360

    End If
    
    lErro = Calcula_Comissoes()
    If lErro <> SUCESSO Then Error 58439
    
    lErro = Calcula_Descontos
    If lErro <> SUCESSO Then Error 58440
    
    Parcela.Caption = ""
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case Err

        Case 31359, 31360, 58439, 58440 'Tratados na Rotinas chamadas

        Case 31361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160835)

      End Select

    Exit Sub

End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dValorNFsSelecionadas As Double

On Error GoTo Erro_Condicaopagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then
        Call Grid_Limpa(objGridParcelas)
        Exit Sub
    End If
    
    'Verifica se é uma Condicaopagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 31362

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 31363
        If lErro = 19205 Then Error 31365

        'Testa se pode ser usada em Contas a Pagar
        If objCondicaoPagto.iEmPagamento = 0 Then Error 31366

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)
        
        'Calcula o Total do Valor das Notas Selecionadas
        Call Calcula_Valor_Notas_Selecionadas(dValorNFsSelecionadas)
        
        'Se ValorTotal e EmissaoFatura estiverem preenchidos, preenche GridParcelas
        If Len(Trim(dValorNFsSelecionadas)) > 0 And Len(Trim(EmissaoFatura.ClipText)) > 0 Then

                'Limpa o Grid Parcelas
                Call Grid_Limpa(objGridParcelas)

                'Preenche o GridParcelas
                lErro = GridParcelas_Preenche(objCondicaoPagto, dValorNFsSelecionadas)
                If lErro <> SUCESSO Then Error 31364

        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then Error 31367
    
    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True


    Select Case Err

       Case 31362, 31363, 31364

       Case 31365
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            Else
            End If

        Case 31366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", Err, objCondicaoPagto.iCodigo)

        Case 31367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagamento.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160836)

    End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto, dValorNFsSelecionadas As Double) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim dtDataVenctoReal As Date
Dim dValorIRRF As Double, dValorINSS As Double
Dim dValorISSRetido As Double
Dim iIndice As Integer

On Error GoTo Erro_GridParcelas_Preenche

    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
    If Len(Trim(INSSValor.Text)) > 0 Then dValorINSS = CDbl(INSSValor)

    If ISSRetido.Value = 1 Then
        If Len(Trim((ISSValor.Caption))) > 0 Then dValorISSRetido = CDbl(ISSValor.Caption)
    End If
    
    'Valor a Pagar
    dValorPagar = dValorNFsSelecionadas - dValorIRRF - dValorISSRetido - dValorINSS

    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        objCondicaoPagto.dValorTotal = dValorPagar
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then Error 31368

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(EmissaoFatura.ClipText)) > 0 Then

        objCondicaoPagto.dtDataEmissao = CDate(EmissaoFatura.Text)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then Error 31369

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then Error 31370

            'Coloca Data de Vencimento Real no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        Next

    End If

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = Err

    Select Case Err

        Case 31368, 31369, 31370 'Tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160837)

    End Select

End Function

Private Function Move_TabSelecao_Memoria(objGeracaoFatura As ClassGeracaoFatura) As Long
'Move os dados do Tab de Selecao para que seja Preenchida as Notas

Dim objcliente As New ClassCliente, lErro As Long

On Error GoTo Erro_Move_TabSelecao_Memoria

    If Len(Trim(CodCliente.ClipText)) > 0 Then

        objcliente.sNomeReduzido = CodCliente.Text

        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 41527

        'Não encontrou Cliente --> erro
        If lErro = 12348 Then Error 41528

        objGeracaoFatura.lCliente = objcliente.lCodigo

    End If
    
    'Filial Cliente
    If Len(Trim(FilialCliente.Text)) > 0 And optTodasFiliais.Value = vbUnchecked Then
        objGeracaoFatura.iFilialCli = Codigo_Extrai(FilialCliente.Text)
    Else
        objGeracaoFatura.iFilialCli = 0
    End If
    
    'Série
    objGeracaoFatura.sSerie = Serie.Text
    
    'Nota Fiscal Inicial
    If Len(Trim(NFiscalInicial.ClipText)) > 0 Then objGeracaoFatura.lNumeroNFDe = CLng(NFiscalInicial.Text)
    
    'Nota Fiscal Final
    If Len(Trim(NFiscalFinal.ClipText)) > 0 Then objGeracaoFatura.lNumeroNFAte = CLng(NFiscalFinal.Text)
    
    'Data De
    If Len(Trim(DataDe.ClipText)) > 0 Then
        objGeracaoFatura.dtEmissaoNFDe = CDate(DataDe.Text)
    Else
        objGeracaoFatura.dtEmissaoNFDe = DATA_NULA
    End If

    'Data Ate
    If Len(Trim(DataAte.ClipText)) > 0 Then
        objGeracaoFatura.dtEmissaoNFAte = CDate(DataAte.Text)
    Else
        objGeracaoFatura.dtEmissaoNFAte = DATA_NULA
    End If
    
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = Err

    Select Case Err

        Case 41527

        Case 41528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, CodCliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160838)

    End Select

    Exit Function

End Function

Private Function Trata_TabTributacao() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iClick As Integer
Dim dSomaISS As Integer
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Trata_TabTributacao

    dSomaISS = 0
    iIndice = 0

    'Para cada nota fiscal na coleção de Notas Fiscais
    For iIndice = 1 To objGridNFiscal.iLinhasExistentes
        iClick = 0
        'Verifica se ela está selecionada no Grid
        If Len(Trim(GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col))) > 0 Then iClick = CInt(GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col))
        'Atualiza o Valor
        If iClick = 1 Then
            
            Set objNFiscal.objTributacaoNF = New ClassTributacaoDoc
            
            objNFiscal.lNumNotaFiscal = GridNFiscal.TextMatrix(iIndice, iGrid_Numero_Col)
            objNFiscal.sSerie = GridNFiscal.TextMatrix(iIndice, iGrid_Serie_Col)
            objNFiscal.iFilialEmpresa = giFilialEmpresa
    
            'Lê o NumIntDoc da NFiscal
            lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
            If lErro <> SUCESSO And lErro <> 22867 Then Error 58409

            dSomaISS = dSomaISS + objNFiscal.objTributacaoNF.dISSValor
        
        End If
    Next

    ISSValor.Caption = Format(dSomaISS, "Fixed")

    Trata_TabTributacao = SUCESSO

    Exit Function

Erro_Trata_TabTributacao:
    
    Trata_TabTributacao = Err
    
    Select Case Err
        
        Case 58409 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160839)

    End Select

    Exit Function

End Function

Private Sub UpDownParcela_DownClick()

Dim iIndice As Integer
Dim lErro As Long
Dim iParcela As Integer

On Error GoTo Erro_UpDownParcela_DownClick

    BotaoGerar.SetFocus
    DoEvents
    
    'Verifica se existem Parcelas no Grid
    If objGridParcelas.iLinhasExistentes = 0 Then
        Parcela.Caption = ""
        Exit Sub
    End If

    'Verifica se já existe alguma Parcela na Tela
    If Len(Trim(Parcela.Caption)) > 0 Then
        iParcela = CInt(Parcela.Caption)
    Else
        iParcela = 0
    End If

    'Verifica se existe uma Parcela inferior a Parcela da Tela
    If iParcela - 1 > 0 Then

        'Recolhe os dados da Parcela da Tela
        lErro = Recolhe_Parcela_Tela(iParcela)
        If lErro <> SUCESSO Then Error 31371

        'Limpa o TabDescontosComissoes
        Call Limpa_Tab_DescontosComissoes

        'Coloca na Tela os dados da nova parcela a tratar
        lErro = Traz_Parcela_Tela(iParcela - 1)
        If lErro <> SUCESSO Then Error 31372

    End If

    Exit Sub

Erro_UpDownParcela_DownClick:

    Select Case Err

        Case 31371, 31372

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160840)

    End Select

    Exit Sub

End Sub


Private Sub UpDownParcela_UpClick()

Dim iIndice As Integer
Dim lErro As Long
Dim iParcela As Integer

On Error GoTo Erro_UpDownParcela_UpClick

    BotaoGerar.SetFocus

    DoEvents

    'Verifica se existe alguma parcela no GridParcelas
    If objGridParcelas.iLinhasExistentes = 0 Then
        Parcela.Caption = ""
        Exit Sub
    End If

    'Verifica se já existe alguma Parcela na tela
    If Len(Trim(Parcela.Caption)) > 0 Then
        iParcela = CInt(Parcela.Caption)
    Else
        iParcela = 0
    End If

    'Verifica se existe a Parcela (iParcela+1)
    If iParcela + 1 <= objGridParcelas.iLinhasExistentes Then

        If iParcela <> 0 Then
            'Recolhe os dados da parcela que está na tela
            lErro = Recolhe_Parcela_Tela(iParcela)
            If lErro <> SUCESSO Then Error 31373
        End If

        'Limpa o TabDescontosComissoes
        Call Limpa_Tab_DescontosComissoes

        'Traz para telas os dados da Parcela seguinte
        lErro = Traz_Parcela_Tela(iParcela + 1)
        If lErro <> SUCESSO Then Error 31374

    End If

    Exit Sub

Erro_UpDownParcela_UpClick:

    Select Case Err

        Case 31373, 31374

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160841)

    End Select

    Exit Sub

End Sub

Private Function Recolhe_Parcela_Tela(iParcela As Integer) As Long
'Recolhe da Tela os dados da Parcela

Dim iIndice As Integer
Dim lErro As Long
Dim objDesconto As New ClassDesconto
Dim colDesconto As New colDesconto
Dim objComissao As New ClassComissao
Dim colComissao As New colComissao
Dim objVendedor As New ClassVendedor
Dim lTamanho As Long

    If iParcela > objGridParcelas.iLinhasExistentes Then Exit Function

    'Verifica se existem descontos no GridDescontos
    If objGridDesconto.iLinhasExistentes > 0 Then

        'Loop de armazenamento dos dados em colDesconto
        For iIndice = 1 To objGridDesconto.iLinhasExistentes

            Set objDesconto = New ClassDesconto

            'Recolhe os Dados do Grid
            If Len(Trim(GridDescontos.TextMatrix(iIndice, 1))) > 0 Then objDesconto.iCodigo = Codigo_Extrai(GridDescontos.TextMatrix(iIndice, 1))

            If objDesconto.iCodigo = VALOR_FIXO Or objDesconto.iCodigo = VALOR_ANT_DIA Or objDesconto.iCodigo = VALOR_ANT_DIA_UTIL Then
                If Len(Trim(GridDescontos.TextMatrix(iIndice, 3))) > 0 Then objDesconto.dValor = CDbl(GridDescontos.TextMatrix(iIndice, 3))
            ElseIf objDesconto.iCodigo = Percentual Or objDesconto.iCodigo = PERC_ANT_DIA Or objDesconto.iCodigo = PERC_ANT_DIA_UTIL Then
                lTamanho = Len(Trim(GridDescontos.TextMatrix(iIndice, 4)))
                If lTamanho > 0 Then objDesconto.dValor = PercentParaDbl(GridDescontos.TextMatrix(iIndice, iGrid_PercentualDesconto_Col))
            End If

            If Len(Trim(GridDescontos.TextMatrix(iIndice, 2))) > 0 Then objDesconto.dtData = CDate(GridDescontos.TextMatrix(iIndice, 2))

            ' Adiciona em colDesconto
            colDesconto.Add objDesconto.iCodigo, objDesconto.dtData, objDesconto.dValor
        Next

    End If

    'Guarda em colcolDesconto
    If colcolDesconto.Count >= iParcela Then
        Set colcolDesconto(iParcela).colDesconto = colDesconto
    Else
        colcolDesconto.Add colDesconto
    End If
    
    'Verifica se existem comissões no GridComissoes
    If objGridComissoes.iLinhasExistentes > 0 Then

        'Loop de armazenamento do dados em colComissao
        For iIndice = 1 To objGridComissoes.iLinhasExistentes

            Set objComissao = New ClassComissao

            'Recolhe os Dados do GridComissao

            lTamanho = Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_PercentualComissao_Col)))
            If lTamanho > 0 Then objComissao.dPercentual = PercentParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_PercentualComissao_Col))
            If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_ValorBase_Col))) > 0 Then objComissao.dValorBase = CDbl(GridComissoes.TextMatrix(iIndice, iGrid_ValorBase_Col))
            If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_ValorComissao_Col))) > 0 Then objComissao.dValor = CDbl(GridComissoes.TextMatrix(iIndice, iGrid_ValorComissao_Col))
            If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col))) > 0 Then
                objVendedor.sNomeReduzido = GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)
                lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
                If lErro = SUCESSO Then objComissao.iCodVendedor = objVendedor.iCodigo
            End If

            'Adiciona em colComissao
            colComissao.Add 0, giFilialEmpresa, 0, 0, 0, objComissao.iCodVendedor, DATA_NULA, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor, DATA_NULA

        Next
    
        'Guarda em colcolComissao
        If colcolComissao.Count >= iParcela Then
            Set colcolComissao(iParcela).colComissao = colComissao
        Else
            colcolComissao.Add colComissao
        End If

    End If

    Recolhe_Parcela_Tela = SUCESSO

    Exit Function

End Function

Private Function Valida_GridParcelas() As Long

Dim iIndice As Integer
Dim dSomaParcelas As Double
Dim dtDataVencimento As Date
Dim dValorTitulo As Double
Dim lErro As Long
Dim dValorIRRF As Double, dValorINSS As Double
Dim dValorISSRetido As Double
Dim dValorNFsSelecionadas As Double

On Error GoTo Erro_Valida_GridParcelas

    dSomaParcelas = 0

    'Loop no GridParcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        'Verifica se DataVencimento foi preenchida
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))) = 0 Then Error 31292
        dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))

        'Verifica se DataVencimento é maior ou igual a DataEmissao
        If Len(Trim(EmissaoFatura.ClipText)) > 0 Then
            If dtDataVencimento < CDate(EmissaoFatura) Then Error 31293
        End If

        'Verifica a ordenação das Datas de Vencimento das Parcelas
        If iIndice > 1 Then If dtDataVencimento < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_col)) Then Error 31294

        'Verifica se Valor da Parcela foi preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))) = 0 Then Error 26190

        'Verifica se Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        If lErro <> SUCESSO Then Error 31295

        'Acumula Valor Parcela em dSomaParcelas
        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
    Next

    'Valor a Pagar
    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
    If Len(Trim(INSSValor.Text)) > 0 Then dValorINSS = CDbl(INSSValor)
    If Len(Trim(ISSValor.Caption)) > 0 Then dValorISSRetido = CDbl(IIf(ISSRetido.Value = 1, ISSValor.Caption, 0))
    
    Call Calcula_Valor_Notas_Selecionadas(dValorNFsSelecionadas)
    
    dValorTitulo = dValorNFsSelecionadas - dValorIRRF - dValorISSRetido - dValorINSS
    If dValorTitulo <= 0 Then Error 31296

    'comparar a soma das parcelas com a soma das nfs menos impostos retidos
    If Format(dValorTitulo, "0,00") <> Format(dSomaParcelas, "0,00") Then Error 31297

    Valida_GridParcelas = SUCESSO

    Exit Function

Erro_Valida_GridParcelas:

    Valida_GridParcelas = Err

    Select Case Err

        Case 26190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_NAO_INFORMADA", gErr, iIndice)

        Case 31292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", Err, iIndice)

        Case 31293
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, dtDataVencimento, EmissaoFatura, iIndice)

        Case 31294
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_NAO_ORDENADA", Err)

        Case 31295

        Case 31296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTITULO_MENOS_IMPOSTOS", Err)

        Case 31297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_DIFERENTE", Err, dSomaParcelas, dValorTitulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160842)

    End Select

    Exit Function

End Function

Private Function Valida_Dados_Parcelas() As Long
'Valida os Dados da Parcela para a Gravação

Dim iIndice As Integer
Dim objComissao As ClassComissao
Dim objDesconto As ClassDesconto
Dim dDesconto As Double
Dim colDesconto As colDesconto
Dim colComissao As colComissao
Dim lErro As Long
Dim iIndice2 As Integer

On Error GoTo Erro_Valida_Dados_Parcelas

    'Para cada Parcela existente
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        'Se a coleção de Comissao estiver vazia --> Não faz nada
        If colcolComissao.Count = 0 Then
        
            Set colComissao = colcolComissao(iIndice).colComissao
    
            iIndice2 = 0
    
            'Loop de validação os dados das comissões
            For Each objComissao In colComissao
    
                iIndice2 = iIndice2 + 1
                If objComissao.iCodVendedor = 0 Then Error 31375
                If objComissao.dPercentual <= 0 Then Error 31376
                If objComissao.dValorBase <= 0 Then Error 31377
                If objComissao.dValor <= 0 Then Error 31378
    
            Next
        
        End If
        
        'Se a coleção de Desconto estiver vazia --> Não faz nada
        If colcolDesconto.Count = 0 Then
        
            Set colDesconto = colcolDesconto(iIndice).colDesconto
    
            iIndice2 = 0
            
            'Loop de validação dos descontos
            For Each objDesconto In colDesconto
    
                iIndice2 = iIndice2 + 1
                If objDesconto.dtData = DATA_NULA Then Error 31379
                If objDesconto.dValor <= 0 Then Error 31380
                If objDesconto.iCodigo <= 0 Then Error 26202
    
            Next
        
        End If
        
    Next

    Valida_Dados_Parcelas = SUCESSO

    Exit Function

Erro_Valida_Dados_Parcelas:

    Valida_Dados_Parcelas = Err

    Select Case Err

        Case 31375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)

        Case 31376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)

        Case 31377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)

        Case 31378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)

        Case 31379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", Err, iIndice2, iIndice)

        Case 31380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", Err, iIndice2, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160843)

    End Select

    Exit Function

End Function

Private Sub Atualiza_Colecoes()

Dim iIndice As Integer
Dim colComissao As colComissao
Dim colDesconto As colDesconto

    'Iguala o número de coleções de Comissões ao número de Parcelas
    If colcolComissao.Count < objGridParcelas.iLinhasExistentes Then
        For iIndice = colcolComissao.Count + 1 To objGridParcelas.iLinhasExistentes
            Set colComissao = New colComissao
            colcolComissao.Add colComissao
        Next
    ElseIf colcolComissao.Count > objGridParcelas.iLinhasExistentes Then
        Do While colcolComissao.Count > objGridParcelas.iLinhasExistentes
            colcolComissao.Remove (colcolComissao.Count)
        Loop
    End If

    'Iguala o número de coleções de Descontos ao número de Parcelas
    If colcolDesconto.Count < objGridParcelas.iLinhasExistentes Then
        For iIndice = colcolDesconto.Count + 1 To objGridParcelas.iLinhasExistentes
            Set colDesconto = New colDesconto
            colcolDesconto.Add colDesconto
        Next
    ElseIf colcolDesconto.Count > objGridParcelas.iLinhasExistentes Then
        Do While colcolDesconto.Count > objGridParcelas.iLinhasExistentes
            colcolDesconto.Remove (colcolDesconto.Count)
        Loop
    End If

    Exit Sub

End Sub

Private Sub Limpa_Tela_GeracaoFatura()
'Limpa todos os campos da Tela de Geração de Fatura

    'Chama o Limpa_Tela
    Call Limpa_Tela(Me)

    'Limpa os demais campos que não são limpos no Limpa_Tela
    FilialCliente.Clear

    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridDesconto)
    Call Grid_Limpa(objGridNFiscal)

    ISSRetido.Value = vbUnchecked
    ISSValor.Caption = ""
    Parcela.Caption = ""
    TotalPercentualComissao.Caption = ""
    TotalValorComissao = ""
    TotalNotasSel.Caption = ""
    CondicaoPagamento.Text = ""
    CondicaoPagamento.ListIndex = -1
    Serie.Text = ""
    Serie.ListIndex = -1
    If gobjFAT.iGeraFatTodasFiliais = MARCADO Then
        optTodasFiliais.Value = vbChecked
    Else
        optTodasFiliais.Value = vbUnchecked
    End If
    
    'Zera as coleções
    Set colcolDesconto = New colcolDesconto
    Set colcolComissao = New colcolComissao

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0

    Exit Sub

End Sub

'Início contabilidade
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

'****
Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

'****
Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

'****
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
'Traz o lote selecionado para a tela

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
'Mnemonico da Contabilidade

Dim lErro As Long
Dim iCodigoF As Integer
Dim sFilial As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case CLIENTE1
            If Len(CodCliente.Text) > 0 Then
                objMnemonicoValor.colValor.Add CodCliente.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case FILIAL1
            If Len(FilialCliente.Text) > 0 Then
                iCodigoF = Codigo_Extrai(FilialCliente.Text)
                sFilial = Mid(FilialCliente.Text, Len(iCodigoF) + 1)
                objMnemonicoValor.colValor.Add FilialCliente.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case FILIAL_COD
            If Len(FilialCliente.Text) > 0 Then
                iCodigoF = Codigo_Extrai(FilialCliente.Text)
                objMnemonicoValor.colValor.Add iCodigoF
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case Data1
            If Len(EmissaoFatura.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(EmissaoFatura.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case VALORFATURA
            If Len(TotalNotasSel.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(TotalNotasSel.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_IRRF
            If Len(ValorIRRF.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIRRF.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_INSS
            If Len(INSSValor.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(INSSValor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case Else
            Error 39763

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 39763
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160844)

    End Select

    Exit Function

End Function

Private Function Inicializa_ComissaoParcela(dValorBaseParcela As Double, colComissaoVendedorPercentual As Collection, colComissao As colComissao) As Long
'Inicializa as Comissões das Parcelas das Parcelas

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objVendedor As New ClassVendedor
Dim objcliente As New ClassCliente
Dim objTipoCliente As New ClassTipoCliente
Dim objComissao As New ClassComissao
Dim objComissaoVendedorPercentual As AdmCodigoValor

On Error GoTo Erro_Inicializa_ComissaoParcela
    
    'Para cada Vendedor em objComissaoVendedorPercentual adiciona em colComissao
    For Each objComissaoVendedorPercentual In colComissaoVendedorPercentual
        
        objComissao.iCodVendedor = objComissaoVendedorPercentual.iCodigo
        objComissao.dPercentual = objComissaoVendedorPercentual.dValor
        objComissao.dValorBase = dValorBaseParcela
        
        If objComissao.dValorBase > 0 And objComissao.dPercentual > 0 Then objComissao.dValor = objComissao.dPercentual * objComissao.dValorBase

        'Se vendedor e percentual não forem nulos adiciona a Comissao na coleção
        If objComissao.iCodVendedor <> 0 And objComissao.dPercentual <> 0 Then colComissao.Add 0, giFilialEmpresa, 0, 0, 0, objComissao.iCodVendedor, DATA_NULA, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor, DATA_NULA
    
    Next
    
    Exit Function
        
    Inicializa_ComissaoParcela = SUCESSO

    Exit Function

Erro_Inicializa_ComissaoParcela:

    Inicializa_ComissaoParcela = Err

    Select Case Err

        Case 26473, 26604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)

        Case 27634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", Err, objTipoCliente.iCodigo)

        Case 31350, 31351, 31353, 31354, 31355

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160845)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iParcela As Integer
Dim colParcelaReceber As New colParcelaReceber
Dim colNFiscalMarcado As New Collection
Dim objGeracaoFatura As New ClassGeracaoFatura
Dim iSelecionada As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    If iTabPrincipalAlterado = REGISTRO_ALTERADO Then Error 58410
    
    'Verifica se os campos obrigatórios da tela estão preenchidos
    If Len(Trim(CodCliente.Text)) = 0 Then Error 31288
    If Len(Trim(FilialCliente.Text)) = 0 Then Error 31289
    If Len(Trim(EmissaoFatura.ClipText)) = 0 Then Error 36731

    'Verifica se existe alguma Nota Fiscal no grid de Notas Fiscais
    If objGridNFiscal.iLinhasExistentes = 0 Then Error 31277

    Call Verifica_NFiscal_Selecionada(iSelecionada)
    
    'se nao tiver uma nf selecionada => erro
    If iSelecionada = 0 Then Error 31278

    'Verifica se existe alguma Parcela no TabDescontosComissoes
    If Len(Trim(Parcela.Caption)) > 0 Then

        'Recolhe os dados da Parcela que está na Tela
        iParcela = CInt(Parcela.Caption)

        lErro = Recolhe_Parcela_Tela(iParcela)
        If lErro <> SUCESSO Then Error 31280

    End If

    lErro = Move_GridParcelas_Memoria(objGeracaoFatura)
    If lErro <> SUCESSO Then Error 51126

    Call Atualiza_Colecoes
    
    'Recolhe os dados Básicos
    lErro = Move_Tela_Memoria(objGeracaoFatura)
    If lErro <> SUCESSO Then Error 58437
    
    lErro = Move_NFiscais_Memoria(objGeracaoFatura)
    If lErro <> SUCESSO Then Error 58438
        
    'Valida os Dados do Grid de Parcelas
    lErro = Valida_GridParcelas()
    If lErro <> SUCESSO Then Error 31291

    'Valida os dados particulares de cada Parcela
    lErro = Valida_Dados_Parcelas()
    If lErro <> SUCESSO Then Error 42421

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(EmissaoFatura.Text))
    If lErro <> SUCESSO Then gError 92040

    'Chama a Rotina que vai Gerar a nova Fatura
    lErro = CF("GeracaoFatura_GerarFatura", objGeracaoFatura, objGeracaoFatura.colNFiscalInfo, colcolComissao, colcolDesconto, objContabil)
    If lErro <> SUCESSO Then Error 46184

    GL_objMDIForm.MousePointer = vbDefault
    
    vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_FATURA", objGeracaoFatura.lNumTitulo)
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 31288
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 31289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 31277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURA_SEM_NOTASFISCAIS", Err)

        Case 31278, 58410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTASFISCAIS_NAO_SELECIONADAS", Err)

        Case 31280, 31291, 42421, 46184, 36732, 51126, 58437, 58438, 92040

        Case 36731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FATURA_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160846)

    End Select

    Exit Function

End Function

Function Move_NFiscais_Memoria(objGeracaoFatura As ClassGeracaoFatura) As Long
'Move as Notas Fiscais para o objGeracaoFatura.colNFiscalInfo

Dim lErro As Long
Dim iIndice As Integer
Dim objNFiscal As ClassNFiscal
Dim objNFiscalInfo As ClassNFiscalInfo

On Error GoTo Erro_Move_NFiscais_Memoria
    
    Set objGeracaoFatura.colNFiscalInfo = New Collection
    
    'Para Cada Nota Fiscal do Grid de Notas Fiscais
    For iIndice = 1 To objGridNFiscal.iLinhasExistentes
        
        'Verifica se a Nota está Marcada
        If GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col) = S_MARCADO Then
                            
            Set objNFiscal = New ClassNFiscal
            Set objNFiscalInfo = New ClassNFiscalInfo
                
            objNFiscal.sSerie = GridNFiscal.TextMatrix(iIndice, iGrid_Serie_Col)
            objNFiscal.lNumNotaFiscal = GridNFiscal.TextMatrix(iIndice, iGrid_Numero_Col)
            objNFiscal.iFilialEmpresa = giFilialEmpresa
            objNFiscal.dtDataEmissao = CDate(GridNFiscal.TextMatrix(iIndice, iGrid_DataEmissao_Col))
                
            'Lê o NumIntDoc da NFiscal
            lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
            If lErro <> SUCESSO And lErro <> 43676 Then Error 58433
    
            'Se não encontrou
            If lErro = 43676 Then Error 58434
            
            'Adiciona na Colecao o NumintDoc
            objNFiscalInfo.lNumIntDoc = objNFiscal.lNumIntDoc
                
            objGeracaoFatura.colNFiscalInfo.Add objNFiscalInfo
             
        End If
    
    Next

    Move_NFiscais_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_NFiscais_Memoria:

    Move_NFiscais_Memoria = Err
                      
    Select Case Err

        Case 58433 'Tratado na Rotina Chamada
        
        Case 58434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NUM_SERIE_NAO_CADASTRADA", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160847)

    End Select

    Exit Function
                
End Function


Function Move_Tela_Memoria(objGeracaoFatura As ClassGeracaoFatura) As Long
'Move os dados Basicos da Geracao para a Memoria

Dim dValorNFsSelecionadas As Double
Dim objcliente As New ClassCliente
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objGeracaoFatura.dtDataEmissao = CDate(EmissaoFatura.Text)
    objGeracaoFatura.dValorIRRF = StrParaDbl(ValorIRRF.Text)
    
    If ISSRetido.Value = 1 Then
        objGeracaoFatura.dValorISS = StrParaDbl(ISSValor.Caption)
    End If
    
    Call Calcula_Valor_Notas_Selecionadas(dValorNFsSelecionadas)
    objGeracaoFatura.dValorNFsSelecionadas = dValorNFsSelecionadas
    
    objGeracaoFatura.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
    
    'Lê o Codigo do0 Cliente atraves do Nome Reduzido
    objcliente.sNomeReduzido = CodCliente.Text
        
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 58430
    
    If lErro = 12348 Then Error 58431
    
    objGeracaoFatura.lCliente = objcliente.lCodigo
    objGeracaoFatura.iFilialCli = Codigo_Extrai(FilialCliente.Text)

    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err
        
        Case 58430
        
        Case 58431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objcliente.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160848)

    End Select

    Exit Function

End Function

Sub Rotina_Grid_Enable(iLinha As Integer, objControle As Object, iCaminho As Integer)

Dim iTipo As Integer

    'Pesquisa a controle da coluna em questão
    Select Case objControle.Name

        Case Data.Name
            
            'Se o Tipo ja foi Preenchido
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then
                Data.Enabled = True
            Else
                Data.Enabled = False
            End If

        Case ValorDesconto.Name
            
            'Se o Tipo Foi Preenchido
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then

                iTipo = Codigo_Extrai(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))
                
                'Se é do Tipo valor
                If iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO Then
                    objControle.Enabled = True
                Else
                    objControle.Enabled = False
                End If
            Else
                objControle.Enabled = False
            End If


        Case Percentual1.Name
                
            'Se o Tipo foi Preenchido
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then

                iTipo = Codigo_Extrai(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))
                
                'Se é do Tipo Percentual
                If iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual Then
                    objControle.Enabled = True
                Else
                    objControle.Enabled = False
                End If
            Else
                objControle.Enabled = False
            End If

        Case ValorParcela.Name
            'Se o vencimento estiver preenchido, habilita o controle
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))) = 0 Then
                objControle.Enabled = False
            Else
                objControle.Enabled = True
            End If

    End Select

    Exit Sub

End Sub

Private Function Move_GridParcelas_Memoria(objGeracaoFatura As ClassGeracaoFatura) As Long
'Move para a memória os dados existentes no Grid

Dim iIndice As Integer
Dim objParcelaReceber As ClassParcelaReceber
Dim lErro As Long
Dim colParcelaReceber As colParcelaReceber

    Set colParcelaReceber = New colParcelaReceber
    
    'Para cada item do Grid de Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaReceber = New ClassParcelaReceber

        'Preenche objParcelaReceber com a linha do GridParcelas
        objParcelaReceber.iNumParcela = iIndice
        If Len(Trim((GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col)))) <> 0 Then objParcelaReceber.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))
        If Len(Trim((GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col)))) <> 0 Then objParcelaReceber.dtDataVencimentoReal = CDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))

        objParcelaReceber.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))

        'Adiciona objParcelaReceber à coleção colParcelas
        With objParcelaReceber
        '##############################################
        'ALTERADO POR WAGNER
            colParcelaReceber.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValor
        '##############################################
        End With
    Next

    Set objGeracaoFatura.colParcelas = colParcelaReceber

    Move_GridParcelas_Memoria = SUCESSO

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_GERACAO_FATURA_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Geração Fatura"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoFatura"

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

Function Calcula_Comissoes() As Long
'Recalcula Todas as Comissões

Dim lErro As Long
Dim colComissaoVendedorPercentual As New Collection
Dim colComissao As colComissao
Dim dValor As Double
Dim iIndice As Integer

On Error GoTo Erro_Calcula_Comissoes
    
    'Reinicializa as Coleções
    Set colcolComissao = New colcolComissao
        
    'Se não tem Condicao de Pagamento Selecionada não faz sentido ter Comissao para parcela
    If objGridParcelas.iLinhasExistentes <= 0 Then Exit Function
    
    'Le todas as comissoes para as Notas selecionadas e retorna uma colecao de
    lErro = NFiscal_Monta_ColecaoComissao(colComissaoVendedorPercentual)
    If lErro <> SUCESSO Then Error 58400
                        
    'Para cada parcela do GridParcelas, Calcula as Comissoes
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        Set colComissao = New colComissao

        'Pega o valor da Parcela em questão
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))) > 0 Then dValor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        
        'Monta a Colecao de Comissao para cada parcela -- Passa: (Valor da Parcela, Colecao de Vendedor e Percentual), Devolve: (Colecao de Vendedores formando a comisao da Parcela)
        lErro = Inicializa_ComissaoParcela(dValor, colComissaoVendedorPercentual, colComissao)
        If lErro <> SUCESSO Then Error 58401
        
        'Adiciona na coleção global de Comissões
        colcolComissao.Add colComissao
        
    Next
   
    Calcula_Comissoes = SUCESSO
    
    Exit Function
    
Erro_Calcula_Comissoes:
    
    Calcula_Comissoes = Err
    
    Select Case Err
        
        Case 58400, 58401 'Tratados nas Rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160849)
        
    End Select
    
    Exit Function
    
End Function

Function Calcula_Descontos()
'Recalcula todos os Descontos

Dim lErro As Long
Dim iIndice As Integer
Dim objDescontoPadrao As New ClassDescontoPadrao
Dim colDescontoPadrao As Collection
Dim colDesconto As colDesconto

On Error GoTo Erro_Calcula_Descontos
    
    'Reinicializa a coleçào de Descontos
    Set colcolDesconto = New colcolDesconto
        
    'Se não tem Condicao de Pagamento Selecionada não faz sentido ter Desconto para parcela
    If objGridParcelas.iLinhasExistentes <= 0 Then Exit Function
    
    'Para cada Parcela no GridParcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set colDescontoPadrao = New Collection
        Set colDesconto = New colDesconto
        
        'Le os Percentuais de descontos
        lErro = CF("DescontoPadrao_Le", colDescontoPadrao)
        If lErro <> SUCESSO Then Error 58411


        For Each objDescontoPadrao In colDescontoPadrao

            'Se os atributos de objDescontoPadrao padrão estiverem preenchidos adiciona em coldesconto
            If objDescontoPadrao.iCodigo > 0 And objDescontoPadrao.dPercentual > 0 And Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))) > 0 Then colDesconto.Add objDescontoPadrao.iCodigo, CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col)) - objDescontoPadrao.iDias, objDescontoPadrao.dPercentual

        Next
        
        'Adiciona a ColDescontos na Colecao Global de Descontos
        colcolDesconto.Add colDesconto

    Next
   
    Calcula_Descontos = SUCESSO
    
    Exit Function
    
Erro_Calcula_Descontos:

    Calcula_Descontos = Err
    
    Select Case Err
        
        Case 58411 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160850)
        
    End Select
    
    Exit Function
        
End Function

Sub Calcula_Comissoes_Parcela(iParcela As Integer, dValorBaseParcela As Double)
'Recalcula a Comissao de acordo com o Valor da Parcela

Dim lErro As Long
Dim objComissao As ClassComissao
        
    If iParcela > colcolComissao.Count Then
                
        lErro = Calcula_Comissoes()
        If lErro <> SUCESSO Then Error 58439
    
        lErro = Calcula_Descontos
        If lErro <> SUCESSO Then Error 58440
    
    End If
    
    'Para cada Comissao da Parcela passada
    For Each objComissao In colcolComissao.Item(iParcela).colComissao
        
        'Atualiza o Valor Base
        objComissao.dValorBase = dValorBaseParcela
        
        'Recalcula o Valor da Comissão
        If objComissao.dValorBase > 0 And objComissao.dPercentual > 0 Then objComissao.dValor = objComissao.dPercentual * objComissao.dValorBase
    
    Next
    
End Sub

Sub Calcula_Valor_Notas_Selecionadas(dValorNotasSelecionadas As Double)
'Soma os valores das NotaFiscal selecionadas no grid

Dim iIndice As Integer
    
    For iIndice = 1 To objGridNFiscal.iLinhasExistentes
        
        If GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col) = S_MARCADO Then
            
            dValorNotasSelecionadas = dValorNotasSelecionadas + GridNFiscal.TextMatrix(iIndice, iGrid_ValorTotal_Col)
        
        End If
    
    Next
    
End Sub

Sub Verifica_NFiscal_Selecionada(iSelecionada As Integer)
'Verifica se tem pelo menos uma NotaFiscal selecionada no grid
'Se tiver delvolve iSelecionada = 1

Dim iIndice As Integer
    
    For iIndice = 1 To objGridNFiscal.iLinhasExistentes
        
        If GridNFiscal.TextMatrix(iIndice, iGrid_GeraFatura_Col) = S_MARCADO Then
            
            iSelecionada = 1
            Exit Sub
        
        End If
    Next
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Serie Then
            Call LabelSerie_Click
        ElseIf Me.ActiveControl Is CodCliente Then
            Call LblCliente_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call BotaoNFiscal_Click
        End If
    
    End If

End Sub


Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LabelSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerie, Source, X, Y)
End Sub

Private Sub LabelSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerie, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LblCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblCliente, Source, X, Y)
End Sub

Private Sub LblCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblCliente, Button, Shift, X, Y)
End Sub

Private Sub LblFilialCli_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilialCli, Source, X, Y)
End Sub

Private Sub LblFilialCli_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilialCli, Button, Shift, X, Y)
End Sub

Private Sub LabelTotalNotaSel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotalNotaSel, Source, X, Y)
End Sub

Private Sub LabelTotalNotaSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotalNotaSel, Button, Shift, X, Y)
End Sub

Private Sub TotalNotasSel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalNotasSel, Source, X, Y)
End Sub

Private Sub TotalNotasSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalNotasSel, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LabelTotaisComissoes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotaisComissoes, Source, X, Y)
End Sub

Private Sub LabelTotaisComissoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotaisComissoes, Button, Shift, X, Y)
End Sub

Private Sub TotalValorComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorComissao, Source, X, Y)
End Sub

Private Sub TotalValorComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorComissao, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentualComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualComissao, Source, X, Y)
End Sub

Private Sub TotalPercentualComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualComissao, Button, Shift, X, Y)
End Sub

Private Sub Parcela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Parcela, Source, X, Y)
End Sub

Private Sub Parcela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Parcela, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub ISSValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ISSValor, Source, X, Y)
End Sub

Private Sub ISSValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ISSValor, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub


Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = CodCliente 'atenção este nome de campo é diferente dos outros do sistema (Cliente)
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134019

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134019

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160851)

    End Select
    
    Exit Sub

End Sub

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




