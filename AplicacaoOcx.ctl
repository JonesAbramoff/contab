VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl AplicacaoOcx 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   9405
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4605
      Index           =   1
      Left            =   195
      TabIndex        =   0
      Top             =   825
      Width           =   9075
      Begin VB.Frame Frame3 
         Caption         =   "Dados Principais"
         Height          =   2250
         Left            =   225
         TabIndex        =   43
         Top             =   285
         Width           =   8235
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2880
            Picture         =   "AplicacaoOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Numeração Automática"
            Top             =   450
            Width           =   300
         End
         Begin VB.ComboBox TipoAplicacao 
            Height          =   315
            Left            =   5325
            TabIndex        =   3
            Top             =   405
            Width           =   2595
         End
         Begin VB.ComboBox CodContaCorrente 
            Height          =   315
            Left            =   1785
            TabIndex        =   4
            Top             =   1065
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1785
            TabIndex        =   1
            Top             =   435
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   300
            Left            =   6495
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1035
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ValorAplicado 
            Height          =   285
            Left            =   1770
            TabIndex        =   6
            Top             =   1725
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataAplicacao 
            Height          =   300
            Left            =   5325
            TabIndex        =   5
            Top             =   1050
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
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
            Left            =   4785
            TabIndex        =   52
            Top             =   1095
            Width           =   480
         End
         Begin VB.Label Label4 
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
            Left            =   1185
            TabIndex        =   53
            Top             =   1755
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Aplicado:"
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
            Left            =   3915
            TabIndex        =   54
            Top             =   1770
            Width           =   1350
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
            Left            =   1035
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   55
            Top             =   465
            Width           =   660
         End
         Begin VB.Label Label15 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4815
            TabIndex        =   56
            Top             =   450
            Width           =   450
         End
         Begin VB.Label LabelCtaCorrente 
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente:"
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
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   57
            Top             =   1110
            Width           =   1350
         End
         Begin VB.Label LabelSaldoAplicado 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5325
            TabIndex        =   58
            Top             =   1740
            Width           =   1740
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pagamento"
         Height          =   1590
         Left            =   225
         TabIndex        =   44
         Top             =   2700
         Width           =   8250
         Begin VB.CommandButton BotaoImprimir 
            Caption         =   "Imprimir Cheque"
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
            Height          =   540
            Left            =   6120
            Picture         =   "AplicacaoOcx.ctx":00EA
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   630
            Width           =   1725
         End
         Begin VB.ComboBox TipoMeioPagto 
            Height          =   315
            Left            =   1740
            TabIndex        =   7
            Top             =   390
            Width           =   1695
         End
         Begin VB.ComboBox Favorecido 
            Height          =   315
            Left            =   1770
            TabIndex        =   9
            Top             =   990
            Width           =   4020
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   300
            Left            =   4680
            TabIndex        =   8
            Top             =   375
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label Label12 
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
            Height          =   195
            Left            =   3900
            TabIndex        =   59
            Top             =   420
            Width           =   705
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Forma:"
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
            Left            =   1095
            TabIndex        =   60
            Top             =   450
            Width           =   585
         End
         Begin VB.Label LabelFavorecido 
            AutoSize        =   -1  'True
            Caption         =   "Favorecido:"
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
            Left            =   660
            TabIndex        =   61
            Top             =   1020
            Width           =   1020
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7155
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "AplicacaoOcx.ctx":01EC
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "AplicacaoOcx.ctx":0346
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "AplicacaoOcx.ctx":04D0
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "AplicacaoOcx.ctx":0A02
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4605
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   810
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4920
         TabIndex        =   88
         Tag             =   "1"
         Top             =   1560
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
         TabIndex        =   23
         Top             =   405
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
         TabIndex        =   21
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   930
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
         TabIndex        =   22
         Top             =   90
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   31
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
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   33
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   32
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   35
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   45
         Top             =   3450
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   62
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1125
            TabIndex        =   63
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   64
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   65
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
         Left            =   3450
         TabIndex        =   26
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
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
         TabIndex        =   30
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
         TabIndex        =   28
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
         TabIndex        =   46
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   34
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
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   24
         Top             =   720
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   45
         TabIndex        =   66
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   67
         Top             =   120
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
         TabIndex        =   68
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   69
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   70
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
         TabIndex        =   71
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
         TabIndex        =   72
         Top             =   945
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
         TabIndex        =   73
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   74
         Top             =   1275
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
         TabIndex        =   75
         Top             =   1275
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1800
         TabIndex        =   76
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   77
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   78
         Top             =   3030
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
         TabIndex        =   79
         Top             =   555
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   80
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   81
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4605
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   810
      Visible         =   0   'False
      Width           =   9075
      Begin VB.Frame Frame4 
         Caption         =   "Complemento"
         Height          =   1605
         Left            =   360
         TabIndex        =   47
         Top             =   240
         Width           =   8235
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1710
            TabIndex        =   11
            Top             =   405
            Width           =   5085
         End
         Begin MSMask.MaskEdBox NumRefExterna 
            Height          =   300
            Left            =   1725
            TabIndex        =   12
            Top             =   1020
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
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
            Left            =   810
            TabIndex        =   82
            Top             =   465
            Width           =   825
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Externo:"
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
            Left            =   450
            TabIndex        =   83
            Top             =   1050
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Resgate"
         Height          =   1965
         Left            =   360
         TabIndex        =   49
         Top             =   2160
         Width           =   8235
         Begin VB.CommandButton BotaoHistorico 
            Caption         =   "Histórico..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   4605
            Picture         =   "AplicacaoOcx.ctx":0B80
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   945
            Width           =   1635
         End
         Begin MSMask.MaskEdBox TaxaPrevista 
            Height          =   300
            Left            =   5475
            TabIndex        =   14
            Top             =   405
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorResgatePrevisto 
            Height          =   300
            Left            =   1650
            TabIndex        =   15
            Top             =   1185
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   300
            Left            =   2730
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataResgatePrevista 
            Height          =   300
            Left            =   1650
            TabIndex        =   13
            Top             =   435
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
            Caption         =   "Taxa (%):"
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
            Left            =   4170
            TabIndex        =   84
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Valor Previsto:"
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
            Left            =   300
            TabIndex        =   85
            Top             =   1215
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Data Prevista:"
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
            Left            =   330
            TabIndex        =   86
            Top             =   465
            Width           =   1230
         End
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4980
      Left            =   60
      TabIndex        =   51
      Top             =   480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8784
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
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
Attribute VB_Name = "AplicacaoOcx"
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

Private Const CODIGO1 As String = "Codigo"
Private Const CONTACORRENTE1 As String = "Conta_Corrente"
Private Const VALOR1 As String = "Valor_Aplicado"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const HISTORICO1 As String = "Historico"
Private Const VALORPREVISTO1 As String = "Valor_Resg_Prev"
Private Const TIPO1 As String = "Tipo_Aplicacao"
Private Const CTACONTACORRENTE As String = "Cta_Conta_Corrente"
Private Const CTATIPOAPLICACAO As String = "Cta_Tipo_Aplicacao"

'fim contabilidade
Dim iFrameAtual As Integer
Public iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1
Private WithEvents objEventoHistorico As AdmEvento
Attribute objEventoHistorico.VB_VarHelpID = -1
Private WithEvents objEventoFavorecido As AdmEvento
Attribute objEventoFavorecido.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Complemento = 2
Private Const TAB_Contabilizacao = 3

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objBanco As New ClassBanco
Dim sLayoutCheque As String
Dim objInfoChequePag As New ClassInfoChequePag
Dim dtDataEmissao As Date
Dim lNumImpressao As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 87019
    If Len(Trim(TipoAplicacao.Text)) = 0 Then gError 87020
    If Len(Trim(CodContaCorrente.Text)) = 0 Then gError 87021
    If Len(Trim(DataAplicacao.ClipText)) = 0 Then gError 87022
    If Len(Trim(ValorAplicado.Text)) = 0 Then gError 87023
    If Len(Trim(Favorecido.Text)) = 0 Then gError 87024
    
    'Retira o código da combo e passa para iCodigo
    iCodigo = Codigo_Extrai(CodContaCorrente.Text)
     
    'Le a Conta Corrente a partir de iCodigo passado como parâmetro
    lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 87025

    'Caso a Conta Corrente não tiver sido encontrada dispara erro
    If lErro = 11807 Then gError 87026

    'Caso a Conta Corrente não for bancária dispara erro
    If objContaCorrenteInt.iCodBanco = 0 Then gError 87027
    
    'Atribui o valor retornado de objContaCorrenteInt.iCodBanco a objBanco.iCodBanco
    objBanco.iCodBanco = objContaCorrenteInt.iCodBanco
          
    'Le o Banco a partir de objBanco.iCodBanco
    lErro = CF("Banco_Le", objBanco)
    If lErro <> SUCESSO And lErro <> 16091 Then gError 87028
        
    'Caso o banco não tiver sido encontrado dispara erro
    If lErro = 16091 Then gError 87029
        
    'Atribui retorno de objBanco.sLayoutCheque a variavel sLayoutCheque
    sLayoutCheque = objBanco.sLayoutCheque
                                                                                               
    'Recolhe os dados do cheque da tela para objInfoChequePag
    Call Move_tela_Cheque(objInfoChequePag, dtDataEmissao)

    'Chama a função que prepara a impressão do cheque
    lErro = CF("PreparaImpressao_Cheque", lNumImpressao, objInfoChequePag)
    If lErro <> SUCESSO Then gError 87030

    'Chama a função responsável pela impressão do cheque
    lErro = ImprimirCheques(lNumImpressao, sLayoutCheque, dtDataEmissao)
    If lErro <> SUCESSO Then gError 87031
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 87025, 87028, 87030, 87031

        Case 87019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 87020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_NAO_PREENCHIDO", gErr)

        Case 87021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)
                                           
        Case 87022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_APLICACAO_NAO_INFORMADA", gErr)

        Case 87023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)

        Case 87024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_NAO_PREENCHIDO", gErr)
        
        Case 87026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", gErr, CodContaCorrente.Text)
        
        Case 87027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", gErr)

        Case 87029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, objBanco.iCodBanco)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142916)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera o proximo codigo da aplicacao
    lErro = CF("Aplicacao_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 57543

    'Mostra na tela o proximo numero de aplicacao disponível
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57543
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142917)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro  As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objAplicacao As New ClassAplicacao

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Codigo foi informado
    If Len(Trim(Codigo.Text)) = 0 Then Error 17437

    objAplicacao.lCodigo = CLng(Codigo.Text)

    'Pede confirmacao da exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_APLICACAO", objAplicacao.lCodigo)

    If vbMsgRes = vbYes Then

        'Chama a rotina de exclusao
        lErro = CF("Aplicacao_Exclui", objAplicacao, objContabil)
        If lErro <> SUCESSO Then Error 17438

        Call Limpa_Tela_Aplicacao

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17437, 17438

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142918)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objAplicacao As New ClassAplicacao

On Error GoTo Erro_BotaoGravar_Click

    'Chama a rotina de gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 17319

    Call Limpa_Tela_Aplicacao

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 17319

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142919)

    End Select

    Exit Sub

End Sub

Private Sub BotaoHistorico_Click()

Dim colSelecao As New Collection

    If Len(Trim(Codigo.Text)) > 0 Then

        colSelecao.Add CInt(Codigo.Text)

        Call Chama_Tela("ResgateLista_Aplicacao", colSelecao)

    End If

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se algum campo da tela foi modificado
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17464

    'Limpa a tela
    Call Limpa_Tela_Aplicacao

    BotaoImprimir.Enabled = False

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 17464

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142920)

    End Select

    Exit Sub

End Sub

Private Sub CodContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodContaCorrente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_CodContaCorrente_Validate

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox CodContacOrrente
    If CodContaCorrente.Text = CodContaCorrente.List(CodContaCorrente.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(CodContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17356
    
    objContaCorrenteInt.iCodigo = iCodigo
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 17357

        'Não encontrou a Conta Corrente no BD
        If lErro = 11807 Then Error 17358
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43526

        End If
        
        'Encontrou a Conta Corrente no BD, coloca no Text da Combo
        CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 17359
    
    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 17356, 17357

        Case 17358
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Contas Correntes
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 17359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE1", Err, CodContaCorrente.Text)

        Case 43526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodContaCorrente.Text, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142921)

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

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Codigo.Text) Then Error 55955

        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 55956

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 55955, 55956
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO1", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142922)

    End Select

    Exit Sub

End Sub

Private Sub DataAplicacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAplicacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAplicacao, iAlterado)

End Sub

Private Sub DataAplicacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAplicacao_Validate

    'Verifica se a data de aplicacao está preenchida
    If Len(DataAplicacao.ClipText) <> 0 Then

        'Verifica se a data final é válida
        lErro = Data_Critica(DataAplicacao.Text)
        If lErro <> SUCESSO Then Error 17305

    End If

    Exit Sub

Erro_DataAplicacao_Validate:

    Cancel = True


    Select Case Err

        Case 17305

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142923)

    End Select

    Exit Sub

End Sub

Private Sub DataResgatePrevista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataResgatePrevista_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataResgatePrevista, iAlterado)

End Sub

Private Sub DataResgatePrevista_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataResgatePrevista_Validate

    'Verifica se a data de resgate prevista está preenchida
    If Len(DataResgatePrevista.ClipText) <> 0 Then

        'Verifica se a data final é válida
        lErro = Data_Critica(DataResgatePrevista.Text)
        If lErro <> SUCESSO Then Error 17307

    End If

    Exit Sub

Erro_DataResgatePrevista_Validate:

    Cancel = True


    Select Case Err

        Case 17307

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142924)

    End Select

    Exit Sub

End Sub

Private Sub Favorecido_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Favorecido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFavorecido As New ClassFavorecidos
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Favorecido_Validate

    'Verifica se foi preenchido o Favorecido
    If Len(Trim(Favorecido.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox Favorecido
    If Favorecido.Text = Favorecido.List(Favorecido.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(Favorecido, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17366

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objFavorecido.iCodigo = iCodigo

        lErro = CF("Favorecido_Le", objFavorecido)
        If lErro <> SUCESSO And lErro <> 17015 Then Error 17367

        'Não encontrou o Favorecido no BD
        If lErro = 17015 Then Error 17368

        'Encontrou o Favorecido no BD, coloca no Text da Combo
        Favorecido.Text = CStr(objFavorecido.iCodigo) & SEPARADOR & objFavorecido.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 17369

    Exit Sub

Erro_Favorecido_Validate:

    Cancel = True


    Select Case Err

        Case 17366, 17367

        Case 17368
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FAVORECIDO_INEXISTENTE", objFavorecido.iCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Favorecidos
                Call Chama_Tela("Favorecidos", objFavorecido)
            Else
            End If

        Case 17369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_INEXISTENTE1", Err, Favorecido.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142925)

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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoCodigo = Nothing
    Set objEventoContaCorrenteInt = Nothing
    Set objEventoHistorico = Nothing
    Set objEventoFavorecido = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Historico_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iTamanho As Integer
Dim iCodigo As Integer
Dim iIndice As Integer
Dim sDescricao As Long
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Historico_Validate

    'Verifica o tamanho do texto em historico
    iTamanho = Len(Trim(Historico.Text))

    If iTamanho = 0 Then Exit Sub

    'Verifica se é maior que o tamanho maximo
    If iTamanho > 50 Then Error 17313
    
    'Verifica se o que foi digitado é numerico
    If IsNumeric(Trim(Historico.Text)) Then

        lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
        If lErro <> SUCESSO Then Error 40711
        
        'peenche o objeto
        objHistMovCta.iCodigo = CInt(Trim(Historico.Text))
                
        'verifica se existe hitorico relacionado com codigo passado
        lErro = CF("HistMovCta_Le", objHistMovCta)
        If lErro <> SUCESSO And lErro <> 15011 Then Error 40796
        
        'se nao existir -----> Erro
        If lErro = 15011 Then Error 40797
        
        Historico.Text = objHistMovCta.sDescricao
        
    End If
   
    Exit Sub

Erro_Historico_Validate:

    Cancel = True


    Select Case Err

        Case 17313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", Err)
        
        Case 40711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, Historico.Text)
            
        Case 40796
        
        Case 40797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142926)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    End If
    
    iFrameAtual = 1

    Set objEventoCodigo = New AdmEvento
    Set objEventoContaCorrenteInt = New AdmEvento
    Set objEventoHistorico = New AdmEvento
    Set objEventoFavorecido = New AdmEvento
    
    iAlterado = 0

    'Lê os tipos de aplicacao com codigo e a descricao existentes no BD e carrega na ComboBox
    lErro = Carrega_TiposDeAplicacao()
    If lErro <> SUCESSO Then Error 17477

    'Lê as contas correntes  com codigo e o nome reduzido existentes no BD e carrega na ComboBox
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then Error 17194

    'Lê os tipos de pagamentos ativos existentes no BD
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then Error 17228

    'Lê os favorecidos com codigo e o nome existentes no BD e carrega na ComboBox
    lErro = Carrega_Favorecidos()
    If lErro <> SUCESSO Then Error 17225

    'Lê os historicos com codigo e o nome existentes no BD e carrega na ComboBox
    lErro = Carrega_Historico()
    If lErro <> SUCESSO Then Error 17226
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_TESOURARIA)
    If lErro <> SUCESSO Then Error 39547
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 17194, 17225, 17226, 17228, 17447, 39547

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142927)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_TiposDeAplicacao() As Long
'Carrega os tipos de aplicacao na combo de tipos de aplicacao

Dim lErro As Long
Dim colTiposDeAplicacao As New Collection
Dim objTiposDeAplicacao As ClassTiposDeAplicacao

On Error GoTo Erro_Carrega_TiposDeAplicacao

    'Lê os tipos de aplicacao ativos existentes no BD
    lErro = CF("TiposDeAplicacao_Le_Ativos", colTiposDeAplicacao)
    If lErro <> SUCESSO Then Error 17189

    'Preenche a ComboBox TipoAplicacao com os objetos da colecao colTipoAplicacao
    For Each objTiposDeAplicacao In colTiposDeAplicacao

       'Insere na combo de tipos de aplicacao
       TipoAplicacao.AddItem CStr(objTiposDeAplicacao.iCodigo) & SEPARADOR & objTiposDeAplicacao.sDescricao
       TipoAplicacao.ItemData(TipoAplicacao.NewIndex) = objTiposDeAplicacao.iCodigo

    Next

    Carrega_TiposDeAplicacao = SUCESSO

    Exit Function

Erro_Carrega_TiposDeAplicacao:

    Carrega_TiposDeAplicacao = Err

    Select Case Err

        Case 17189

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142928)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39549

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39550

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 39549
        
        Case 39550
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objAplicacao As New ClassAplicacao
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) = 0 Then
        objAplicacao.lCodigo = 0
    Else
        objAplicacao.lCodigo = CLng(Codigo.Text)
    End If

    Call Chama_Tela("AplicacaoLista", colSelecao, objAplicacao, objEventoCodigo)

End Sub

Private Sub LabelCtaCorrente_Click()

Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim colSelecao As Collection

    If Len(CodContaCorrente.Text) = 0 Then
        objContaCorrenteInt.iCodigo = 0
    Else
        If CodContaCorrente.ListIndex <> -1 Then objContaCorrenteInt.iCodigo = CodContaCorrente.ItemData(CodContaCorrente.ListIndex)
    End If

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContaCorrenteInt, objEventoContaCorrenteInt)

End Sub

Private Sub LabelSaldoAplicado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub NumRefExterna_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1

    CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo)

    Call CodContaCorrente_Validate(bSGECancelDummy)

    iAlterado = 0

    Me.Show

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
                Parent.HelpContextID = IDH_APLICACAO_ID
                
            Case TAB_Complemento
                Parent.HelpContextID = IDH_APLICACAO_COMPLEMENTO
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_APLICACAO_CONTABILIZACAO
                        
        End Select

    End If

End Sub

Private Function Carrega_TipoMeioPagto() As Long
'Carrega na Combo TipoMeioPagto os tipo de meio de pagamento ativos

Dim lErro As Long
Dim colTipoMeioPagto As New Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Le todos os tipo de pagamento.
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 17229

    For Each objTipoMeioPagto In colTipoMeioPagto

        'Verifica se estao ativos
        If objTipoMeioPagto.iInativo = TIPOMEIOPAGTO_ATIVO Then

            'coloca na combo
            TipoMeioPagto.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
            TipoMeioPagto.ItemData(TipoMeioPagto.NewIndex) = objTipoMeioPagto.iTipo

        End If

    Next

    Carrega_TipoMeioPagto = SUCESSO

    Exit Function

Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = Err

    Select Case Err

        Case 17229

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142929)

    End Select

    Exit Function

End Function

Private Function Carrega_Historico() As Long
'Carrega a combo de historicos com os historicos da tabela "HistPadraMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'Le o Codigo e a descricao de todos os historicos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then Error 17230

    For Each objCodigoNome In colCodigoNome

        'Insere na Combo de historicos
        Historico.AddItem objCodigoNome.sNome
        Historico.ItemData(Historico.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_Historico = SUCESSO

    Exit Function

Erro_Carrega_Historico:

    Carrega_Historico = Err

    Select Case Err

        Case 17230

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142930)

    End Select
    
    Exit Function

End Function

Private Function Carrega_Favorecidos() As Long
'Carrega os favorecidos ativos na combo de Favorecidos

Dim lErro As Long
Dim objFavorecidos As ClassFavorecidos
Dim colFavorecidos As New Collection

On Error GoTo Erro_Carrega_Favorecidos

    'Le todos os favorecidos
    lErro = CF("Favorecidos_Le_Todos", colFavorecidos)
    If lErro <> SUCESSO Then Error 17231

    For Each objFavorecidos In colFavorecidos

        'Verifica se esta ativo
        If objFavorecidos.iInativo = FAVORECIDO_ATIVO Then

            'Insere na combo de Favorecidos
            Favorecido.AddItem CStr(objFavorecidos.iCodigo) & SEPARADOR & objFavorecidos.sNome
            Favorecido.ItemData(Favorecido.NewIndex) = objFavorecidos.iCodigo

        End If
    Next

    Carrega_Favorecidos = SUCESSO

    Exit Function

Erro_Carrega_Favorecidos:

    Carrega_Favorecidos = Err

    Select Case Err

        Case 17231

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142931)

    End Select

    Exit Function

End Function

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Carrega_CodContaCorrente

    'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 17132

    For Each objCodigoNome In colCodigoNomeConta

        'Insere na combo de contas correntes
        CodContaCorrente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 17132

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142932)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objAplicacao As ClassAplicacao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há uma aplicacao selecionada, exibir seus dados
    If Not (objAplicacao Is Nothing) Then

        'Verifica se tipo de aplicacao existe
        lErro = CF("Aplicacao_Le", objAplicacao)
        If lErro <> SUCESSO And lErro <> 17241 Then Error 17243

        'Se não encontrou a aplicação em questão
        If lErro = 17241 Then Error 17262

        'Verifica se a aplicacao esta ativa
        If objAplicacao.iStatus = APLICACAO_EXCLUIDA Then Error 17245

        lErro = Traz_Aplicacao_Tela(objAplicacao)
        If lErro <> SUCESSO Then Error 17248

    Else

        DataAplicacao.Text = Format(gdtDataAtual, "dd/mm/yy")

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 17243

        Case 17245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_EXCLUIDA", Err, objAplicacao.lCodigo)

        Case 17248
              Call Limpa_Tela_Aplicacao

        Case 17262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_INEXISTENTE", Err, objAplicacao.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142933)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objAplicacao As New ClassAplicacao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à tela
    sTabela = "Aplicacoes"

    If Len(Trim(Codigo.ClipText)) > 0 Then
        objAplicacao.lCodigo = CLng(Codigo.Text)
    Else
        objAplicacao.lCodigo = 0
    End If

    If Len(TipoAplicacao.Text) > 0 Then
        objAplicacao.iTipoAplicacao = Codigo_Extrai(TipoAplicacao.Text)
    Else
        objAplicacao.iTipoAplicacao = 0
    End If

    objAplicacao.dtDataAplicacao = CDate(DataAplicacao.Text)

    If Len(ValorAplicado.Text) > 0 Then
        objAplicacao.dValorAplicado = CDbl(ValorAplicado.Text)
    Else
        objAplicacao.dValorAplicado = 0
    End If

    If Len(Trim(LabelSaldoAplicado.Caption)) > 0 Then
        objAplicacao.dSaldoAplicado = CDbl(LabelSaldoAplicado.Caption)
    Else
        objAplicacao.dSaldoAplicado = 0
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "Codigo", objAplicacao.lCodigo, 0, "Codigo"
    colCampoValor.Add "TipoAplicacao", objAplicacao.iTipoAplicacao, 0, "TipoAplicacao"
    colCampoValor.Add "DataAplicacao", objAplicacao.dtDataAplicacao, 0, "DataAplicacao"
    colCampoValor.Add "ValorAplicado", objAplicacao.dValorAplicado, 0, "ValorAplicado"
    colCampoValor.Add "SaldoAplicado", objAplicacao.dSaldoAplicado, 0, "SaldoAplicado"

    colSelecao.Add "Status", OP_DIFERENTE, APLICACAO_EXCLUIDA

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142934)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objAplicacao As New ClassAplicacao

On Error GoTo Erro_Tela_Preenche

    objAplicacao.lCodigo = colCampoValor.Item("Codigo").vValor

    If objAplicacao.lCodigo <> 0 Then

        'Verifica se tipo de aplicacao existe
        lErro = CF("Aplicacao_Le", objAplicacao)
        If lErro <> SUCESSO And lErro <> 17241 Then Error 34668

        'Se não encontrou a aplicação em questão
        If lErro = 17241 Then Error 34669

        'Verifica se a aplicacao esta ativa
        If objAplicacao.iStatus = APLICACAO_EXCLUIDA Then Error 34670

        'Preenche a tela com os dados retornados
        lErro = Traz_Aplicacao_Tela(objAplicacao)
        If lErro <> SUCESSO Then Error 34671

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 34668

        Case 34669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_INEXISTENTE", Err, objAplicacao.lCodigo)

        Case 34670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_EXCLUIDA", Err, objAplicacao.lCodigo)

        Case 34671
              Call Limpa_Tela_Aplicacao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142935)

    End Select

    Exit Sub

End Sub

Private Function Traz_Aplicacao_Tela(objAplicacao As ClassAplicacao) As Long
'Coloca na Tela os dados da Aplicacao passada como parametro

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao
Dim objMovContaCorrente As New ClassMovContaCorrente
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objFavorecidos As New ClassFavorecidos
Dim objTipoMeioPagto As New ClassTipoMeioPagto

On Error GoTo Erro_Traz_Aplicacao_Tela

    objMovContaCorrente.lNumMovto = objAplicacao.lNumMovto

    'Carrega os dados do movimento relativos a aplicacao a partir da chave
    lErro = CF("MovContaCorrente_Le", objMovContaCorrente)
    If lErro <> SUCESSO And lErro <> 11893 Then Error 17244

    'Se o movimento não estiver cadastrado
    If lErro = 11893 Then Error 17246

    If objMovContaCorrente.iExcluido = EXCLUIDO Then Error 17478

    If objMovContaCorrente.iTipo <> MOVCCI_APLICACAO Then Error 17479

    Codigo.PromptInclude = False
    Codigo.Text = CStr(objAplicacao.lCodigo)
    Codigo.PromptInclude = True

    Numero.PromptInclude = False
    Numero.Text = CStr(objMovContaCorrente.lNumero)
    Numero.PromptInclude = True

    ValorAplicado.Text = CStr(objAplicacao.dValorAplicado)
    DataAplicacao.Text = Format(objAplicacao.dtDataAplicacao, "dd/MM/yy")
    LabelSaldoAplicado.Caption = Format(objAplicacao.dSaldoAplicado, "Standard")

    If objMovContaCorrente.sHistorico <> "" Then
        Historico.Text = objMovContaCorrente.sHistorico
    Else
        Historico.Text = ""
    End If

    If objMovContaCorrente.sNumRefExterna <> "" Then
        NumRefExterna.Text = objMovContaCorrente.sNumRefExterna
    Else
        NumRefExterna.Text = ""
    End If

    If objAplicacao.dtDataResgatePrevista <> DATA_NULA Then
        DataResgatePrevista = Format(objAplicacao.dtDataResgatePrevista, "dd/MM/yy")
    End If
    If objAplicacao.dValorResgatePrevisto <> 0 Then
        ValorResgatePrevisto.Text = CStr(objAplicacao.dValorResgatePrevisto)
    Else
        ValorResgatePrevisto.Text = ""
    End If

    If objAplicacao.dTaxaPrevista <> 0 Then
        TaxaPrevista.Text = Format(CDbl(objAplicacao.dTaxaPrevista), "Fixed")
    Else
        TaxaPrevista.Text = ""
    End If

    objTiposDeAplicacao.iCodigo = objAplicacao.iTipoAplicacao

    'Verifica se o tipo de aplicacao existe
    lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
    If lErro <> SUCESSO And lErro <> 17291 Then Error 17249

    If lErro = 17291 Then Error 17283

    If objTiposDeAplicacao.iInativo = TIPOAPLICACAO_INATIVO Then Error 17292

    TipoAplicacao.Text = CStr(objAplicacao.iTipoAplicacao) & SEPARADOR & objTiposDeAplicacao.sDescricao

    'Verifica se a conta corrente existe
    lErro = CF("ContaCorrenteInt_Le", objMovContaCorrente.iCodConta, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 17480

    If lErro = 11807 Then Error 17481

    CodContaCorrente.Text = CStr(objMovContaCorrente.iCodConta) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    'Verifica se o TiPoMeioPago existe
    objTipoMeioPagto.iTipo = objMovContaCorrente.iTipoMeioPagto

    lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
    If lErro <> SUCESSO And lErro <> 11909 Then Error 17482

    If lErro = 11909 Then Error 17483

    TipoMeioPagto.Text = CStr(objMovContaCorrente.iTipoMeioPagto) & SEPARADOR & objTipoMeioPagto.sDescricao

    'Verifica se o favorecido existe
    If objMovContaCorrente.iFavorecido <> 0 Then
        objFavorecidos.iCodigo = objMovContaCorrente.iFavorecido

        lErro = CF("Favorecido_Le", objFavorecidos)
        If lErro <> SUCESSO And lErro <> 17015 Then Error 17484

        If lErro = 11807 Then Error 17485

        Favorecido.Text = CStr(objFavorecidos.iCodigo) & SEPARADOR & objFavorecidos.sNome
    Else
        Favorecido.Text = ""
    End If

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objAplicacao.lNumMovto)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 39548
    
    iAlterado = 0

    Traz_Aplicacao_Tela = SUCESSO

    Exit Function

Erro_Traz_Aplicacao_Tela:

    Traz_Aplicacao_Tela = Err

    Select Case Err

        Case 17244, 17249, 17480, 17482, 17484, 39548

        Case 17246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovContaCorrente.lNumMovto)

        Case 17283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INEXISTENTE1", Err, objMovContaCorrente.iTipo)

        Case 17292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INATIVO", Err, objTiposDeAplicacao.iCodigo)

        Case 17478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVCONTACORRENTE_EXCLUIDO", Err, objMovContaCorrente.iCodConta, objMovContaCorrente.lSequencial)

        Case 17479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_APLICACAO", Err, objMovContaCorrente.lSequencial)

        Case 17481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objMovContaCorrente.iCodConta)

        Case 17483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objMovContaCorrente.iTipoMeioPagto)

        Case 17485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_INEXISTENTE", Err, objMovContaCorrente.iFavorecido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142936)

    End Select

    Exit Function

End Function

Private Sub TaxaPrevista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaPrevista_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dValAplic As Double
Dim dValResgPrev As Double
Dim dPercentual As Double

On Error GoTo Erro_TaxaPrevista_Validate

    'Verifica se taxa prevista está preenchida
    If Len(TaxaPrevista.Text) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TaxaPrevista.Text)
        If lErro <> SUCESSO Then Error 17327

        'Verifica se valor aplicado está preenchido
        If Len(Trim(ValorAplicado.Text)) > 0 Then

           dPercentual = CDbl(TaxaPrevista.Text) / 100

           dValAplic = CDbl(ValorAplicado.Text)

           dValResgPrev = (1 + dPercentual) * dValAplic

           'Coloca o valor resgate previsto na tela
           ValorResgatePrevisto.Text = Format(dValResgPrev, "Fixed")

        ElseIf Len(Trim(ValorResgatePrevisto.Text)) > 0 Then

            dValResgPrev = CDbl(ValorResgatePrevisto.Text)

            dPercentual = CDbl(TaxaPrevista.Text) / 100

            dValAplic = dValResgPrev / (1 + dPercentual)

            'Coloca o valor aplicado na tela
            ValorAplicado.Text = Format(dValAplic, "Fixed")

        End If

    End If

    Exit Sub

Erro_TaxaPrevista_Validate:

    Cancel = True


    Select Case Err

        Case 17327

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142937)

    End Select

    Exit Sub

End Sub

Private Sub TipoAplicacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoAplicacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoAplicacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoAplicacao_Validate

    If Len(Trim(TipoAplicacao.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox Tipo de Aplicacao
    If TipoAplicacao.Text = TipoAplicacao.List(TipoAplicacao.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(TipoAplicacao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17382

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTiposDeAplicacao.iCodigo = iCodigo

        lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
        If lErro <> SUCESSO And lErro <> 15068 Then Error 17293

        'se não encontrou o tipo de aplicação
        If lErro = 15068 Then Error 17294

        If objTiposDeAplicacao.iInativo = TIPOAPLICACAO_INATIVO Then Error 17296

        TipoAplicacao.Text = CStr(objTiposDeAplicacao.iCodigo) & SEPARADOR & objTiposDeAplicacao.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 17383

    Exit Sub

Erro_TipoAplicacao_Validate:

    Cancel = True


    Select Case Err

        Case 17293, 17382

        Case 17294
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODTIPOAPLICACAO_INEXISTENTE", objTiposDeAplicacao.iCodigo)
            If vbMsgRes = vbYes Then
                'Lembrar de manter na tela o numero passado como parametro
                Call Chama_Tela("TipoAplicacao", objTiposDeAplicacao)
            Else
            End If

        Case 17296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INATIVO", Err, objTiposDeAplicacao.iCodigo)

        Case 17383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INEXISTENTE2", Err, TipoAplicacao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142938)

    End Select

    Exit Sub

End Sub

Private Sub TipoMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMeioPagto_Click()

    iAlterado = REGISTRO_ALTERADO
    Call ValidaBotao_Cheque

End Sub

Private Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    'verifica se foi preenchido o TipoMeioPagto
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub

    'Tenta selecionar o TipoMeioPagto com o codigo digitado
    lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17362

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoMeioPagto.iTipo = iCodigo

        'Pesquisa no BD a existencia do tipo passado por parametro
        lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
        If lErro <> SUCESSO And lErro <> 11909 Then Error 17363

        'Se não existir o tipomeiopagto  ==> Erro
        If lErro = 11909 Then Error 17364

        TipoMeioPagto.Text = CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 17365

    Call ValidaBotao_Cheque

    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True


    Select Case Err

        Case 17362, 17363

        Case 17364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)
            
        Case 17365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE1", Err, TipoMeioPagto.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142939)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    DataAplicacao.SetFocus

    If Len(DataAplicacao.ClipText) > 0 Then

        sData = DataAplicacao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 17468

        DataAplicacao.Text = sData

    End If

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 17468

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142940)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    DataAplicacao.SetFocus

    If Len(DataAplicacao.ClipText) > 0 Then

        sData = DataAplicacao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 17469

        DataAplicacao.Text = sData

    End If

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 17469

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142941)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_DownClick

    DataResgatePrevista.SetFocus

    If Len(DataResgatePrevista.ClipText) > 0 Then

        sData = DataResgatePrevista.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 17471

        DataResgatePrevista.Text = sData

    End If

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 17471

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142942)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_UpClick

    DataResgatePrevista.SetFocus

    If Len(DataResgatePrevista.ClipText) > 0 Then

        sData = DataResgatePrevista.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 17470

        DataResgatePrevista.Text = sData

    End If

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 17470

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142943)

    End Select

    Exit Sub

End Sub

Private Sub ValorAplicado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorAplicado_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dValAplic As Double
Dim dValResgPrev As Double
Dim dPercentual As Double

On Error GoTo Erro_ValorAplicado_Validate

    'Verifica se valor aplicado está preenchido
    If Len(ValorAplicado.Text) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorAplicado.Text)
        If lErro <> SUCESSO Then Error 17321

        dValAplic = CDbl(ValorAplicado.Text)

        ValorAplicado.Text = Format(ValorAplicado.Text, "Fixed")

        'Verifica se valor resgate previsto está preenchido
        If Len(ValorResgatePrevisto.Text) > 0 Then

           dValResgPrev = CDbl(ValorResgatePrevisto.Text)

           'Verifica se valor resgate previsto é maior que o valor aplicado
           If dValResgPrev < dValAplic Then Error 17323

           'Calcula a taxa prevista
           dPercentual = (dValResgPrev - dValAplic) / dValAplic * 100

           'Coloca a taxa prevista na tela
           TaxaPrevista.Text = Format(dPercentual, "Fixed")

        Else

           'Verifica se taxa prevista está preenchida
           If Len(TaxaPrevista.Text) > 0 Then

               dPercentual = CDbl(TaxaPrevista.Text) / 100

               'Calcula o valor do resgate previsto
               dValResgPrev = (dPercentual + 1) * dValAplic

               'Coloca o valor do resgate previsto na tela
               ValorResgatePrevisto.Text = Format(dValResgPrev, "Fixed")

           End If

        End If

    End If

    Exit Sub

Erro_ValorAplicado_Validate:

    Cancel = True


    Select Case Err

        Case 17321

        Case 17323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALRESGATE_MENOR_VALAPLICADO", Err, dValResgPrev, dValAplic)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142944)

    End Select

    Exit Sub

End Sub

Private Sub ValorResgatePrevisto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAplicacao As ClassAplicacao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objAplicacao = obj1

    lErro = Traz_Aplicacao_Tela(objAplicacao)
    If lErro <> SUCESSO Then Error 17284

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 17284

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142945)

    End Select

    Exit Sub

End Sub

Private Sub ValorResgatePrevisto_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dValAplic As Double
Dim dValResgPrev As Double
Dim dPercentual As Double

On Error GoTo Erro_ValorResgatePrevisto_Validate

    'Verifica se valor do resgate previsto está preenchido
    If Len(ValorResgatePrevisto.Text) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorResgatePrevisto.Text)
        If lErro <> SUCESSO Then Error 17324

        dValResgPrev = CDbl(ValorResgatePrevisto.Text)

        ValorResgatePrevisto.Text = Format(ValorResgatePrevisto.Text, "Fixed")

        'Verifica se valor aplicado está preenchido
        If Len(Trim(ValorAplicado.Text)) > 0 Then

           dValAplic = CDbl(ValorAplicado.Text)

           'Verifica se valor resgate previsto é maior que o valor aplicado
           If dValResgPrev < dValAplic Then Error 17326

           'Calcula a taxa prevista
           dPercentual = (dValResgPrev - dValAplic) / dValAplic * 100

           'Coloca a taxa na tela
           TaxaPrevista.Text = Format(CDbl(dPercentual), "Fixed")
                      
        ElseIf Len(Trim(TaxaPrevista.Text)) > 0 Then

            dPercentual = CDbl(TaxaPrevista.Text) / 100

            dValResgPrev = CDbl(ValorResgatePrevisto.Text)

            dValAplic = dValResgPrev / (1 + dPercentual)

            'Coloca o valor aplicado na tela
            ValorAplicado.Text = Format(dValAplic, "Fixed")

        End If

    End If

    Exit Sub

Erro_ValorResgatePrevisto_Validate:

    Cancel = True


    Select Case Err

        Case 17324

        Case 17326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALRESGATE_MENOR_VALAPLICADO", Err, dValResgPrev, dValAplic)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142946)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValAplic As Double
Dim dValResgPrev As Double
Dim dtDataAplic As Date
Dim dtDataResg As Date
Dim objAplicacao As New ClassAplicacao
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos essencias da tela foram preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then Error 17340

    If Len(Trim(TipoAplicacao.Text)) = 0 Then Error 17341

    If Len(Trim(DataAplicacao.ClipText)) = 0 Then Error 17342

    dtDataAplic = CDate(DataAplicacao.Text)

    If Len(Trim(ValorAplicado.Text)) = 0 Then Error 17343

    dValAplic = CDbl(ValorAplicado.Text)

    If dValAplic < 0 Then Error 17344

    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 17393

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Error 17345

    If Len(Trim(DataResgatePrevista.ClipText)) <> 0 Then

        dtDataResg = CDate(DataResgatePrevista.Text)

       If dtDataResg < dtDataAplic Then Error 17346

    End If

    If Len(Trim(ValorResgatePrevisto.Text)) <> 0 Then

       dValResgPrev = CDbl(ValorResgatePrevisto.Text)

       If dValResgPrev < dValAplic Then Error 17347

    End If

    'Passa os dados da Tela para objAplicacao e objMovContaCorrente
    lErro = Move_Tela_Memoria(objAplicacao, objMovContaCorrente)
    If lErro <> SUCESSO Then Error 17348

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataAplicacao.Text))
    If lErro <> SUCESSO Then Error 20827

    'Rotina encarregada de gravar a aplicacao
    lErro = CF("Aplicacao_Grava", objAplicacao, objMovContaCorrente, objContabil)
    If lErro <> SUCESSO Then Error 17349

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 17341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_NAO_PREENCHIDO", Err)

        Case 17342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_APLICACAO_NAO_PREENCHIDA", Err)

        Case 17343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORAPLICADO_NAO_INFORMADO", Err)

        Case 17344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, ValorAplicado.Text)

        Case 17345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PREENCHIDA", Err)

        Case 17346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATARESGPREV_MENOR_DATAAPLIC", Err, dtDataResg, dtDataAplic)

        Case 17347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALRESGATE_MENOR_VALAPLICADO", Err, dValResgPrev, dValAplic)

        Case 17348, 17349, 20827

        Case 17393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142947)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Aplicacao() As Long

    Call Limpa_Tela(Me)

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
    CodContaCorrente.Text = ""
    TipoAplicacao.Text = ""
    TipoMeioPagto.Text = ""
    Favorecido.Text = ""
    Historico.Text = ""
    LabelSaldoAplicado.Caption = ""
    DataAplicacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

End Function

Function Move_Tela_Memoria(objAplicacao As ClassAplicacao, objMovContaCorrente As ClassMovContaCorrente) As Long
'Move os dados da tela para memoria.

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objAplicacao e objMovContaCorrente
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    objAplicacao.lCodigo = CLng(Codigo.Text)
    objAplicacao.iTipoAplicacao = Codigo_Extrai(TipoAplicacao.Text)
    objAplicacao.dtDataAplicacao = CDate(DataAplicacao.Text)
    objAplicacao.dValorAplicado = CDbl(ValorAplicado.Text)
    If Len(Trim(LabelSaldoAplicado.Caption)) > 0 Then
        objAplicacao.dSaldoAplicado = CDbl(LabelSaldoAplicado.Caption)
    Else
        objAplicacao.dSaldoAplicado = 0
    End If

    objMovContaCorrente.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)

    If Len(Trim(Numero.Text)) > 0 Then
        objMovContaCorrente.lNumero = CLng(Numero.Text)
    Else
        objMovContaCorrente.lNumero = 0
    End If

    If Len(Trim(Favorecido.Text)) > 0 Then
        objMovContaCorrente.iFavorecido = Codigo_Extrai(Favorecido.Text)
    Else
        objMovContaCorrente.iFavorecido = 0
    End If

    If Len(Trim(NumRefExterna.Text)) > 0 Then
        objMovContaCorrente.sNumRefExterna = NumRefExterna.Text
    End If

    If Len(Trim(Historico.Text)) > 0 Then
        objMovContaCorrente.sHistorico = Historico.Text
    End If

    If Len(Trim(DataResgatePrevista.ClipText)) > 0 Then
        objAplicacao.dtDataResgatePrevista = CDate(DataResgatePrevista.Text)
    Else
        objAplicacao.dtDataResgatePrevista = DATA_NULA
    End If

    If Len(Trim(ValorResgatePrevisto.Text)) > 0 Then
        objAplicacao.dValorResgatePrevisto = CDbl(ValorResgatePrevisto.Text)
    End If

    If Len(Trim(TaxaPrevista.Text)) > 0 Then
        objAplicacao.dTaxaPrevista = CDbl(TaxaPrevista.Text)
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142948)

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

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case CTATIPOAPLICACAO
        
            If Len(Trim(TipoAplicacao.Text)) > 0 Then
                
                objTiposDeAplicacao.iCodigo = Codigo_Extrai(TipoAplicacao.Text)
                
                'Le  a conta no BD
                lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
                If lErro <> SUCESSO And lErro <> 15068 Then gError 64410
                
                'Se não encontrou ---> Erro
                If lErro = 15068 Then gError 64411
                
                If objTiposDeAplicacao.sContaAplicacao <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTiposDeAplicacao.sContaAplicacao, sContaTela)
                    If lErro <> SUCESSO Then gError 64446
                
                Else
                
                    sContaTela = ""
                     
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case CTACONTACORRENTE
            If Len(Trim(CodContaCorrente.Text)) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64412
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64413
                
                If objContasCorrentesInternas.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContasCorrentesInternas.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 64447
                    
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case VALOR1
            If Len(ValorAplicado.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorAplicado.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALORPREVISTO1
            If Len(ValorResgatePrevisto.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorResgatePrevisto.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case CODIGO1
            If Len(Trim(Codigo.ClipText)) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Codigo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
                
        Case CONTACORRENTE1
            If Len(CodContaCorrente.Text) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64414
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64415
                
                objMnemonicoValor.colValor.Add objContasCorrentesInternas.sContaContabil
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                
        Case TIPO1
            If Len(Trim(TipoAplicacao.Text)) > 0 Then
                
                If TipoAplicacao.ListIndex >= 0 Then
                    objTiposDeAplicacao.iCodigo = TipoAplicacao.ItemData(TipoAplicacao.ListIndex)
            
                    'Le  a conta no BD
                    lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
                    If lErro <> SUCESSO And lErro <> 15068 Then gError 64416
                    
                    'Se não encontrou ---> Erro
                    If lErro = 15068 Then gError 64417
                    
                    objMnemonicoValor.colValor.Add objTiposDeAplicacao.sDescricao
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                
        Case FORMA1
            If Len(TipoMeioPagto.Text) > 0 Then
                objMnemonicoValor.colValor.Add TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
                
        Case HISTORICO1
            If Len(Historico.Text) > 0 Then
                objMnemonicoValor.colValor.Add Historico.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
       Case Else
            gError 39551

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
        
        Case 39551
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64410, 64416, 64412, 64414, 64446, 64447
        
        Case 64411, 64417
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INEXISTENTE1", gErr, objTiposDeAplicacao.iCodigo)
        
        Case 64413, 64415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContasCorrentesInternas.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142949)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_APLICACAO_ID
    Set Form_Load_Ocx = Me
    Caption = "Aplicação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Aplicacao"
    
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
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is CodContaCorrente Then
            Call LabelCtaCorrente_Click
        End If
    
    End If

End Sub



Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub LabelCtaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCtaCorrente, Source, X, Y)
End Sub

Private Sub LabelCtaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCtaCorrente, Button, Shift, X, Y)
End Sub

Private Sub LabelSaldoAplicado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSaldoAplicado, Source, X, Y)
End Sub

Private Sub LabelSaldoAplicado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSaldoAplicado, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

'Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label11, Source, X, Y)
'End Sub
'
'Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
'End Sub

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

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Function Move_tela_Cheque(objInfoChequePag As ClassInfoChequePag, dtDataEmissao As Date) As Long

    'Recolhe os dados do cheque
    objInfoChequePag.sFavorecido = Nome_Extrai(Favorecido.Text)
    objInfoChequePag.dValor = StrParaDbl(ValorAplicado.Text)
    objInfoChequePag.lNumRealCheque = StrParaLong(Numero.Text)
    dtDataEmissao = DataAplicacao.Text
    
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

'Fernando subir Função, ela esta na tela AntecipPag
Public Function PreparaImpressao_Cheque(lNumImpressao As Long, objInfoChequePag As ClassInfoChequePag) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long

On Error GoTo Erro_PreparaImpressao_Cheque

    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 80406

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80407

    'obtem sequencial que identifica "geracao" de impressao
    lErro = CF("CPRConfig_ObterNumInt", "NUM_PROX_GERACAO_CHEQUES", lNumImpressao)
    If lErro <> SUCESSO Then gError 80408

    'limpa a tabela
    lErro = Comando_Executar(lComando, "DELETE FROM GeracaoDeCheques WHERE CodGeracao = ?", lNumImpressao)
    If lErro <> AD_SQL_SUCESSO Then gError 80409

    'Insere o novo registro a tabela
    lErro = Comando_Executar(lComando, "INSERT INTO GeracaoDeCheques (CodGeracao, SeqCheque, Favorecido, Valor, NumCheque) VALUES (?,?,?,?,?)", lNumImpressao, 1, objInfoChequePag.sFavorecido, objInfoChequePag.dValor, objInfoChequePag.lNumRealCheque)
    If lErro <> AD_SQL_SUCESSO Then gError 80410

    'Confirma transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 80411

    'Fecha Comando
    lErro = Comando_Fechar(lComando)

    PreparaImpressao_Cheque = SUCESSO
    
    Exit Function

Erro_PreparaImpressao_Cheque:

    PreparaImpressao_Cheque = gErr

    Select Case gErr

        Case 80406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 80407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80408 'Tratado na rotina chamadora
            
        Case 80409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GERACAO_CHEQUES", gErr)

        Case 80410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GERACAO_CHEQUES", gErr)

        Case 80411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142950)

    End Select

    Exit Function

End Function

'Fernando subir Função, ela esta na tela AntecipPag
Function ImprimirCheques(lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long
'chama a impressao de cheques

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String
Dim lErro As Long, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ImprimirCheques

    'a cidade deve vir do endereco da filial que está emitindo, se entrar como EMPRESA_TODA pegar da matriz
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 19466
    
    lErro = objRelatorio.ExecutarDireto("Cheques", "", 0, sLayoutCheques, "NIMPRESSAO", CStr(lNumImpressao), "DEMISSAO", CStr(dtDataEmissao), "TCIDADE", objFilialEmpresa.objEndereco.sCidade)
    If lErro <> SUCESSO Then Error 7431

    ImprimirCheques = SUCESSO

    Exit Function

Erro_ImprimirCheques:

    ImprimirCheques = Err

    Select Case Err

        Case 7431, 19466

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142951)

    End Select

    Exit Function

End Function

Public Sub ValidaBotao_Cheque()
'Apenas verifica o conteúdo da combo TipoMeioPagto e
'se a condição for satisfeita habilita o botão Imprimir

Dim iCodigo As Integer

    'Atribui o valor retornado de Codigo_Extrai a variavel iCodigo
    iCodigo = Codigo_Extrai(TipoMeioPagto.Text)
    
    'Verifica se iCodigo é igual a Constante Cheque
    If iCodigo <> Cheque Then
        'Se for diferente desabilita o botão
        BotaoImprimir.Enabled = False
    Else
        'Se for igual habilita o botão
        BotaoImprimir.Enabled = True
    End If
    
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


'####################################################################
'Inserido por Wagner 26/06/2006
Private Sub LabelFavorecido_Click()

Dim colSelecao As New Collection
Dim objFavorecido As New ClassFavorecidos

    objFavorecido.iCodigo = Codigo_Extrai(Favorecido.Text)
       
    Call Chama_Tela("FavorecidosLista", colSelecao, objFavorecido, objEventoFavorecido)
    
End Sub

Private Sub objEventoFavorecido_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFavorecido As ClassFavorecidos

On Error GoTo Erro_objEventoSaque_evSelecao

    Set objFavorecido = obj1
    
    Favorecido.Text = CStr(objFavorecido.iCodigo)
    
    Call Favorecido_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub
    
Erro_objEventoSaque_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub
'###################################################################

