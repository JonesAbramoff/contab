VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl TransferenciaOcx 
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9390
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4515
      Index           =   2
      Left            =   165
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   9060
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4800
         TabIndex        =   78
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
         Left            =   6300
         TabIndex        =   17
         Top             =   450
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
         Left            =   6300
         TabIndex        =   15
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   19
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
         Height          =   300
         Left            =   7740
         TabIndex        =   16
         Top             =   90
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   25
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
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   34
         Top             =   3435
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   47
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   48
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
            TabIndex        =   49
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
            TabIndex        =   50
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2790
         Left            =   6330
         TabIndex        =   29
         Top             =   1500
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   26
         Top             =   2175
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   27
         Top             =   2565
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
         Left            =   3495
         TabIndex        =   20
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   21
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
         TabIndex        =   24
         Top             =   1905
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
         Left            =   2295
         TabIndex        =   23
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
         TabIndex        =   22
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
         Left            =   1680
         TabIndex        =   35
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
         Left            =   600
         TabIndex        =   14
         Top             =   540
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
         Left            =   5610
         TabIndex        =   13
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
         Left            =   3810
         TabIndex        =   12
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
         Left            =   15
         TabIndex        =   28
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
         Left            =   6330
         TabIndex        =   30
         Top             =   1500
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
         TabIndex        =   31
         Top             =   1500
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
         Left            =   6330
         TabIndex        =   18
         Top             =   720
         Width           =   690
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
         TabIndex        =   51
         Top             =   150
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
         TabIndex        =   52
         Top             =   150
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
         TabIndex        =   53
         Top             =   570
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   54
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   55
         Top             =   3015
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
         TabIndex        =   56
         Top             =   3030
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
         TabIndex        =   57
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
         TabIndex        =   58
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
         TabIndex        =   59
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
         TabIndex        =   60
         Top             =   930
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
         TabIndex        =   61
         Top             =   570
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   62
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   63
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
         Left            =   4230
         TabIndex        =   64
         Top             =   585
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   65
         Top             =   105
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
         TabIndex        =   66
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TransferenciaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TransferenciaOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "TransferenciaOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "TransferenciaOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4845
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   720
      Width           =   9060
      Begin VB.CommandButton BotaoConsultaTransf 
         Caption         =   "Transferências"
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
         Left            =   3405
         TabIndex        =   73
         Top             =   4410
         Width           =   1920
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Principais"
         Height          =   1650
         Left            =   420
         TabIndex        =   32
         Top             =   75
         Width           =   7770
         Begin VB.ComboBox ContaDestino 
            Height          =   315
            Left            =   5115
            TabIndex        =   2
            Top             =   330
            Width           =   1695
         End
         Begin VB.ComboBox ContaOrigem 
            Height          =   315
            Left            =   1815
            TabIndex        =   1
            Top             =   315
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1785
            TabIndex        =   3
            Top             =   810
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown SpinData 
            Height          =   300
            Left            =   2940
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   795
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   315
            Left            =   5100
            TabIndex        =   4
            Top             =   810
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Seq. Destino:"
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
            Left            =   3885
            TabIndex        =   77
            Top             =   1305
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Seq. Origem:"
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
            Left            =   585
            TabIndex        =   76
            Top             =   1305
            Width           =   1110
         End
         Begin VB.Label LabelSeqDestino 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5100
            TabIndex        =   75
            Top             =   1275
            Width           =   1170
         End
         Begin VB.Label LabelSeqOrigem 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1785
            TabIndex        =   74
            Top             =   1275
            Width           =   1170
         End
         Begin VB.Label LblContaDestino 
            AutoSize        =   -1  'True
            Caption         =   "Conta Destino:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   375
            Width           =   1275
         End
         Begin VB.Label LblContaOrigem 
            AutoSize        =   -1  'True
            Caption         =   "Conta Origem:"
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
            Left            =   495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   375
            Width           =   1215
         End
         Begin VB.Label Label15 
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
            Left            =   1215
            TabIndex        =   41
            Top             =   855
            Width           =   480
         End
         Begin VB.Label Label14 
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
            Left            =   4545
            TabIndex        =   42
            Top             =   840
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Operação"
         Height          =   1215
         Left            =   435
         TabIndex        =   38
         Top             =   1845
         Width           =   7785
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
            Height          =   510
            Left            =   5985
            Picture         =   "TransferenciaOcx.ctx":0994
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   495
            Width           =   1680
         End
         Begin VB.ComboBox Favorecido 
            Height          =   315
            Left            =   1740
            TabIndex        =   7
            Top             =   765
            Width           =   4020
         End
         Begin VB.ComboBox TipoMeioPagto 
            Height          =   315
            Left            =   1755
            TabIndex        =   5
            Top             =   255
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   300
            Left            =   4575
            TabIndex        =   6
            Top             =   255
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##########"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Left            =   630
            TabIndex        =   67
            Top             =   795
            Width           =   1020
         End
         Begin VB.Label Label11 
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
            Left            =   3780
            TabIndex        =   43
            Top             =   300
            Width           =   720
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
            Left            =   1020
            TabIndex        =   44
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Complemento"
         Height          =   1110
         Left            =   420
         TabIndex        =   33
         Top             =   3165
         Width           =   7785
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1815
            TabIndex        =   9
            Top             =   225
            Width           =   5085
         End
         Begin MSMask.MaskEdBox NumRefExterna 
            Height          =   300
            Left            =   1800
            TabIndex        =   10
            Top             =   675
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
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
            Left            =   870
            TabIndex        =   45
            Top             =   270
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
            Left            =   495
            TabIndex        =   46
            Top             =   720
            Width           =   1185
         End
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5250
      Left            =   120
      TabIndex        =   37
      Top             =   390
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   9260
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
Attribute VB_Name = "TransferenciaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Responsavel: Mario
'Revisado em 20/8/98

'1 - ?? A parte de contabilização não está funcionando;

Option Explicit
'inicio contabilidade

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'Mnemônicos
Private Const CONTAORIGEM1 As String = "Conta_Origem"
Private Const CONTADESTINO1 As String = "Conta_Destino"
Private Const DATA1 As String = "Data"
Private Const VALOR1 As String = "Valor_Transf"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const NUMERO1 As String = "Numero"
Private Const HISTORICO1 As String = "Historico"
Private Const DOC_EXTERNO As String = "Doc_Externo"
Private Const CTACONTAORIGEM As String = "Cta_Conta_Origem"
Private Const CTACONTADESTINO As String = "Cta_Conta_Destino"

'fim da contabilidade

Public iAlterado As Integer
Dim iFrameAtual As Integer
Public gsBuffer As String

Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1
Private WithEvents objEventoContaCorrenteInt1 As AdmEvento
Attribute objEventoContaCorrenteInt1.VB_VarHelpID = -1
Private WithEvents objEventoTransferencia As AdmEvento
Attribute objEventoTransferencia.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2

Private Sub BotaoConsultaTransf_Click()

Dim lErro As Long
Dim objTransferencia As New ClassMovContaCorrente
Dim colSelecao As New Collection

On Error GoTo Erro_TrazTranferencia_Click

    'Chama a Tela
    Call Chama_Tela("TransferenciaLista", colSelecao, objTransferencia, objEventoTransferencia)

    Exit Sub

Erro_TrazTranferencia_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175529)

    End Select

    Exit Sub
End Sub

'Maristela(Antes)
Private Sub BotaoExcluir_Click()

Dim lErro  As Long

Dim vbMsgRes As VbMsgBoxResult
Dim objTransferencia As New ClassTransferencia

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Carrega o Obj para a exclusão
    lErro = Move_Tela_Memoria(objTransferencia)
    If lErro <> SUCESSO Then gError 90499
    
    If objTransferencia.lSeqOrigem = 0 And objTransferencia.lSeqDestino = 0 Then gError 90503
    
    'Pede a confirmação da exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TRANSFERENCIA", objTransferencia.iCodContaOrigem, objTransferencia.iCodContaDestino)

    If vbMsgRes = vbYes Then
        
        'Chama a rotina de exclusao
        lErro = CF("MovCCI_Exclui_Transferencia", objTransferencia, objContabil)
        If lErro <> SUCESSO Then gError 90500
        
        Call Limpa_Tela_Transferencia

        iAlterado = 0
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 90499, 90500
        
        Case 90503
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSFERENCIA_INEXISTENTE1", gErr, objTransferencia.iCodContaOrigem, objTransferencia.iCodContaDestino)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175530)

    End Select

    Exit Sub

End Sub
'Maristela(Depois)

Private Sub ContaDestino_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaDestino_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaDestino_GotFocus()
    
    gsBuffer = ContaDestino.Text

End Sub

Private Sub ContaOrigem_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaOrigem_GotFocus()
    
    gsBuffer = ContaOrigem.Text

End Sub

Private Sub Data_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        Opcao.Visible = False
    End If
    
    Set objEventoContaCorrenteInt = New AdmEvento
    Set objEventoContaCorrenteInt1 = New AdmEvento
    Set objEventoTransferencia = New AdmEvento

    
    'Carrega a combo dos Tipos de meio de Pagamento
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then gError 18122

    'permite fazer transferencia entre as contas de toda a empresa e não somente da filial em questão
    iFilialEmpresa = giFilialEmpresa
    giFilialEmpresa = EMPRESA_TODA

    'Carrega a combo com os codigos e nomes das contas correntes
    lErro = Carrega_CodContaCorrente(ContaOrigem)
    If lErro <> SUCESSO Then gError 18123

    lErro = Carrega_CodContaCorrente(ContaDestino)
    If lErro <> SUCESSO Then gError 18128

    giFilialEmpresa = iFilialEmpresa

    'carrega a combo de historico
    lErro = Carrega_Historico()
    If lErro <> SUCESSO Then gError 18124
    
    'Lê os favorecidos com codigo e o nome existentes no BD e carrega na ComboBox
    lErro = Carrega_Favorecidos()
    If lErro <> SUCESSO Then gError 87032

    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    iFrameAtual = 1
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_TESOURARIA)
    If lErro <> SUCESSO Then gError 39610
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 18122, 18123, 18124, 18128, 18244, 39610, 87032
        
        Case Else
        
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Private Function Carrega_TipoMeioPagto() As Long
'Carrega na Combo TipoMeioPagto os tipo de meio de pagamento ativos

Dim lErro As Long
Dim colTipoMeioPagto As New Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Le todos os tipo de pagamento
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 18125

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

        Case 18125

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175531)

    End Select

    Exit Function

End Function

Private Function Carrega_CodContaCorrente(objComboBox As Object) As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_CodContaCorrente
    
   'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoDescricao)
    If lErro <> SUCESSO Then Error 18126

    For Each objCodigoNome In colCodigoDescricao

        'Insere na combo de contas correntes
        objComboBox.AddItem objCodigoNome.iCodigo & SEPARADOR & objCodigoNome.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 18126

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175532)

    End Select

    Exit Function

End Function

Private Function Carrega_Historico() As Long
'Carrega a combo de historicos com os historicos da tabela "HistPadraMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'le o Codigo e a descricao de todos os historicos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then Error 18127

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

        Case 18127

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175533)

    End Select

End Function

'Maristela(Antes)
Function Trata_Parametros(Optional objTransferencia As ClassMovContaCorrente) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum movimento foi passado por parametro
    If Not (objTransferencia Is Nothing) Then
            
        'Traz os dados do movimento passado por parametro
        lErro = Traz_Transferencia_Tela(objTransferencia)
        If lErro <> SUCESSO Then gError 90462
    Else
        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 90462
              Call Limpa_Tela_Transferencia
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175534)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function
'Maristela(depois)

'Maristela(antes)
Private Function Traz_Transferencia_Tela(objMovtoCta As ClassMovContaCorrente) As Long
'Coloca na Tela os dados da Transferencia passado por parametro

Dim lErro As Long
Dim iIndice As Integer
Dim objTransferencia As New ClassTransferencia
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim objFavorecidos As New ClassFavorecidos
Dim lNumMovtoSaida As Long

On Error GoTo Erro_Traz_Transferencia_Tela

    objTransferencia.lNumIntDoc = objMovtoCta.lNumMovto
    
    'Le a Conta de Origem(Saida), a partir da Conta de Destino(Entrada) ou vice versa
    'Wagner - Não necessariamente objMovtoCta vem com a movimentação de entrada
    'Logo a função abaixo vai pegar o movimento faltante, quer seja ele de entrada ou saída
    lErro = Transferencia_Le_MovtoSaida(objTransferencia)
    If lErro <> SUCESSO And lErro <> 90476 Then gError 90463
    
    If lErro = 90476 Then gError 90482
    
    'Se pegou a entrada completa a saída
    If objTransferencia.iCodContaOrigem = 0 Then
        objTransferencia.iCodContaOrigem = objMovtoCta.iCodConta
        objTransferencia.lSeqOrigem = objMovtoCta.lSequencial
        lNumMovtoSaida = objMovtoCta.lNumMovto
    Else
        objTransferencia.iCodContaDestino = objMovtoCta.iCodConta
        objTransferencia.lSeqDestino = objMovtoCta.lSequencial
        lNumMovtoSaida = objTransferencia.lNumIntDoc
    End If
    
    'Passa os dados para a Tela
    Data.Text = Format(objMovtoCta.dtDataMovimento, "dd/MM/yy")
    Valor.Text = (objMovtoCta.dValor)
    If objMovtoCta.lNumero <> 0 Then
        Numero.Text = CStr(objMovtoCta.lNumero)
    End If
    Historico.Text = objMovtoCta.sHistorico
    NumRefExterna.Text = objMovtoCta.sNumRefExterna
                  
    'Verifica se o TiPoMeioPagot existe
    objTipoMeioPagto.iTipo = objMovtoCta.iTipoMeioPagto

    lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
    If lErro <> SUCESSO And lErro <> 11909 Then gError 90468

    If lErro = 11909 Then gError 90469

    TipoMeioPagto.Text = CStr(objMovtoCta.iTipoMeioPagto) & SEPARADOR & objTipoMeioPagto.sDescricao

    If objMovtoCta.iFavorecido <> 0 Then

        objFavorecidos.iCodigo = objMovtoCta.iFavorecido

        'Verifica se o favorecido existe
        lErro = CF("Favorecido_Le", objFavorecidos)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 90470

        If lErro = 11807 Then gError 90471

        Favorecido.Text = CStr(objFavorecidos.iCodigo) & SEPARADOR & objFavorecidos.sNome

    End If
    
    'Verifica se a conta corrente(Destino) existe
    lErro = CF("ContaCorrenteInt_Le", objTransferencia.iCodContaDestino, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 90464
    
    If lErro = 11807 Then gError 90465
    
    For iIndice = 0 To ContaDestino.ListCount - 1
        If ContaDestino.List(iIndice) = CStr(objTransferencia.iCodContaDestino) & SEPARADOR & objContaCorrenteInt.sNomeReduzido Then
            ContaDestino.ListIndex = iIndice
            Exit For
        End If
    Next

    'Verifica se a conta corrente(Origem) existe
    lErro = CF("ContaCorrenteInt_Le", objTransferencia.iCodContaOrigem, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 90466

    If lErro = 11807 Then gError 90467

    For iIndice = 0 To ContaOrigem.ListCount - 1
        If ContaOrigem.List(iIndice) = CStr(objTransferencia.iCodContaOrigem) & SEPARADOR & objContaCorrenteInt.sNomeReduzido Then
            ContaOrigem.ListIndex = iIndice
            Exit For
        End If
    Next
    
    LabelSeqOrigem.Caption = objTransferencia.lSeqOrigem
    LabelSeqDestino.Caption = objTransferencia.lSeqDestino
    
    objMovtoCta.lNumMovto = lNumMovtoSaida
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovtoCta.lNumMovto)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 90472
    
    Traz_Transferencia_Tela = SUCESSO

    Exit Function

Erro_Traz_Transferencia_Tela:

    Traz_Transferencia_Tela = gErr

    Select Case gErr

        Case 90463, 90470, 90472, 90464, 90468, 90466
        
        Case 90482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", gErr)
        
        Case 90471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_INEXISTENTE", gErr, objMovtoCta.iFavorecido)
           
        Case 90465 '???
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTEDESTINO_INEXISTENTE", gErr, objMovtoCta.iCodConta)
        
        Case 90467 '???
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTEORIGEM_INEXISTENTE", gErr, objTransferencia.iCodContaOrigem)
        
        Case 90469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", gErr, objMovtoCta.iTipoMeioPagto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175535)

    End Select

    Exit Function

End Function
'Maristela(depois)

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_UnLoad(Cancel As Integer)
    
    Set objEventoContaCorrenteInt = Nothing
    Set objEventoContaCorrenteInt1 = Nothing

    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    Set objEventoTransferencia = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing

End Sub

Private Sub Data_GotFocus()
    
    gsBuffer = Data.Text
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Favorecido_GotFocus()

    gsBuffer = Favorecido.Text

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO
 
End Sub

Private Sub Historico_Click()
       
    iAlterado = REGISTRO_ALTERADO
   
End Sub

Private Sub Historico_GotFocus()
    
    gsBuffer = Historico.Text
    
End Sub

Private Sub LblContaDestino_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub LblContaDestino_Click()

Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim colSelecao As Collection

    If Len(ContaDestino.Text) = 0 Then
        objContaCorrenteInt.iCodigo = 0
    Else
        objContaCorrenteInt.iCodigo = Codigo_Extrai(ContaDestino.Text)
    End If

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContaCorrenteInt, objEventoContaCorrenteInt1)

End Sub

Private Sub LblContaOrigem_Click()

Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim colSelecao As Collection

    If Len(ContaOrigem.Text) = 0 Then
        objContaCorrenteInt.iCodigo = 0
    Else
        objContaCorrenteInt.iCodigo = Codigo_Extrai(ContaOrigem.Text)
    End If

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContaCorrenteInt, objEventoContaCorrenteInt)

End Sub

Private Sub Numero_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()
    
    gsBuffer = Numero.Text
    
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)
   
   If gsBuffer <> Numero.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

End Sub

Private Sub NumRefExterna_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumRefExterna_GotFocus()

    gsBuffer = NumRefExterna.Text
    
End Sub

Private Sub NumRefExterna_Validate(Cancel As Boolean)
    
    If gsBuffer <> NumRefExterna.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

End Sub

Private Sub objEventoContaCorrenteInt1_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1

    ContaDestino.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    ContaDestino.SetFocus

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoTransferencia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTransferencia As ClassMovContaCorrente

On Error GoTo Erro_objEventoTranferencia_evSelecao

    Set objTransferencia = obj1
    
    lErro = Traz_Transferencia_Tela(objTransferencia)
    If lErro <> SUCESSO Then gError 90504
           
    Me.Show

    Exit Sub

Erro_objEventoTranferencia_evSelecao:

    Select Case gErr
        
        Case 90504
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175536)

    End Select

    Exit Sub

End Sub

Private Sub TipoMeioPagto_Click()
        
    iAlterado = REGISTRO_ALTERADO
    
    Call ValidaBotao_Cheque

End Sub

Private Sub TipoMeioPagto_GotFocus()
    
    gsBuffer = TipoMeioPagto.Text
    
End Sub

Private Sub Valor_Change()
    
     iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1

    ContaOrigem.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    ContaOrigem.SetFocus

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub ContaOrigem_Click()
        
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If gsBuffer <> Data.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    'verifica se a data está vazia
    If Len(Data.ClipText) = 0 Then Exit Sub

    'verifica se a data é válida
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then Error 18130

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err
        
        Case 18130

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175537)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_DownClick()
Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_DownClick

    Data.SetFocus
    
    If gsBuffer <> Data.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        'Diminui a data
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 18131

        Data.PromptInclude = False
        Data.Text = sData
        Data.PromptInclude = True

    End If

    Exit Sub

Erro_SpinData_DownClick:

    Select Case Err

        Case 18131

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175538)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_UpClick

    Data.SetFocus
    
    If gsBuffer <> Data.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If
   
    'verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text
        
        'Aumenta a data
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 18132

        Data.PromptInclude = False
        Data.Text = sData
        Data.PromptInclude = True

    End If

    Exit Sub

Erro_SpinData_UpClick:

    Select Case Err

        Case 18132

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175539)

    End Select

    Exit Sub

End Sub

Private Sub Valor_GotFocus()
    
    gsBuffer = Format(Valor.Text, "Fixed")

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate


    If gsBuffer <> Format(Valor.Text, "Fixed") Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    'Verifica se há um valor digitado
    If Len(Trim(Valor.Text)) > 0 Then
    
        'Critica o valor digitado
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 18133

        Valor.Text = Format(Valor.Text, "Fixed")
        
    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 18133

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175540)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iTamanho As Integer
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Historico_Validate

    If gsBuffer <> Historico.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    'Verifica o tamanho do texto em historico
    iTamanho = Len(Trim(Historico.Text))

    If iTamanho = 0 Then Exit Sub
    
    'Verifica se é maior que o tamanho maximo
    If iTamanho > 50 Then Error 40733
    
    'Verifica se o que foi digitado é numerico
    If IsNumeric(Trim(Historico.Text)) Then
        
        'verifica se é inteiro o codigo passado
        lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
        If lErro <> SUCESSO Then Error 40734
        
        'preenche o objeto
        objHistMovCta.iCodigo = CInt(Trim(Historico.Text))
        
        'verifica a existencia dele pelo codigo passado
        lErro = CF("HistMovCta_Le", objHistMovCta)
        If lErro <> SUCESSO And lErro <> 15011 Then Error 40735
        
        'se não existir -----> Erro
        If lErro = 15011 Then Error 40741
        
        Historico.Text = objHistMovCta.sDescricao
                       
    End If
       
    Exit Sub

Erro_Historico_Validate:

    Cancel = True


    Select Case Err
    
        Case 40733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", Err)
          
        Case 40734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, Historico.Text)
        
        Case 40735
        
        Case 40741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175541)

    End Select

    Exit Sub

End Sub

Private Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    If gsBuffer <> TipoMeioPagto.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    'verifica se foi preenchido o TipoMeioPagto
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub

    'Tenta selecionar o TipoMeioPagto com o codigo digitado
    lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18136

    If lErro = 6730 Then

        objTipoMeioPagto.iTipo = iCodigo
    
        'Pesquisa no BD a existencia do tipo passado por parametro
        lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
        If lErro <> SUCESSO And lErro <> 11909 Then Error 18137
        
        'Se não existir ==> Erro
        If lErro = 11909 Then Error 18138
            
        'Coloca o dado na tela
        TipoMeioPagto.Text = CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
    
    End If
    
    Call ValidaBotao_Cheque
    
    If lErro = 6731 Then Error 18135

    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True



    Select Case Err

        Case 18135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE1", Err, TipoMeioPagto.Text)
         
        Case 18138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)
            
        Case 18136, 18137
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175542)

    End Select

    Exit Sub

End Sub

Private Sub ContaOrigem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaOrigem_Validate

    If gsBuffer <> ContaOrigem.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    If Len(Trim(ContaOrigem.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox ContaOrigem
    If ContaOrigem.Text = ContaOrigem.List(ContaOrigem.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(ContaOrigem, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18140

    If lErro = 6730 Then
    
        objContaCorrenteInt.iCodigo = iCodigo
        
        'Tenta ler conta com esse codigo no BD
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 18141
        
        'Se não encontrou a Conta Corrente --> Erro
        If lErro = 11807 Then Error 18142
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43537
        
        End If
        
        ContaOrigem.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    
    End If

    If lErro = 6731 Then Error 18139

    Exit Sub

Erro_ContaOrigem_Validate:

    Cancel = True


    Select Case Err

        Case 18139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaOrigem.Text)
        
        Case 18140, 18141

        Case 18142
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTAORIGEM_INEXISTENTE", ContaOrigem.Text)
        
            If vbMsgRes = vbYes Then
                'Lembrar de manter na tela o numero passado como parametro
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If
            
        Case 43537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaOrigem.Text, giFilialEmpresa)
            
        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175543)

    End Select

    Exit Sub

End Sub

Private Sub ContaDestino_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaDestino_Validate

    If gsBuffer <> ContaDestino.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    If Len(Trim(ContaDestino.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox ContaDestino
    If ContaDestino.Text = ContaDestino.List(ContaDestino.ListIndex) Then Exit Sub

    'Tenta encontrara e selecionar o item na combo
    lErro = Combo_Seleciona(ContaDestino, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18144

    If lErro = 6730 Then
    
        objContaCorrenteInt.iCodigo = iCodigo
        'tenta ler a conta com esse codigo no BD
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 18145
        
        'Se não encontrou
        If lErro = 11807 Then Error 18146
    
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43538
        
        End If
        
            ContaDestino.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
                
    End If
    
    If lErro = 6731 Then Error 18143

    Exit Sub

Erro_ContaDestino_Validate:

    Cancel = True


    Select Case Err

        Case 18143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaDestino.Text)
        
        Case 18144, 18145

        Case 18146
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTADESTINO_INEXISTENTE", objContaCorrenteInt.iCodigo)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If
                    
        Case 43538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaDestino.Text, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175544)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a rotina de gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 18147

    Call Limpa_Tela_Transferencia

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 18147

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175545)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_Transferencia() As Long
'Limpa todos os campos datela  e coloca em data a data atual do sistema

    Call Limpa_Tela(Me)

    ContaOrigem.Text = ""
    ContaDestino.Text = ""
    LabelSeqOrigem.Caption = ""
    LabelSeqDestino.Caption = ""
    TipoMeioPagto.Text = ""
    Favorecido.Text = ""
    Historico.Text = ""
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    
    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 18148

    'Limpa a tela
    Call Limpa_Tela_Transferencia

    BotaoImprimir.Enabled = False

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 18148

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175546)

    End Select

    Exit Sub
End Sub

Public Function Gravar_Registro() As Long

Dim objTransferencia As New ClassTransferencia
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a conta origem foi preenchida
    If Len(Trim(ContaOrigem.Text)) = 0 Then Error 18149
    
    'Verifica se a conta destino foi preenchida
    If Len(Trim(ContaDestino.Text)) = 0 Then Error 18150
    
    'Verifica se a data foi preenchida
    If Len(Data.ClipText) = 0 Then Error 18151
    
    If Len(Trim(Valor.ClipText)) = 0 Then Error 18236
    
    'Verifica se o valor é válido
    lErro = Valor_Positivo_Critica(Valor.Text)
    If lErro <> SUCESSO Then Error 18152
    
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 18238
    
    'Move os dados da tela para a variável objtransferencia
    lErro = Move_Tela_Memoria(objTransferencia)
    If lErro <> SUCESSO Then Error 18153
    
    'Verifica se  a conta origem e igual a conta Destino
    If objTransferencia.iCodContaOrigem = objTransferencia.iCodContaDestino Then Error 18199
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then Error 20836
    
    'Chama a rotina que grava o movimento no BD
    lErro = CF("MovCCI_Grava_Transferencia", objTransferencia, objContabil)
    If lErro <> SUCESSO Then Error 18154
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 18149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAORIGEM_NAO_DIGITADA", Err)
        
        Case 18150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTADESTINO_NAO_DIGITADA", Err)
        
        Case 18151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
        
        Case 18152, 18153, 18154, 20836
        
        Case 18199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCI_IGUAIS", Err)
        
        Case 18236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", Err)
        
        Case 18238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175547)
    
    End Select
    
    Exit Function
    
End Function

Function Move_Tela_Memoria(objTransferencia As ClassTransferencia) As Long
'Move os dados da tela para memória

    'Move os dados da tela para objTransferencia
    objTransferencia.iCodContaOrigem = Codigo_Extrai(ContaOrigem.Text)

    objTransferencia.iCodContaDestino = Codigo_Extrai(ContaDestino.Text)

    objTransferencia.dtData = CDate(Data.Text)
    objTransferencia.dValor = CDbl(Valor.Text)

    objTransferencia.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)

    If Len(Trim(Numero.Text)) > 0 Then objTransferencia.lNumero = CLng(Numero.Text)
    
    objTransferencia.sHistorico = Historico.Text
    objTransferencia.sNumRefExterna = NumRefExterna.Text
    objTransferencia.iFavorecido = Codigo_Extrai(Favorecido.Text)
    objTransferencia.lSeqOrigem = StrParaLong(LabelSeqOrigem.Caption)
    objTransferencia.lSeqDestino = StrParaLong(LabelSeqDestino.Caption)
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
End Function

Private Sub Opcao_Click()

    'Se Frame selecionado não for o atual esconde o frame atual, mostra o novo.
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
                Parent.HelpContextID = IDH_TRANSFERENCIA_ID
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_TRANSFERENCIA_CONTABILIZACAO
                        
        End Select
    
    End If

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 36242

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 18591

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 18591
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 36242

    End Select

    Exit Function

End Function

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

'****
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
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case CTACONTADESTINO
            If Len(ContaDestino.Text) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(ContaDestino.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64434
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64435
                
                If objContasCorrentesInternas.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContasCorrentesInternas.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 64442
                    
                Else
                
                    sContaTela = ""
                    
                End If
                                    
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case CTACONTAORIGEM
            If Len(ContaOrigem.Text) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(ContaOrigem.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64436
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64437
                
                If objContasCorrentesInternas.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContasCorrentesInternas.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 64443
                    
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
            
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case DOC_EXTERNO
            If Len(Trim(NumRefExterna.Text)) > 0 Then
                objMnemonicoValor.colValor.Add NumRefExterna.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case VALOR1
            If Len(Valor.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Valor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CONTAORIGEM1
            If Len(ContaOrigem.Text) Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(ContaOrigem.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64438
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64439
                
                objMnemonicoValor.colValor.Add objContasCorrentesInternas.sNomeReduzido
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case CONTADESTINO1
            If Len(ContaDestino.Text) > 0 Then
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(ContaDestino.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64440
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64441
                
                objMnemonicoValor.colValor.Add objContasCorrentesInternas.sNomeReduzido
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case DATA1
            If Len(Data.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case FORMA1
            If Len(TipoMeioPagto) > 0 Then
                objMnemonicoValor.colValor.Add TipoMeioPagto.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case NUMERO1
            If Len(Numero.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Numero.Text)
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
            gError 39611
            
    End Select
    
    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
    
    Select Case gErr
    
        Case 39611
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64434, 64436, 64438, 64440, 64442, 64443
        
        Case 64435, 64437, 64439, 64441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContasCorrentesInternas.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175548)
            
    End Select
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TRANSFERENCIA_ID
    Set Form_Load_Ocx = Me
    Caption = "Transferência"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Transferencia"
    
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
        
        If Me.ActiveControl Is ContaOrigem Then
            Call LblContaOrigem_Click
        ElseIf Me.ActiveControl Is ContaDestino Then
            Call LblContaDestino_Click
        End If
    
    End If
    
End Sub




Private Sub LblContaDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblContaDestino, Source, X, Y)
End Sub

Private Sub LblContaDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblContaDestino, Button, Shift, X, Y)
End Sub

Private Sub LblContaOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblContaOrigem, Source, X, Y)
End Sub

Private Sub LblContaOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblContaOrigem, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
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

    If gsBuffer <> Favorecido.Text Then
        LabelSeqOrigem.Caption = ""
        LabelSeqDestino.Caption = ""
    End If

    'Verifica se foi preenchido o Favorecido
    If Len(Trim(Favorecido.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox Favorecido
    If Favorecido.Text = Favorecido.List(Favorecido.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(Favorecido, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 87033

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objFavorecido.iCodigo = iCodigo

        lErro = CF("Favorecido_Le", objFavorecido)
        If lErro <> SUCESSO And lErro <> 17015 Then gError 87034

        'Não encontrou o Favorecido no BD
        If lErro = 17015 Then gError 87035

        'Encontrou o Favorecido no BD, coloca no Text da Combo
        Favorecido.Text = CStr(objFavorecido.iCodigo) & SEPARADOR & objFavorecido.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 87036

    Exit Sub

Erro_Favorecido_Validate:

    Cancel = True


    Select Case gErr

        Case 87033, 87034

        Case 87035
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FAVORECIDO_INEXISTENTE", objFavorecido.iCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Favorecidos
                Call Chama_Tela("Favorecidos", objFavorecido)
            Else
            End If

        Case 87036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_INEXISTENTE1", gErr, Favorecido.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175549)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Favorecidos() As Long
'Carrega os favorecidos ativos na combo de Favorecidos

Dim lErro As Long
Dim objFavorecidos As ClassFavorecidos
Dim colFavorecidos As New Collection

On Error GoTo Erro_Carrega_Favorecidos

    'Le todos os favorecidos
    lErro = CF("Favorecidos_Le_Todos", colFavorecidos)
    If lErro <> SUCESSO Then gError 87036

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

    Carrega_Favorecidos = gErr

    Select Case gErr

        Case 87036

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175550)

    End Select

    Exit Function

End Function

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim iCodigoOrigem As Integer
Dim iCodigoDestino As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objBanco As New ClassBanco
Dim sLayoutCheque As String
Dim objInfoChequePag As New ClassInfoChequePag
Dim dtDataEmissao As Date
Dim lNumImpressao As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(ContaOrigem.Text)) = 0 Then gError 87038
    If Len(Trim(ContaDestino.Text)) = 0 Then gError 87039
    If Len(Trim(Data.ClipText)) = 0 Then gError 87040
    If Len(Trim(Valor.Text)) = 0 Then gError 87041
    'If Len(Trim(TipoMeioPagto.Text)) = 0 Then gError 87042
    If Len(Trim(Favorecido.Text)) = 0 Then gError 87043
    
    'Retira o código da combo e passa para iCodigo
    iCodigoOrigem = Codigo_Extrai(ContaOrigem.Text)
    iCodigoDestino = Codigo_Extrai(ContaDestino.Text)
    
    'Le a Conta Corrente a partir de iCodigoOrigem passado como parâmetro
    lErro = CF("ContaCorrenteInt_Le", iCodigoOrigem, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 87044

    'Caso a Conta Corrente de Origem não tiver sido encontrada dispara erro
    If lErro = 11807 Then gError 87046

    'Caso a Conta Corrente de Origem não for bancária dispara erro
    If objContaCorrenteInt.iCodBanco = 0 Then gError 87047
    
    'Atribui o valor retornado de objContaCorrenteInt.iCodBanco a objBanco.iCodBanco
    objBanco.iCodBanco = objContaCorrenteInt.iCodBanco

    'Le a Conta Corrente a partir de iCodigoDestino passado como parâmetro
    lErro = CF("ContaCorrenteInt_Le", iCodigoDestino, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 87045

    'Caso a Conta Corrente não tiver sido encontrada dispara erro
    If lErro = 11807 Then gError 87048

    'Caso a Conta Corrente não for bancária dispara erro
    If objContaCorrenteInt.iCodBanco = 0 Then gError 87049
                                                                         
    'Le o Banco a partir de objBanco.iCodBanco
    lErro = CF("Banco_Le", objBanco)
    If lErro <> SUCESSO And lErro <> 16091 Then gError 87050
                          
    'Caso o banco não tiver sido encontrado dispara erro
    If lErro = 16091 Then gError 87051
                                
    'Atribui retorno de objBanco.sLayoutCheque a variavel sLayoutCheque
    sLayoutCheque = objBanco.sLayoutCheque
                                                                                                                                                                                                                          
    'Recolhe os dados do cheque da tela para objInfoChequePag
    Call Move_tela_Cheque(objInfoChequePag, dtDataEmissao)

    'Chama a função que prepara a impressão do cheque
    lErro = CF("PreparaImpressao_Cheque", lNumImpressao, objInfoChequePag)
    If lErro <> SUCESSO Then gError 87052
    
    'Chama a função responsável pela impressão do cheque
    lErro = ImprimirCheques(lNumImpressao, sLayoutCheque, dtDataEmissao)
    If lErro <> SUCESSO Then gError 87053
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 87044, 87045, 87050, 87052, 87053

        Case 87038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAORIGEM_NAO_DIGITADA", gErr)

        Case 87039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTADESTINO_NAO_DIGITADA", gErr)

        Case 87040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
                                           
        Case 87041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)

        Case 87043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_NAO_PREENCHIDO", gErr)
        
        Case 87046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTEORIGEM_NAO_ENCONTRADA", gErr, ContaOrigem.Text)
        
        Case 87047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTEORIGEM_NAO_BANCARIA", gErr)

        Case 87048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTEDESTINO_NAO_ENCONTRADA", gErr, ContaDestino.Text)
                
        Case 87049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTEDESTINO_NAO_BANCARIA", gErr)
        
        Case 87051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, objBanco.iCodBanco)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175551)

    End Select

    Exit Sub

End Sub

Function Move_tela_Cheque(objInfoChequePag As ClassInfoChequePag, dtDataEmissao As Date) As Long

    'Recolhe os dados do cheque
    objInfoChequePag.sFavorecido = Nome_Extrai(Favorecido.Text)
    objInfoChequePag.dValor = StrParaDbl(Valor.Text)
    objInfoChequePag.lNumRealCheque = StrParaLong(Numero.Text)
    dtDataEmissao = Data.Text
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175552)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175553)

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

Function Transferencia_Le_MovtoSaida(objTransferenciaSaida As ClassTransferencia) As Long
'Le o movimento de conta corrente passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim iCodContaOrigem As Integer
Dim lSeqOrigem As Long
Dim lNumMovto As Long
Dim iTipo As Integer

On Error GoTo Erro_Transferencia_Le_MovtoSaida

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90473
        
    'Le o movimento com o codigo passado como Parametro
    lErro = Comando_Executar(lComando, "SELECT CodConta, Sequencial, NumMovto, Tipo FROM MovimentosContaCorrente WHERE NumMovtoTransf = ?", iCodContaOrigem, lSeqOrigem, lNumMovto, iTipo, objTransferenciaSaida.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 90474

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90475

    If lErro = AD_SQL_SEM_DADOS Then gError 90476
    
    'Armazena em objTransferenciaSaida a Conta de Origem
    If iTipo = MOVCCI_SAIDA_TRANSFERENCIA Then
        objTransferenciaSaida.iCodContaDestino = 0
        objTransferenciaSaida.lSeqDestino = 0
        objTransferenciaSaida.iCodContaOrigem = iCodContaOrigem
        objTransferenciaSaida.lSeqOrigem = lSeqOrigem
    Else
        objTransferenciaSaida.iCodContaDestino = iCodContaOrigem
        objTransferenciaSaida.lSeqDestino = lSeqOrigem
        objTransferenciaSaida.iCodContaOrigem = 0
        objTransferenciaSaida.lSeqOrigem = 0
    End If
    objTransferenciaSaida.lNumIntDoc = lNumMovto
    
    Call Comando_Fechar(lComando)

    Transferencia_Le_MovtoSaida = SUCESSO

    Exit Function

Erro_Transferencia_Le_MovtoSaida:

    Transferencia_Le_MovtoSaida = gErr

    Select Case gErr

        Case 90473
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90474, 90475
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOSCONTACORRENTE", gErr)

        Case 90476

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175554)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Atenção: Já existe em RotinasCPR, mas foram feitas mudanças
Private Function MovCCI_Grava_Transferencia_BD(objTransferencia As ClassTransferencia) As Long
'auxiliar a MovCCI_Grava_Transferencia

Dim lErro As Long
Dim lNumMovto1 As Long
Dim lNumMovto2 As Long
Dim lSeq1 As Long
Dim lSeq2 As Long
Dim lComando As Long
Dim objContaCorrenteInterna As New ClassContasCorrentesInternas

On Error GoTo Erro_MovCCI_Grava_Transferencia_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 18242

    'Testa se já extiste um transferencia
    lErro = Transferencia_Testa_Existencia(lComando, objTransferencia)
    If lErro <> SUCESSO Then gError 90477
    
    'Gera o proximo sequencial de movimento daconta origem
    lErro = CF("CtaCorrente_Sequencial_Automatico", objTransferencia.iCodContaOrigem, lSeq1)
    If lErro <> SUCESSO Then gError 18171
    
    'Gera um Numero de movimento para o movimento de saida transferencia
    lErro = CF("MovCCI_Automatico", lNumMovto1)
    If lErro <> SUCESSO Then gError 18175
    
    objTransferencia.lNumIntDoc = lNumMovto1
    
    'Gera o proximo sequencial de movimento da conta destino
    lErro = CF("CtaCorrente_Sequencial_Automatico", objTransferencia.iCodContaDestino, lSeq2)
    If lErro <> SUCESSO Then gError 18177
      
    'Gera um Numero de movimento para o movimento de entrada transferencia
    lErro = CF("MovCCI_Automatico", lNumMovto2)
    If lErro <> SUCESSO Then gError 18180
    
    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objTransferencia.iCodContaOrigem, objContaCorrenteInterna)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 82276
    
    If lErro = 11807 Then gError 82277
    
    'Faz a insersao na tabela MovimentosContaCorrente do movimento na conta origem
    lErro = Comando_Executar(lComando, "INSERT INTO MovimentosContaCorrente (NumMovto, FilialEmpresa, CodConta, Sequencial, Tipo, TipoMeioPagto, Numero, DataMovimento, Valor, Historico, Favorecido, NumRefExterna, NumMovtoTransf) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumMovto1, objContaCorrenteInterna.iFilialEmpresa, objTransferencia.iCodContaOrigem, lSeq1, MOVCCI_SAIDA_TRANSFERENCIA, objTransferencia.iTipoMeioPagto, objTransferencia.lNumero, objTransferencia.dtData, objTransferencia.dValor, objTransferencia.sHistorico, objTransferencia.iFavorecido, objTransferencia.sNumRefExterna, lNumMovto2)
    If lErro <> AD_SQL_SUCESSO Then gError 18176
    
    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objTransferencia.iCodContaDestino, objContaCorrenteInterna)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 82278
    
    If lErro = 11807 Then gError 82279
    
    'Insere na tabela MovimentosContaCorrente do movimento na conta destino
    lErro = Comando_Executar(lComando, "INSERT INTO MovimentosContaCorrente (NumMovto, FilialEmpresa, CodConta, Sequencial, Tipo, TipoMeioPagto, Numero, DataMovimento, Valor, Historico, Favorecido, NumRefExterna, NumMovtoTransf) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumMovto2, objContaCorrenteInterna.iFilialEmpresa, objTransferencia.iCodContaDestino, lSeq2, MOVCCI_ENTRADA_TRANSFERENCIA, objTransferencia.iTipoMeioPagto, objTransferencia.lNumero, objTransferencia.dtData, objTransferencia.dValor, objTransferencia.sHistorico, objTransferencia.iFavorecido, objTransferencia.sNumRefExterna, lNumMovto1)
    If lErro <> AD_SQL_SUCESSO Then gError 18181

    Call Comando_Fechar(lComando)

    MovCCI_Grava_Transferencia_BD = SUCESSO
        
    Exit Function
    
Erro_MovCCI_Grava_Transferencia_BD:

    MovCCI_Grava_Transferencia_BD = gErr
    
    Select Case gErr
    
        Case 18171, 18175, 18177, 18180, 82276, 82278, 90477
        
        Case 18176
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MOVIMENTOSCONTACORRENTE", gErr, objTransferencia.iCodContaOrigem)
            
        Case 18181
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MOVIMENTOSCONTACORRENTE", gErr, objTransferencia.iCodContaDestino)
            
        Case 18242
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 82277
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objTransferencia.iCodContaOrigem)
    
        Case 82279
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objTransferencia.iCodContaDestino)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175555)
        
    End Select
    
    Call Comando_Fechar(lComando)
     
    Exit Function
    
End Function

'Atenção: Já existe em RotinasCPR
Function MovCCI_Grava_Transferencia(objTransferencia As ClassTransferencia, objContabil As ClassContabil) As Long
'grava uma transferencia de dinheiro feita pela tesouraria entre contas correntes da empresa

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim dtDataInicialOrigem As Date
Dim dtDataInicialDestino As Date
Dim iInativo As Integer
Dim iExigeNumero As Integer
Dim dtData As Date

On Error GoTo Erro_MovCCI_Grava_Transferencia
               
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 18155
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18156
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 18240
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 18241
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 18240

    'Verifica se a conta origem existe no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT DataSaldoInicial FROM ContasCorrentesInternas WHERE Codigo = ?", 0, dtDataInicialOrigem, objTransferencia.iCodContaOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 18157
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18158
    
    'Se nao existir ==>> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18159
    
    'Loca a conta origem
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 18160
    
    'Verifica se a conta destino Existe no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT DataSaldoInicial FROM ContasCorrentesInternas WHERE Codigo = ?", 0, dtDataInicialDestino, objTransferencia.iCodContaDestino)
    If lErro <> AD_SQL_SUCESSO Then Error 18161
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18162
    
    'Se nao existir --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18163
    
    'Loca a conta destino
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 18164
    
    'Verifica se a data do movimento e menor que a data inicial da conta origem
    If objTransferencia.dtData < dtDataInicialOrigem Then Error 18165
    
    'Verifica se a data do movimento e menor que a data inicial da conta destino
    If objTransferencia.dtData < dtDataInicialDestino Then Error 18166
    
    'Verifica se o tipo de pagamento esta cadastrado no BD
    lErro = Comando_ExecutarPos(lComando2, "SELECT Inativo, ExigeNumero FROM TipoMeioPagto WHERE Tipo = ?", 0, iInativo, iExigeNumero, objTransferencia.iTipoMeioPagto)
    If lErro <> AD_SQL_SUCESSO Then Error 18167
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18168
    
    'Se nao existir --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18169
    
    'Loca o tipo de pagamento
    lErro = Comando_LockShared(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 18170
        
    'Verifica se é um tipo de pagamento ativo
    If iInativo = TIPOMEIOPAGTO_INATIVO Then Error 18186
        
    If iExigeNumero = TIPOMEIOPAGTO_EXIGENUMERO Then
        
        If objTransferencia.lNumero = 0 Then Error 18205
        
        lErro = Comando_ExecutarPos(lComando3, "SELECT DataMovimento FROM MovimentosContaCorrente WHERE CodConta = ? AND TipoMeioPagto = ? AND Numero = ?", 0, dtDataInicialOrigem, objTransferencia.iCodContaOrigem, objTransferencia.iTipoMeioPagto, objTransferencia.lNumero)
        If lErro <> SUCESSO Then Error 18245
        
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18246
        
        If lErro = AD_SQL_SUCESSO Then Error 18247

    Else
        objTransferencia.lNumero = 0
    End If
    
    lErro = CF("MovCCI_Grava_Transferencia_BD", objTransferencia)
    If lErro <> SUCESSO Then Error 18243
    
    dtData = objTransferencia.dtData
    
    'Atualiza o movimento nas tabelas CCIMov e CCIMovDia
    lErro = CF("CCIMovDia_Grava", objTransferencia.iCodContaOrigem, dtData, -objTransferencia.dValor)
    If lErro <> SUCESSO Then Error 18182
    
    lErro = CF("CCIMov_Grava", objTransferencia.iCodContaOrigem, Year(dtData), Month(dtData), -objTransferencia.dValor)
    If lErro <> SUCESSO Then Error 18183
    
    lErro = CF("CCIMovDia_Grava", objTransferencia.iCodContaDestino, dtData, objTransferencia.dValor)
    If lErro <> SUCESSO Then Error 18184
    
    lErro = CF("CCIMov_Grava", objTransferencia.iCodContaDestino, Year(dtData), Month(dtData), objTransferencia.dValor)
    If lErro <> SUCESSO Then Error 18185
    
    'Grava os dados contábeis (contabilidade)
    lErro = objContabil.Contabil_Gravar_Registro(objTransferencia.lNumIntDoc, 0, 0, DATA_NULA)
    If lErro <> SUCESSO Then Error 20526
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 18187
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    MovCCI_Grava_Transferencia = SUCESSO
    
    Exit Function
    
Erro_MovCCI_Grava_Transferencia:

    MovCCI_Grava_Transferencia = Err
    
    Select Case Err
            
        Case 18155
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 18156, 18240, 18241
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 18157, 18158, 18245, 18246
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTASCORRENTESINTERNAS1", Err, objTransferencia.iCodContaOrigem)
        
        Case 18161, 18162
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTASCORRENTESINTERNAS1", Err, objTransferencia.iCodContaDestino)
        
        Case 18159
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objTransferencia.iCodContaOrigem)
        
        Case 18160
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTASCORRENTESINTERNAS", Err, objTransferencia.iCodContaOrigem)
              
        Case 18163
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objTransferencia.iCodContaDestino)
        
        Case 18164
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTASCORRENTESINTERNAS", Err, objTransferencia.iCodContaDestino)
        
        Case 18165
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAMOVIMENTO_MENOR", Err, objTransferencia.dtData, dtDataInicialOrigem, objTransferencia.iCodContaOrigem)
        
        Case 18166
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAMOVIMENTO_MENOR", Err, objTransferencia.dtData, dtDataInicialDestino, objTransferencia.iCodContaDestino)
        
        Case 18167, 18168
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOMEIOPAGTO1", Err, objTransferencia.iTipoMeioPagto)
            
        Case 18169
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTransferencia.iTipoMeioPagto)
        
        Case 18170
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOMEIOPAGTO", Err, objTransferencia.iTipoMeioPagto)
        
        Case 18182, 18183, 18184, 18185, 18243, 20526
        
        Case 18186
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INATIVO", Err, objTransferencia.iTipoMeioPagto)
        
        Case 18187
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case 18205
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_EXIGENUMERO", Err, objTransferencia.iTipoMeioPagto)
        
        Case 18247
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_JA_UTILIZADO", Err, objTransferencia.iCodContaOrigem, objTransferencia.iTipoMeioPagto, objTransferencia.lNumero)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175556)
            
    End Select
        
    Call Transacao_Rollback
         
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function
    
End Function

'Maristela(Antes)
Function Transferencia_Testa_Existencia(lComando As Long, objTransferencia As ClassTransferencia) As Long
'Verifica se já existe uma Transferencia, com os mesmos dados.

'Daniel - 20/09/2001
'Alterado SELECT (incluido o valor)

Dim lErro As Long
Dim dValor As Double
Dim iTipoMeioPagto As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Transferencia_Testa_Existencia
    
    'Alteracao Daniel - 20/09/2001
    'Verifica se já existe transferencia com esses dados
    
    '30/11/01 Marcelo incluido na clásula where o TipoMeioPagto, MovEntrada.Tipo , MovSaida.Tipo
    lErro = Comando_Executar(lComando, "SELECT MovEntrada.TipoMeioPagto, MovEntrada.Valor FROM MovimentosContaCorrente AS MovEntrada, MovimentosContaCorrente AS MovSaida WHERE MovEntrada.NumMovto = MovSaida.NumMovtoTransf AND MovEntrada.CodConta = ? AND MovSaida.CodConta = ? AND MovEntrada.DataMovimento = ? AND MovEntrada.Valor = ? AND MovEntrada.TipoMeioPagto = ? AND MovEntrada.Tipo = ? AND MovSaida.Tipo = ?", iTipoMeioPagto, dValor, objTransferencia.iCodContaDestino, objTransferencia.iCodContaOrigem, objTransferencia.dtData, objTransferencia.dValor, objTransferencia.iTipoMeioPagto, TIPOMOVCCI_CREDITA, TIPOMOVCCI_DEBITA)
    If lErro <> SUCESSO Then gError 90478
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90479
            
    'Existe
    If lErro = AD_SQL_SUCESSO Then
           
       'Dá a Mensagem de aviso que a transferencia será alterada
       vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_TRANSFERENCIA", objTransferencia.iCodContaOrigem, objTransferencia.iCodContaDestino, objTransferencia.iTipoMeioPagto, objTransferencia.dValor)
         
       If vbMsgRes = vbNo Then gError 90480
         
        
    End If
      
    Transferencia_Testa_Existencia = SUCESSO

    Exit Function
    
Erro_Transferencia_Testa_Existencia:

    Transferencia_Testa_Existencia = gErr
    
    Select Case gErr
    
        Case 90478, 90479
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOSCONTACORRENTE", gErr)
        
        Case 90480
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175557)
            
    End Select
    
    Exit Function

End Function
'Maristela(depois)

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


