VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl OSApontamentoOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   2
      Left            =   135
      TabIndex        =   25
      Top             =   825
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4695
         TabIndex        =   74
         Tag             =   "1"
         Top             =   2415
         Width           =   870
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descri��o do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   35
         Top             =   3675
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   39
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   3180
         Left            =   6360
         TabIndex        =   34
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   33
         Top             =   1605
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   32
         Top             =   1995
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
         TabIndex        =   31
         Top             =   1035
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padr�o"
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
         Left            =   6450
         TabIndex        =   29
         Top             =   420
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
         Left            =   6450
         TabIndex        =   28
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   27
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
         Left            =   7905
         TabIndex        =   26
         Top             =   60
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   5520
         TabIndex        =   30
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
         TabIndex        =   40
         Top             =   1290
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
         TabIndex        =   41
         Top             =   1320
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
         TabIndex        =   42
         Top             =   1260
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
         TabIndex        =   43
         Top             =   1305
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
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   45
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
         TabIndex        =   46
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
         Left            =   3795
         TabIndex        =   47
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
         TabIndex        =   48
         Top             =   1305
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
         Height          =   3180
         Left            =   6360
         TabIndex        =   49
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5609
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   3180
         Left            =   6360
         TabIndex        =   50
         Top             =   1560
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5609
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   555
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   64
         Top             =   3135
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   63
         Top             =   3135
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
         TabIndex        =   62
         Top             =   3150
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
         TabIndex        =   61
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
         TabIndex        =   60
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Hist�ricos"
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
         Caption         =   "Lan�amentos"
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
         TabIndex        =   58
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exerc�cio:"
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
         TabIndex        =   57
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   56
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   55
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Per�odo:"
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
         TabIndex        =   54
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   53
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
         TabIndex        =   52
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
         Left            =   6510
         TabIndex        =   51
         Top             =   690
         Width           =   690
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   150
      Left            =   3000
      TabIndex        =   75
      Top             =   150
      Visible         =   0   'False
      Width           =   735
      Begin VB.ComboBox Etapa 
         Height          =   315
         Left            =   4755
         TabIndex        =   77
         Top             =   0
         Visible         =   0   'False
         Width           =   3525
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
         Left            =   2505
         TabIndex        =   76
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSMask.MaskEdBox Projeto 
         Height          =   300
         Left            =   735
         TabIndex        =   78
         Top             =   15
         Visible         =   0   'False
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
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
         Left            =   0
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   80
         Top             =   60
         Visible         =   0   'False
         Width           =   675
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
         Left            =   4125
         TabIndex        =   79
         Top             =   45
         Visible         =   0   'False
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5070
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9270
      Begin VB.Frame FrameItens 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   3285
         Index           =   1
         Left            =   75
         TabIndex        =   83
         Top             =   1710
         Width           =   9105
         Begin VB.ComboBox Benef 
            Height          =   315
            ItemData        =   "OSApontamento.ctx":0000
            Left            =   60
            List            =   "OSApontamento.ctx":0007
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   2175
            Width           =   1260
         End
         Begin VB.CommandButton BotaoServicos 
            Caption         =   "Servi�os"
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
            Left            =   3456
            TabIndex        =   114
            Top             =   2910
            Width           =   915
         End
         Begin VB.ComboBox FilialOP 
            Height          =   315
            Left            =   2475
            TabIndex        =   103
            Top             =   975
            Width           =   2160
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   1005
            Width           =   645
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   3570
            MaxLength       =   50
            TabIndex        =   101
            Top             =   1455
            Width           =   2640
         End
         Begin VB.CheckBox Estorno 
            Height          =   210
            Left            =   4890
            TabIndex        =   100
            Top             =   1830
            Width           =   870
         End
         Begin VB.TextBox OPCodigo 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   7110
            MaxLength       =   9
            TabIndex        =   97
            Top             =   1095
            Width           =   1260
         End
         Begin VB.CommandButton BotaoPlanoConta 
            Caption         =   "Contas"
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
            Left            =   6522
            TabIndex        =   92
            Top             =   2910
            Width           =   840
         End
         Begin VB.CommandButton BotaoCcls 
            Caption         =   "Ccl"
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
            Left            =   5545
            TabIndex        =   91
            Top             =   2910
            Width           =   810
         End
         Begin VB.CommandButton BotaoOP 
            Caption         =   "OS"
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
            Left            =   2494
            TabIndex        =   90
            Top             =   2910
            Width           =   795
         End
         Begin VB.CommandButton BotaoLote 
            Caption         =   "Lotes"
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
            Left            =   4538
            TabIndex        =   89
            Top             =   2910
            Width           =   840
         End
         Begin VB.CommandButton BotaoSerie 
            Caption         =   "N�m.S�ries ..."
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
            Left            =   7530
            TabIndex        =   88
            Top             =   2910
            Width           =   1545
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
            Height          =   315
            Left            =   1292
            TabIndex        =   87
            Top             =   2910
            Width           =   1035
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
            Height          =   315
            Left            =   15
            TabIndex        =   86
            Top             =   2910
            Width           =   1110
         End
         Begin MSMask.MaskEdBox Lote 
            Height          =   255
            Left            =   5910
            TabIndex        =   98
            Top             =   1050
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   240
            Left            =   6405
            TabIndex        =   99
            Top             =   1470
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaContabilProducao 
            Height          =   270
            Left            =   1230
            TabIndex        =   104
            Top             =   1440
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox ContaContabilEst 
            Height          =   240
            Left            =   90
            TabIndex        =   105
            Top             =   1815
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox ProdutoOP 
            Height          =   240
            Left            =   3600
            TabIndex        =   106
            Top             =   1815
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   240
            Left            =   90
            TabIndex        =   107
            Top             =   1440
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   423
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
            Height          =   240
            Left            =   0
            TabIndex        =   108
            Top             =   1020
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   240
            Left            =   2580
            TabIndex        =   109
            Top             =   1815
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   423
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
         Begin MSFlexGridLib.MSFlexGrid GridMovs 
            Height          =   885
            Left            =   10
            TabIndex        =   85
            Top             =   45
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   1561
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
            Height          =   300
            Left            =   2235
            TabIndex        =   127
            Top             =   2565
            Width           =   1740
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade Dispon�vel:"
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
            TabIndex        =   126
            Top             =   2610
            Width           =   2025
         End
      End
      Begin VB.Frame FrameItens 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   3285
         Index           =   3
         Left            =   75
         TabIndex        =   82
         Top             =   1710
         Visible         =   0   'False
         Width           =   9105
         Begin VB.TextBox Maquina 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1095
            TabIndex        =   122
            Top             =   1215
            Width           =   2895
         End
         Begin VB.CommandButton BotaoServicos 
            Caption         =   "Servi�os"
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
            Index           =   2
            Left            =   2250
            TabIndex        =   118
            Top             =   2910
            Width           =   915
         End
         Begin VB.CommandButton BotaoOP 
            Caption         =   "OS"
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
            Index           =   2
            Left            =   1425
            TabIndex        =   117
            Top             =   2910
            Width           =   795
         End
         Begin VB.CommandButton BotaoMaquinas 
            Caption         =   "M�quinas"
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
            Left            =   0
            TabIndex        =   94
            Top             =   2910
            Width           =   1380
         End
         Begin MSMask.MaskEdBox MaqQtd 
            Height          =   315
            Left            =   3375
            TabIndex        =   121
            Top             =   1230
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaqOS 
            Height          =   315
            Left            =   4905
            TabIndex        =   123
            Top             =   1260
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaqServico 
            Height          =   315
            Left            =   6270
            TabIndex        =   124
            Top             =   1260
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaqHoras 
            Height          =   315
            Left            =   3090
            TabIndex        =   125
            Top             =   1770
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaq 
            Height          =   885
            Left            =   10
            TabIndex        =   96
            Top             =   45
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   1561
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame FrameItens 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   3285
         Index           =   2
         Left            =   75
         TabIndex        =   84
         Top             =   1710
         Visible         =   0   'False
         Width           =   9105
         Begin VB.TextBox MOTipo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1095
            TabIndex        =   128
            Top             =   1590
            Width           =   1785
         End
         Begin VB.CommandButton BotaoServicos 
            Caption         =   "Servi�os"
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
            Index           =   1
            Left            =   2265
            TabIndex        =   116
            Top             =   2910
            Width           =   915
         End
         Begin VB.CommandButton BotaoOP 
            Caption         =   "OS"
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
            Index           =   1
            Left            =   1425
            TabIndex        =   115
            Top             =   2910
            Width           =   795
         End
         Begin VB.TextBox MONomeRed 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1095
            TabIndex        =   112
            Top             =   1125
            Width           =   1785
         End
         Begin VB.TextBox MOCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            TabIndex        =   111
            Top             =   1110
            Width           =   1395
         End
         Begin VB.CommandButton BotaoMO 
            Caption         =   "M�o-de-obra"
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
            Left            =   15
            TabIndex        =   93
            Top             =   2910
            Width           =   1380
         End
         Begin MSMask.MaskEdBox MOHoras 
            Height          =   315
            Left            =   3375
            TabIndex        =   110
            Top             =   1140
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MOOS 
            Height          =   315
            Left            =   4905
            TabIndex        =   113
            Top             =   1170
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MOServico 
            Height          =   315
            Left            =   6270
            TabIndex        =   120
            Top             =   1170
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridMO 
            Height          =   885
            Left            =   10
            TabIndex        =   95
            Top             =   45
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   1561
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3645
         Left            =   15
         TabIndex        =   81
         Top             =   1380
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   6429
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Pe�as"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "M�o-de-obra"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "M�quinas"
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
      Begin VB.Frame Frame2 
         Caption         =   "Gera��o Autom�tica"
         Height          =   555
         Left            =   15
         TabIndex        =   3
         Top             =   765
         Width           =   9240
         Begin VB.CommandButton BotaoGeraReq 
            Caption         =   "Gerar"
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
            Left            =   8220
            TabIndex        =   5
            Top             =   150
            Width           =   960
         End
         Begin VB.TextBox OP 
            Height          =   285
            Left            =   1335
            MaxLength       =   9
            TabIndex        =   4
            Top             =   195
            Width           =   1230
         End
         Begin MSMask.MaskEdBox ProdutoOPGera 
            Height          =   285
            Left            =   3495
            TabIndex        =   6
            Top             =   195
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeOP 
            Height          =   285
            Left            =   6930
            TabIndex        =   7
            Top             =   195
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   4815
            TabIndex        =   12
            Top             =   255
            Width           =   480
         End
         Begin VB.Label LblUM 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5310
            TabIndex        =   11
            Top             =   195
            Width           =   675
         End
         Begin VB.Label OPLabel 
            AutoSize        =   -1  'True
            Caption         =   "OS:"
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
            Left            =   960
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   10
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Quant.:"
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
            Left            =   6225
            TabIndex        =   9
            Top             =   240
            Width           =   645
         End
         Begin VB.Label ProdutoOPLabel 
            AutoSize        =   -1  'True
            Caption         =   "Servi�o:"
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
            Left            =   2745
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   8
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.TextBox OPCodigoPadrao 
         Height          =   285
         Left            =   8235
         MaxLength       =   9
         TabIndex        =   2
         Top             =   420
         Width           =   1005
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2175
         Picture         =   "OSApontamento.ctx":0017
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numera��o Autom�tica"
         Top             =   15
         Width           =   300
      End
      Begin MSMask.MaskEdBox AlmoxPadrao 
         Height          =   285
         Left            =   1365
         TabIndex        =   13
         Top             =   420
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   6420
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   5340
         TabIndex        =   15
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclPadrao 
         Height          =   285
         Left            =   5340
         TabIndex        =   16
         Top             =   420
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1365
         TabIndex        =   17
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Hora 
         Height          =   300
         Left            =   8235
         TabIndex        =   18
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label OPPadraoLabel 
         AutoSize        =   -1  'True
         Caption         =   "OS Padr�o:"
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
         Left            =   7245
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   450
         Width           =   975
      End
      Begin VB.Label CodigoLabel 
         AutoSize        =   -1  'True
         Caption         =   "C�digo:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   45
         Width           =   660
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
         Height          =   180
         Left            =   4800
         TabIndex        =   22
         Top             =   45
         Width           =   480
      End
      Begin VB.Label CclPadraoLabel 
         AutoSize        =   -1  'True
         Caption         =   "C.Custo Padr�o:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label AlmoxPadraoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Almox. Padr�o:"
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
         Left            =   30
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   450
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Left            =   7755
         TabIndex        =   19
         Top             =   45
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7290
      ScaleHeight     =   495
      ScaleWidth      =   2100
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   60
      Width           =   2160
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "OSApontamento.ctx":0101
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "OSApontamento.ctx":028B
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "OSApontamento.ctx":0409
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OSApontamento.ctx":093B
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5580
      Left            =   45
      TabIndex        =   73
      Top             =   375
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   9843
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabiliza��o"
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
Attribute VB_Name = "OSApontamentoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTOSApontamento
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTOSApontamento
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoPlanoConta_Click()
     Call objCT.BotaoPlanoConta_Click
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub ContaContabilEst_Change()
     Call objCT.ContaContabilEst_Change
End Sub

Private Sub ContaContabilEst_GotFocus()
     Call objCT.ContaContabilEst_GotFocus
End Sub

Private Sub ContaContabilEst_KeyPress(KeyAscii As Integer)
     Call objCT.ContaContabilEst_KeyPress(KeyAscii)
End Sub

Private Sub ContaContabilEst_Validate(Cancel As Boolean)
     Call objCT.ContaContabilEst_Validate(Cancel)
End Sub

Private Sub ContaContabilProducao_Change()
     Call objCT.ContaContabilProducao_Change
End Sub

Private Sub ContaContabilProducao_GotFocus()
     Call objCT.ContaContabilProducao_GotFocus
End Sub

Private Sub ContaContabilProducao_KeyPress(KeyAscii As Integer)
     Call objCT.ContaContabilProducao_KeyPress(KeyAscii)
End Sub

Private Sub ContaContabilProducao_Validate(Cancel As Boolean)
     Call objCT.ContaContabilProducao_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub CodigoLabel_Click()
     Call objCT.CodigoLabel_Click
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub OPPadraoLabel_Click()
     Call objCT.OPPadraoLabel_Click
End Sub

Private Sub CclPadraoLabel_Click()
     Call objCT.CclPadraoLabel_Click
End Sub

Private Sub AlmoxPadraoLabel_Click()
     Call objCT.AlmoxPadraoLabel_Click
End Sub

Private Sub OPLabel_Click()
     Call objCT.OPLabel_Click
End Sub

Private Sub ProdutoOPLabel_Click()
     Call objCT.ProdutoOPLabel_Click
End Sub

Private Sub BotaoOP_Click(Index As Integer)
     Call objCT.BotaoOP_Click(Index)
End Sub

Private Sub BotaoCcls_Click()
     Call objCT.BotaoCcls_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub BotaoEstoque_Click()
     Call objCT.BotaoEstoque_Click
End Sub

Function Trata_Parametros(Optional obj1 As Object) As Long
     Trata_Parametros = objCT.Trata_Parametros(obj1)
End Function

Private Sub OP_Validate(bMantemFoco As Boolean)
     Call objCT.OP_Validate(bMantemFoco)
End Sub

Private Sub ProdutoOPGera_Validate(bMantemFoco As Boolean)
     Call objCT.ProdutoOPGera_Validate(bMantemFoco)
End Sub

Private Sub QuantidadeOP_Validate(Cancel As Boolean)
     Call objCT.QuantidadeOP_Validate(Cancel)
End Sub

Private Sub BotaoGeraReq_Click()
     Call objCT.BotaoGeraReq_Click
End Sub

Private Sub OPCodigoPadrao_Validate(Cancel As Boolean)
     Call objCT.OPCodigoPadrao_Validate(Cancel)
End Sub

Private Sub CclPadrao_Validate(Cancel As Boolean)
     Call objCT.CclPadrao_Validate(Cancel)
End Sub

Private Sub AlmoxPadrao_Validate(Cancel As Boolean)
     Call objCT.AlmoxPadrao_Validate(Cancel)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
     Call objCT.Data_Validate(Cancel)
End Sub

Private Sub UpDownData_DownClick()
     Call objCT.UpDownData_DownClick
End Sub

Private Sub UpDownData_UpClick()
     Call objCT.UpDownData_UpClick
End Sub

Private Sub Lote_Change()
     Call objCT.Lote_Change
End Sub

Private Sub Lote_GotFocus()
     Call objCT.Lote_GotFocus
End Sub

Private Sub Lote_KeyPress(KeyAscii As Integer)
     Call objCT.Lote_KeyPress(KeyAscii)
End Sub

Private Sub Lote_Validate(Cancel As Boolean)
     Call objCT.Lote_Validate(Cancel)
End Sub

Private Sub GridMovs_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMovs_KeyDown(KeyCode, Shift)
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        Set objCT.objUserControl = Nothing
        Set objCT = Nothing
    End If
End Sub

Private Sub Ccl_Change()
     Call objCT.Ccl_Change
End Sub

Private Sub CclPadrao_Change()
     Call objCT.CclPadrao_Change
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub Data_Change()
     Call objCT.Data_Change
End Sub

Private Sub DescricaoItem_Change()
     Call objCT.DescricaoItem_Change
End Sub

Private Sub OPCodigo_Change()
     Call objCT.OPCodigo_Change
End Sub

Private Sub OPCodigoPadrao_Change()
     Call objCT.OPCodigoPadrao_Change
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub UnidadeMed_Change()
     Call objCT.UnidadeMed_Change
End Sub

Private Sub Almoxarifado_Change()
     Call objCT.Almoxarifado_Change
End Sub

Private Sub AlmoxPadrao_Change()
     Call objCT.AlmoxPadrao_Change
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub Estorno_Click()
     Call objCT.Estorno_Click
End Sub

Private Sub OP_Change()
     Call objCT.OP_Change
End Sub

Private Sub ProdutoOP_Change()
     Call objCT.ProdutoOP_Change
End Sub

Private Sub ProdutoOPGera_Change()
     Call objCT.ProdutoOPGera_Change
End Sub

Private Sub QuantidadeOP_Change()
     Call objCT.QuantidadeOP_Change
End Sub

Private Sub GridMovs_Click()
     Call objCT.GridMovs_Click
End Sub

Private Sub GridMovs_EnterCell()
     Call objCT.GridMovs_EnterCell
End Sub

Private Sub GridMovs_GotFocus()
     Call objCT.GridMovs_GotFocus
End Sub

Private Sub GridMovs_KeyPress(KeyAscii As Integer)
     Call objCT.GridMovs_KeyPress(KeyAscii)
End Sub

Private Sub GridMovs_LeaveCell()
     Call objCT.GridMovs_LeaveCell
End Sub

Private Sub GridMovs_Validate(Cancel As Boolean)
     Call objCT.GridMovs_Validate(Cancel)
End Sub

Private Sub GridMovs_Scroll()
     Call objCT.GridMovs_Scroll
End Sub

Private Sub GridMovs_RowColChange()
     Call objCT.GridMovs_RowColChange
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

Private Sub Ccl_GotFocus()
     Call objCT.Ccl_GotFocus
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)
     Call objCT.Ccl_KeyPress(KeyAscii)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub Estorno_GotFocus()
     Call objCT.Estorno_GotFocus
End Sub

Private Sub Estorno_KeyPress(KeyAscii As Integer)
     Call objCT.Estorno_KeyPress(KeyAscii)
End Sub

Private Sub Estorno_Validate(Cancel As Boolean)
     Call objCT.Estorno_Validate(Cancel)
End Sub

Private Sub OPCodigo_GotFocus()
     Call objCT.OPCodigo_GotFocus
End Sub

Private Sub OPCodigo_KeyPress(KeyAscii As Integer)
     Call objCT.OPCodigo_KeyPress(KeyAscii)
End Sub

Private Sub OPCodigo_Validate(Cancel As Boolean)
     Call objCT.OPCodigo_Validate(Cancel)
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

Private Sub Quantidade_GotFocus()
     Call objCT.Quantidade_GotFocus
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
     Call objCT.Quantidade_KeyPress(KeyAscii)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub UnidadeMed_GotFocus()
     Call objCT.UnidadeMed_GotFocus
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMed_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)
     Call objCT.UnidadeMed_Validate(Cancel)
End Sub

Private Sub ProdutoOP_GotFocus()
     Call objCT.ProdutoOP_GotFocus
End Sub

Private Sub ProdutoOP_KeyPress(KeyAscii As Integer)
     Call objCT.ProdutoOP_KeyPress(KeyAscii)
End Sub

Private Sub ProdutoOP_Validate(Cancel As Boolean)
     Call objCT.ProdutoOP_Validate(Cancel)
End Sub

Private Sub CTBBotaoModeloPadrao_Click()
     Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Private Sub CTBModelo_Click()
     Call objCT.CTBModelo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub CTBGridContabil_EnterCell()
     Call objCT.CTBGridContabil_EnterCell
End Sub

Private Sub CTBGridContabil_GotFocus()
     Call objCT.CTBGridContabil_GotFocus
End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Private Sub CTBGridContabil_LeaveCell()
     Call objCT.CTBGridContabil_LeaveCell
End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)
     Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Private Sub CTBGridContabil_RowColChange()
     Call objCT.CTBGridContabil_RowColChange
End Sub

Private Sub CTBGridContabil_Scroll()
     Call objCT.CTBGridContabil_Scroll
End Sub

Private Sub CTBConta_Change()
     Call objCT.CTBConta_Change
End Sub

Private Sub CTBConta_GotFocus()
     Call objCT.CTBConta_GotFocus
End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)
     Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)
     Call objCT.CTBConta_Validate(Cancel)
End Sub

Private Sub CTBCcl_Change()
     Call objCT.CTBCcl_Change
End Sub

Private Sub CTBCcl_GotFocus()
     Call objCT.CTBCcl_GotFocus
End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)
     Call objCT.CTBCcl_Validate(Cancel)
End Sub

Private Sub CTBCredito_Change()
     Call objCT.CTBCredito_Change
End Sub

Private Sub CTBCredito_GotFocus()
     Call objCT.CTBCredito_GotFocus
End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)
     Call objCT.CTBCredito_Validate(Cancel)
End Sub

Private Sub CTBDebito_Change()
     Call objCT.CTBDebito_Change
End Sub

Private Sub CTBDebito_GotFocus()
     Call objCT.CTBDebito_GotFocus
End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)
     Call objCT.CTBDebito_Validate(Cancel)
End Sub

Private Sub CTBSeqContraPartida_Change()
     Call objCT.CTBSeqContraPartida_Change
End Sub

Private Sub CTBSeqContraPartida_GotFocus()
     Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
End Sub

Private Sub CTBHistorico_Change()
     Call objCT.CTBHistorico_Change
End Sub

Private Sub CTBHistorico_GotFocus()
     Call objCT.CTBHistorico_GotFocus
End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
     Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)
     Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Private Sub CTBLancAutomatico_Click()
     Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
     Call objCT.CTBAglutina_Click
End Sub

Private Sub CTBAglutina_GotFocus()
     Call objCT.CTBAglutina_GotFocus
End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
     Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)
     Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Private Sub CTBListHistoricos_DblClick()
     Call objCT.CTBListHistoricos_DblClick
End Sub

Private Sub CTBBotaoLimparGrid_Click()
     Call objCT.CTBBotaoLimparGrid_Click
End Sub

Private Sub CTBLote_Change()
     Call objCT.CTBLote_Change
End Sub

Private Sub CTBLote_GotFocus()
     Call objCT.CTBLote_GotFocus
End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)
     Call objCT.CTBLote_Validate(Cancel)
End Sub

Private Sub CTBDataContabil_Change()
     Call objCT.CTBDataContabil_Change
End Sub

Private Sub CTBDataContabil_GotFocus()
     Call objCT.CTBDataContabil_GotFocus
End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)
     Call objCT.CTBDataContabil_Validate(Cancel)
End Sub

Private Sub CTBDocumento_Change()
     Call objCT.CTBDocumento_Change
End Sub

Private Sub CTBDocumento_GotFocus()
     Call objCT.CTBDocumento_GotFocus
End Sub

Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub CTBLabelDoc_Click()
     Call objCT.CTBLabelDoc_Click
End Sub

Private Sub CTBLabelLote_Click()
     Call objCT.CTBLabelLote_Click
End Sub

Private Sub ProdutoOPLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoOPLabel, Source, X, Y)
End Sub
Private Sub ProdutoOPLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoOPLabel, Button, Shift, X, Y)
End Sub
Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub
Private Sub OPLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OPLabel, Source, X, Y)
End Sub
Private Sub OPLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OPLabel, Button, Shift, X, Y)
End Sub
Private Sub LblUM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUM, Source, X, Y)
End Sub
Private Sub LblUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUM, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(iIndex As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(iIndex), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(iIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(iIndex), Button, Shift, X, Y)
End Sub
Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub
Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub
Private Sub OPPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OPPadraoLabel, Source, X, Y)
End Sub
Private Sub OPPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OPPadraoLabel, Button, Shift, X, Y)
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
     Call objCT.Opcao_BeforeClick(Cancel)
End Sub

Private Sub BotaoLote_Click()
     Call objCT.BotaoLote_Click
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

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
     Call objCT.TabStrip1_BeforeClick(Cancel)
End Sub

Private Sub BotaoMaquinas_Click()
     Call objCT.BotaoMaquinas_Click
End Sub

Private Sub BotaoMO_Click()
     Call objCT.BotaoMO_Click
End Sub

Private Sub GridMaq_Click()
     Call objCT.GridMaq_Click
End Sub

Private Sub GridMaq_GotFocus()
     Call objCT.GridMaq_GotFocus
End Sub

Private Sub GridMaq_EnterCell()
     Call objCT.GridMaq_EnterCell
End Sub

Private Sub GridMaq_LeaveCell()
     Call objCT.GridMaq_LeaveCell
End Sub

Private Sub GridMaq_KeyPress(KeyAscii As Integer)
     Call objCT.GridMaq_KeyPress(KeyAscii)
End Sub

Private Sub GridMaq_RowColChange()
     Call objCT.GridMaq_RowColChange
End Sub

Private Sub GridMaq_Scroll()
     Call objCT.GridMaq_Scroll
End Sub

Private Sub GridMaq_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMaq_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridMaq_LostFocus()
     Call objCT.GridMaq_LostFocus
End Sub

Private Sub GridMO_Click()
     Call objCT.GridMO_Click
End Sub

Private Sub GridMO_GotFocus()
     Call objCT.GridMO_GotFocus
End Sub

Private Sub GridMO_EnterCell()
     Call objCT.GridMO_EnterCell
End Sub

Private Sub GridMO_LeaveCell()
     Call objCT.GridMO_LeaveCell
End Sub

Private Sub GridMO_KeyPress(KeyAscii As Integer)
     Call objCT.GridMO_KeyPress(KeyAscii)
End Sub

Private Sub GridMO_RowColChange()
     Call objCT.GridMO_RowColChange
End Sub

Private Sub GridMO_Scroll()
     Call objCT.GridMO_Scroll
End Sub

Private Sub GridMO_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMO_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridMO_LostFocus()
     Call objCT.GridMO_LostFocus
End Sub

Private Sub MOCodigo_Change()
     Call objCT.MOCodigo_Change
End Sub

Private Sub MOCodigo_GotFocus()
     Call objCT.MOCodigo_GotFocus
End Sub

Private Sub MOCodigo_KeyPress(KeyAscii As Integer)
     Call objCT.MOCodigo_KeyPress(KeyAscii)
End Sub

Private Sub MOCodigo_Validate(Cancel As Boolean)
     Call objCT.MOCodigo_Validate(Cancel)
End Sub

Private Sub MOHoras_Change()
     Call objCT.MOHoras_Change
End Sub

Private Sub MOHoras_GotFocus()
     Call objCT.MOHoras_GotFocus
End Sub

Private Sub MOHoras_KeyPress(KeyAscii As Integer)
     Call objCT.MOHoras_KeyPress(KeyAscii)
End Sub

Private Sub MOHoras_Validate(Cancel As Boolean)
     Call objCT.MOHoras_Validate(Cancel)
End Sub

Private Sub MOOS_Change()
     Call objCT.MOOS_Change
End Sub

Private Sub MOOS_GotFocus()
     Call objCT.MOOS_GotFocus
End Sub

Private Sub MOOS_KeyPress(KeyAscii As Integer)
     Call objCT.MOOS_KeyPress(KeyAscii)
End Sub

Private Sub MOOS_Validate(Cancel As Boolean)
     Call objCT.MOOS_Validate(Cancel)
End Sub

Private Sub MOServico_Change()
     Call objCT.MOServico_Change
End Sub

Private Sub MOServico_GotFocus()
     Call objCT.MOServico_GotFocus
End Sub

Private Sub MOServico_KeyPress(KeyAscii As Integer)
     Call objCT.MOServico_KeyPress(KeyAscii)
End Sub

Private Sub MOServico_Validate(Cancel As Boolean)
     Call objCT.MOServico_Validate(Cancel)
End Sub

Private Sub Maquina_Change()
     Call objCT.Maquina_Change
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

Private Sub MaqQtd_Change()
     Call objCT.MaqQtd_Change
End Sub

Private Sub MaqQtd_GotFocus()
     Call objCT.MaqQtd_GotFocus
End Sub

Private Sub MaqQtd_KeyPress(KeyAscii As Integer)
     Call objCT.MaqQtd_KeyPress(KeyAscii)
End Sub

Private Sub MaqQtd_Validate(Cancel As Boolean)
     Call objCT.MaqQtd_Validate(Cancel)
End Sub

Private Sub MaqHoras_Change()
     Call objCT.MaqHoras_Change
End Sub

Private Sub MaqHoras_GotFocus()
     Call objCT.MaqHoras_GotFocus
End Sub

Private Sub MaqHoras_KeyPress(KeyAscii As Integer)
     Call objCT.MaqHoras_KeyPress(KeyAscii)
End Sub

Private Sub MaqHoras_Validate(Cancel As Boolean)
     Call objCT.MaqHoras_Validate(Cancel)
End Sub

Private Sub MaqOS_Change()
     Call objCT.MaqOS_Change
End Sub

Private Sub MaqOS_GotFocus()
     Call objCT.MaqOS_GotFocus
End Sub

Private Sub MaqOS_KeyPress(KeyAscii As Integer)
     Call objCT.MaqOS_KeyPress(KeyAscii)
End Sub

Private Sub MaqOS_Validate(Cancel As Boolean)
     Call objCT.MaqOS_Validate(Cancel)
End Sub

Private Sub MaqServico_Change()
     Call objCT.MaqServico_Change
End Sub

Private Sub MaqServico_GotFocus()
     Call objCT.MaqServico_GotFocus
End Sub

Private Sub MaqServico_KeyPress(KeyAscii As Integer)
     Call objCT.MaqServico_KeyPress(KeyAscii)
End Sub

Private Sub MaqServico_Validate(Cancel As Boolean)
     Call objCT.MaqServico_Validate(Cancel)
End Sub

Private Sub BotaoServicos_Click(Index As Integer)
     Call objCT.BotaoServicos_Click(Index)
End Sub
