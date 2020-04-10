VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LoteEstAtualiza 
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   ScaleHeight     =   5160
   ScaleWidth      =   8790
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   8
      Left            =   8370
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   9240
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   29
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   28
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2790
         Left            =   6330
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
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
         Height          =   510
         Left            =   6330
         TabIndex        =   26
         Top             =   630
         Width           =   1245
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
         Height          =   375
         Left            =   6330
         TabIndex        =   25
         Top             =   120
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   7740
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   810
         Width           =   1260
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   19
         Top             =   3330
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   315
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   21
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   20
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
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
         Height          =   375
         Left            =   7710
         TabIndex        =   18
         Top             =   135
         Width           =   1245
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
         TabIndex        =   17
         Top             =   930
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4800
         TabIndex        =   30
         Top             =   1560
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
         TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   33
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
         TabIndex        =   34
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
      Begin MSComCtl2.UpDown CTBUpDown3 
         Height          =   300
         Left            =   1635
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil3 
         Height          =   300
         Left            =   570
         TabIndex        =   36
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
      Begin MSMask.MaskEdBox CTBLote3 
         Height          =   300
         Left            =   5580
         TabIndex        =   37
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
         Left            =   1845
         TabIndex        =   38
         Top             =   3030
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2790
         Left            =   6330
         TabIndex        =   39
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4921
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2790
         Left            =   6330
         TabIndex        =   40
         Top             =   1560
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4921
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   41
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
         Left            =   3600
         TabIndex        =   58
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4305
         TabIndex        =   57
         Top             =   3075
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
         TabIndex        =   56
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   55
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   1305
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
         Left            =   6360
         TabIndex        =   49
         Top             =   1290
         Visible         =   0   'False
         Width           =   2490
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
         Height          =   195
         Left            =   7755
         TabIndex        =   48
         Top             =   585
         Width           =   690
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
         TabIndex        =   47
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   46
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   45
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
         TabIndex        =   44
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
         Height          =   195
         Left            =   750
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   43
         Top             =   3075
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote3 
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
         TabIndex        =   42
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.CheckBox ExibirLotesAtualizando 
      Caption         =   "Exibir os lotes que estão sendo atualizados"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   930
      Width           =   3495
   End
   Begin VB.TextBox Status 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   7530
      TabIndex        =   7
      Text            =   "Status"
      Top             =   4005
      Width           =   1065
   End
   Begin VB.CommandButton BotaoFechar 
      Caption         =   "Fechar"
      Height          =   585
      Left            =   6810
      Picture         =   "LoteEstAtualiza.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4410
      Width           =   1140
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todos"
      Height          =   570
      Left            =   2925
      Picture         =   "LoteEstAtualiza.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4425
      Width           =   1425
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todos"
      Height          =   570
      Left            =   840
      Picture         =   "LoteEstAtualiza.ctx":1360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4425
      Width           =   1425
   End
   Begin VB.CommandButton BotaoAtualizar 
      Caption         =   "Atualizar"
      Height          =   585
      Left            =   5010
      Picture         =   "LoteEstAtualiza.ctx":237A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4425
      Width           =   1140
   End
   Begin VB.TextBox Lote 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   2115
      TabIndex        =   4
      Text            =   "Lote"
      Top             =   3990
      Width           =   720
   End
   Begin VB.CheckBox Atualiza 
      Height          =   195
      Left            =   1050
      TabIndex        =   3
      Top             =   4005
      Width           =   870
   End
   Begin VB.TextBox Descricao 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   2880
      TabIndex        =   5
      Text            =   "Descricao"
      Top             =   3990
      Width           =   3090
   End
   Begin VB.TextBox NumLancAtual 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   6000
      TabIndex        =   6
      Text            =   "NumLancAtual"
      Top             =   3990
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GridLotesPendentes 
      Height          =   2805
      Left            =   150
      TabIndex        =   8
      Top             =   1245
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   11
      Cols            =   7
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   1
   End
   Begin MSComCtl2.UpDown CTBUpDown 
      Height          =   300
      Left            =   3390
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   375
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox CTBDataContabil 
      Height          =   300
      Left            =   2310
      TabIndex        =   12
      Top             =   375
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
      Left            =   5475
      TabIndex        =   13
      Top             =   375
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label LabelDataContabil 
      AutoSize        =   -1  'True
      Caption         =   "Data de Contabilização:"
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
      Left            =   225
      TabIndex        =   15
      Top             =   405
      Width           =   2055
   End
   Begin VB.Label CTBLabelLote 
      AutoSize        =   -1  'True
      Caption         =   "Lote Contábil:"
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
      Left            =   4200
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   14
      Top             =   420
      Width           =   1200
   End
End
Attribute VB_Name = "LoteEstAtualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'OBSERVACAO:
'O frame de contabilidade foi incluido nesta tela apenas p/evitar duplicidade de codigo p/tratamento de campos de lote e data contabil. Por isso ele fica invisivel.

'Property Variables:
Dim m_Caption As String
Event Unload()

'Colunas do Grid da Tela LoteEstAtualiza
'coluna referente a checkbox atualiza no grid da tela LoteEstAtualiza
Const COL_ATUALIZA = 1
'coluna referente ao Lote no grid da tela LoteEstAtualiza
Const COL_LOTE = 2
'coluna referente a descrição no grid da tela LoteEstAtualiza
Const COL_DESCRICAO = 3
'coluna referente ao número de lançamentos no grid da tela LoteEstAtualiza
Const COL_NUMLANCATUAL = 4
'coluna referente ao Status no grid da tela LoteEstAtualiza
Const COL_STATUS = 5

Public iAlterado As Integer
Dim objGrid As AdmGrid

'Associados a contabilidade
Public objContabil As New ClassContabil
Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Public objGrid1 As AdmGrid

Public Function Trata_Parametros() As Long

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Private Sub BotaoDesmarcarTodos_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoDesmarcarTodos_Click

    'percorre todas as linhas do grid
    For iIndice = 1 To objGrid.iLinhasExistentes
        'marca cada checkbox Atualiza do grid
        GridLotesPendentes.TextMatrix(iIndice, COL_ATUALIZA) = "0"
    Next

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then Error 41215

    Exit Sub

Erro_BotaoDesmarcarTodos_Click:

    Select Case Err

        Case 41215

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162458)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoMarcarTodos_Click

    'percorre todas as linhas do grid
    For iIndice = 1 To objGrid.iLinhasExistentes
        'marca cada checkbox Atualiza do grid
            GridLotesPendentes.TextMatrix(iIndice, COL_ATUALIZA) = "1"
    Next

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then Error 41216

    Exit Sub

Erro_BotaoMarcarTodos_Click:

    Select Case Err

        Case 41216

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162459)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAtualizar_Click()

Dim lErro As Long
Dim colInvLote As New Collection
Dim iIDAtualizacao As Integer
Dim sNomeArqParam As String, dtDataContabil As Date, iLoteContabil As Integer

On Error GoTo Erro_BotaoAtualizar_Click

    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        'Verifica se a data contabil está preenchida
        If Len(CTBDataContabil.ClipText) = 0 Then Error 59414
 
        dtDataContabil = CDate(CTBDataContabil.Text)
        
        'se o lote estiver preenchido ==> não pode ser com o valor zero
        If Len(CTBLote.ClipText) > 0 And giTipoVersao = VERSAO_FULL Then
            iLoteContabil = CInt(CTBLote.ClipText)
        Else
            'se não estiver preenchido o lote ==> atualizacao imediata e o valor do lote será zero internamente
            iLoteContabil = 0
        End If
    
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'le os dados do grid que estao marcados para serem atualizados e coloca na colecao colInvLote
    lErro = GridLotesPendentes_Le(colInvLote)
    If lErro <> SUCESSO Then Error 41052

    'Atualiza o campo IdAtualizacao das tabelas Configuracao e LotePendente
    lErro = CF("InvLotePendente_Atualiza",colInvLote, iIDAtualizacao)
    If lErro <> SUCESSO Then Error 41053

    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 41054

    lErro = CF("Rotina_Atualizacao_InvLote",sNomeArqParam, iIDAtualizacao, dtDataContabil, iLoteContabil)
    If lErro <> SUCESSO Then Error 41055

    'limpa o grid
    Call Grid_Limpa(objGrid)

    Set colInvLote = New Collection

    If ExibirLotesAtualizando.Value = 1 Then

        'le todos os lotes com status = desatualizado na tabela LotePendente e coloca na colecao colInvLote
        lErro = CF("InvLotePendente_Le_Desatualizados",giFilialEmpresa, colInvLote, LOTES_PENDENTES)
        If lErro <> SUCESSO Then Error 41059

    Else

        'leitura dos Lotes desatualizados e IdAtualizacao zerados
        lErro = CF("InvLotePendente_Le_Desatualizados",giFilialEmpresa, colInvLote, LOTE_DESATUALIZADO)
        If lErro <> SUCESSO Then Error 41196

    End If

    'preenche o grid com os dados da colecao
    lErro = Grid_Preenche(colInvLote)
    If lErro <> SUCESSO Then Error 41197

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoAtualizar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 41052, 41053, 41054, 41055, 41059, 41196, 41197

        Case 59414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CONTABIL_NAO_PREENCHIDA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162460)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub CTBUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_CTBUpDown_DownClick

    If Len(CTBDataContabil.ClipText) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(CTBDataContabil, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 69003

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CTBUpDown_DownClick:

    Select Case gErr

        Case 69003

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162461)

    End Select

    Exit Sub

End Sub

Private Sub CTBUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_CTBUpDown_UpClick

    If Len(CTBDataContabil.ClipText) = 0 Then Exit Sub

    lErro = lErro = Data_Up_Down_Click(CTBDataContabil, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 69004

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CTBUpDown_UpClick:

    Select Case gErr

        Case 69004

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162462)

    End Select

    Exit Sub

End Sub

Private Sub ExibirLotesAtualizando_Click()

Dim colInvLote As New Collection
Dim lErro As Long

On Error GoTo Error_ExibirLotesAtualizando_Click

    'limpa o grid
    Call Grid_Limpa(objGrid)

    If ExibirLotesAtualizando.Value = 1 Then

        'le todos os lotes
        lErro = CF("InvLotePendente_Le_Desatualizados",giFilialEmpresa, colInvLote, LOTES_PENDENTES)
        If lErro <> SUCESSO Then Error 41198

    Else

        'leitura dos Lotes com IdAtualizacao = LOTE_DESATUALIZADO
        lErro = CF("InvLotePendente_Le_Desatualizados",giFilialEmpresa, colInvLote, LOTE_DESATUALIZADO)
        If lErro <> SUCESSO Then Error 41199

    End If

    'preenche o grid com os dados da colecao colInvLote
    lErro = Grid_Preenche(colInvLote)
    If lErro <> SUCESSO Then Error 41200

    Exit Sub

Error_ExibirLotesAtualizando_Click:

    Select Case Err

        Case 41198, 41199, 41200

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162463)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim colInvLote As New Collection

On Error GoTo Erro_Form_Load

    Set objGrid = New AdmGrid

    'Leitura dos Lotes desatualizados
    lErro = CF("InvLotePendente_Le_Desatualizados",giFilialEmpresa, colInvLote, LOTE_DESATUALIZADO)
    If lErro <> SUCESSO Then Error 41231

    'se não encontrou lote desatualizado ==> pesquisa se há lotes em processo de atualização
    If colInvLote.Count = 0 Then

        'Tenta ler um invlote pendente
        lErro = CF("InvLotePendente_Le1",giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 41243 Then Error 41244

        'Não encontrou nenhum lote pendente
        If lErro = 41243 Then Error 41245

        'Encontrou algum lote atualizando ==> avisa que só existem lotes sendo atualizados
        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_LOTE_ATUALIZANDO", giFilialEmpresa)

    End If

    'Inicializacao do grid
    Call Inicializa_Grid_LotesPendentes

    'Preenche o grid com os dados da colecao
    lErro = Grid_Preenche(colInvLote)
    If lErro <> SUCESSO Then Error 41246
    
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        'Inicialização da parte de contabilidade
        lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE)
        If lErro <> SUCESSO Then Error 59412
        
        lErro = objContabil.Contabil_Gera_Cabecalho_Automatico
        If lErro <> SUCESSO Then Error 59413
        
    Else
        
        CTBDataContabil.Enabled = False
        LabelDataContabil.Enabled = False
    
    End If
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 41231, 41244, 41246, 59412, 59413

        Case 41245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_LOTE_PENDENTE", Err, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162464)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Atualiza_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Atualiza_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Atualiza_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Atualiza
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid = Nothing
    
    Set objEventoLote = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
End Sub

Private Sub Lote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Descricao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NumLancAtual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub NumLancAtual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub NumLancAtual_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = NumLancAtual
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Status_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Status_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Status_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = Status
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridLotesPendentes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Private Sub GridLotesPendentes_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridLotesPendentes_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridLotesPendentes_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridLotesPendentes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Private Sub GridLotesPendentes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Private Sub GridLotesPendentes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridLotesPendentes_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridLotesPendentes_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case COL_ATUALIZA

                lErro = Saida_Celula_Atualiza(objGridInt)
                If lErro <> SUCESSO Then Error 41217

            Case COL_LOTE

                lErro = Saida_Celula_Lote(objGridInt)
                If lErro <> SUCESSO Then Error 41218

            Case COL_DESCRICAO

                lErro = Saida_Celula_Descricao(objGridInt)
                If lErro <> SUCESSO Then Error 41219

            Case COL_NUMLANCATUAL

                lErro = Saida_Celula_NumLancAtual(objGridInt)
                If lErro <> SUCESSO Then Error 41220

            Case COL_STATUS

                lErro = Saida_Celula_Status(objGridInt)
                If lErro <> SUCESSO Then Error 41221

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 41222

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 41217, 41218, 41219, 41220, 41221
        
        Case 41222
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162465)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Atualiza(objGridInt As AdmGrid) As Long
'faz a critica da celula(checkbox) atualiza do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Atualiza

    Set objGridInt.objControle = Atualiza

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41223

    Saida_Celula_Atualiza = SUCESSO

    Exit Function

Erro_Saida_Celula_Atualiza:

    Saida_Celula_Atualiza = Err

    Select Case Err

        Case 41223
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162466)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'faz a critica da celula Descrição do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = Descricao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41224

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = Err

    Select Case Err

        Case 41224
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162467)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NumLancAtual(objGridInt As AdmGrid) As Long
'faz a critica da celula NumLancAtual do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NumLancAtual

    Set objGridInt.objControle = NumLancAtual

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41225

    Saida_Celula_NumLancAtual = SUCESSO

    Exit Function

Erro_Saida_Celula_NumLancAtual:

    Saida_Celula_NumLancAtual = Err

    Select Case Err

        Case 41225
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162468)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long
'faz a critica da celula Lote do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41226

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = Err

    Select Case Err

        Case 41226
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162469)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Status(objGridInt As AdmGrid) As Long
'faz a critica da celula Status do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Status

    Set objGridInt.objControle = Status

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41227

    Saida_Celula_Status = SUCESSO

Exit Function

Erro_Saida_Celula_Status:

    Saida_Celula_Status = Err

    Select Case Err

        Case 41227
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162470)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_LotesPendentes() As Long

    'tela em questão
    Set objGrid.objForm = Me

    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Atualiza")
    objGrid.colColuna.Add ("Lote")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Numero Lanc. Atual")
    objGrid.colColuna.Add ("Status")

   'campos de edição do grid
    objGrid.colCampo.Add (Atualiza.Name)
    objGrid.colCampo.Add (Lote.Name)
    objGrid.colCampo.Add (Descricao.Name)
    objGrid.colCampo.Add (NumLancAtual.Name)
    objGrid.colCampo.Add (Status.Name)

    objGrid.objGrid = GridLotesPendentes

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid.iLinhasVisiveis = 10

    objGrid.objGrid.ColWidth(0) = 1000

    objGrid.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    lErro_Chama_Tela = SUCESSO

    Inicializa_Grid_LotesPendentes = SUCESSO

End Function

Private Function Grid_Preenche(colInvLote As Collection) As Long
'preenche o GridLotesPendentes e as duas listbox invisiveis, com os dados da colecao colInvLote

Dim lErro As Long
Dim iLinha As Integer
Dim objInvLote As ClassInvLote

On Error GoTo Erro_Grid_Preenche

    If colInvLote.Count < 10 Then
        objGrid.objGrid.Rows = 11
    Else
        objGrid.objGrid.Rows = colInvLote.Count + 1
    End If

    objGrid.iLinhasExistentes = colInvLote.Count

    iLinha = 1

    'pega cada objeto da colecao para fazer os preenchimentos
    For Each objInvLote In colInvLote

        'coloca o Lote no grid da tela
        GridLotesPendentes.TextMatrix(iLinha, COL_LOTE) = CStr(objInvLote.iLote)

        'coloca a Descrição no grid da tela
        GridLotesPendentes.TextMatrix(iLinha, COL_DESCRICAO) = objInvLote.sDescricao

        'coloca o Número de lançamentos atuais no grid da tela
        GridLotesPendentes.TextMatrix(iLinha, COL_NUMLANCATUAL) = objInvLote.iNumItensAtual

        If objInvLote.iIDAtualizacao <> 0 Then

            'coloca o Status = Atualizando no grid da tela
            GridLotesPendentes.TextMatrix(iLinha, COL_STATUS) = LOTE_ATUALIZANDO_TEXTO

        End If

        iLinha = iLinha + 1

    Next

    Call Grid_Inicializa(objGrid)

    Grid_Preenche = SUCESSO

    Exit Function

Erro_Grid_Preenche:

    Grid_Preenche = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162471)

    End Select

    Exit Function

End Function

Private Function GridLotesPendentes_Le(colInvLote As Collection) As Long
'Le o Grid onde a checkbox Atualiza = 1, colocando os dados na colecao

Dim iIndice As Integer
Dim iTemLote As Integer
Dim objInvLote As ClassInvLote
Dim lErro As Long

On Error GoTo Error_GridLotesPendentes_Le

    iTemLote = 0

    'Percorre todas as linhas do grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'seleciona os registros marcados na checkbox atualiza
        If GridLotesPendentes.TextMatrix(iIndice, COL_ATUALIZA) = "1" Then

            iTemLote = 1

            Set objInvLote = New ClassInvLote

                objInvLote.iFilialEmpresa = giFilialEmpresa

                'insere a descricao no objeto
                objInvLote.sDescricao = GridLotesPendentes.TextMatrix(iIndice, COL_DESCRICAO)

                'insere o número de lançamentos atual no objeto
                objInvLote.iNumItensAtual = CInt(GridLotesPendentes.TextMatrix(iIndice, COL_NUMLANCATUAL))

                'insere o lote no objeto
                objInvLote.iLote = CInt(GridLotesPendentes.TextMatrix(iIndice, COL_LOTE))

                If GridLotesPendentes.TextMatrix(iIndice, COL_STATUS) = LOTE_ATUALIZANDO_TEXTO Then
                    objInvLote.iIDAtualizacao = LOTE_ATUALIZANDO
                Else
                    objInvLote.iIDAtualizacao = LOTE_DESATUALIZADO
                End If

                'adiciona o objeto a colecao
                colInvLote.Add objInvLote

        End If

    Next

    If iTemLote = 0 Then Error 41230

    GridLotesPendentes_Le = SUCESSO

    Exit Function

Error_GridLotesPendentes_Le:

    GridLotesPendentes_Le = Err

    Select Case Err

        Case 41230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162472)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LOTE_EST_ATUALIZA
    Set Form_Load_Ocx = Me
    Caption = "Processamento dos Lotes do Inventário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LoteEstAtualiza"
    
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

'**** fim do trecho a ser copiado *****

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
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
'Traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub LabelDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataContabil, Button, Shift, X, Y)
End Sub

Private Sub LabelDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataContabil, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
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

Private Sub CTBLabelLote3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote3, Source, X, Y)
End Sub

Private Sub CTBLabelLote3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote3, Button, Shift, X, Y)
End Sub

