VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl Maquinas 
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
      Caption         =   "Frame5"
      Height          =   5055
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   810
      Visible         =   0   'False
      Width           =   9270
      Begin VB.Frame Frame8 
         Caption         =   "Insumos"
         Height          =   4935
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   9060
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   600
            TabIndex        =   40
            Top             =   975
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox UMTempo 
            Height          =   315
            Left            =   6720
            TabIndex        =   39
            Top             =   1200
            Width           =   975
         End
         Begin MSMask.MaskEdBox BarraSeparadora 
            Height          =   315
            Left            =   6600
            TabIndex        =   38
            Top             =   1200
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin VB.ComboBox UMProduto 
            Height          =   315
            Left            =   5640
            TabIndex        =   36
            Top             =   1200
            Width           =   975
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   4620
            TabIndex        =   35
            Top             =   1470
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1665
            TabIndex        =   34
            Top             =   1290
            Width           =   3000
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
            Height          =   375
            Left            =   165
            TabIndex        =   16
            ToolTipText     =   "Abre o Browse de Produtos"
            Top             =   4470
            Width           =   1380
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1995
            Left            =   165
            TabIndex        =   15
            Top             =   420
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   3519
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label LabelConversao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Taxa de Consumo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6120
            TabIndex        =   37
            Top             =   120
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5070
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   765
      Width           =   9255
      Begin VB.Frame Frame5 
         Caption         =   "Outros"
         Height          =   1050
         Left            =   60
         TabIndex        =   55
         Top             =   3405
         Width           =   9150
         Begin MSMask.MaskEdBox CustoHora 
            Height          =   315
            Left            =   1620
            TabIndex        =   10
            Top             =   225
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Prod 
            Height          =   315
            Left            =   1620
            TabIndex        =   12
            Top             =   645
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Peso 
            Height          =   285
            Left            =   4605
            TabIndex        =   11
            Top             =   195
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin VB.Label Label16 
            Caption         =   "Kg"
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
            Left            =   5910
            TabIndex        =   60
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
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
            Left            =   4005
            TabIndex        =   59
            Top             =   240
            Width           =   495
         End
         Begin VB.Label LabelCustoHora 
            Caption         =   "Custo p/ hora:"
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
            Left            =   345
            TabIndex        =   58
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label LabelProduto 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   585
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   57
            Top             =   675
            Width           =   870
         End
         Begin VB.Label DescProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4080
            TabIndex        =   56
            Top             =   645
            Width           =   5010
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Medidas (em metros)"
         Height          =   720
         Left            =   60
         TabIndex        =   51
         Top             =   2550
         Width           =   9150
         Begin MSMask.MaskEdBox Comprimento 
            Height          =   285
            Left            =   1620
            TabIndex        =   7
            Top             =   300
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Largura 
            Height          =   285
            Left            =   7230
            TabIndex        =   9
            Top             =   270
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Espessura 
            Height          =   285
            Left            =   4590
            TabIndex        =   8
            Top             =   285
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00######"
            PromptChar      =   " "
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Espessura:"
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
            Left            =   3555
            TabIndex        =   54
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Largura:"
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
            Left            =   6435
            TabIndex        =   53
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Comprimento:"
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
            TabIndex        =   52
            Top             =   345
            Width           =   1155
         End
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
         Height          =   435
         Left            =   60
         TabIndex        =   49
         ToolTipText     =   "Abre Browse com os CTs onde esta Máquina está vinculada"
         Top             =   4590
         Width           =   1905
      End
      Begin VB.CommandButton BotaoTaxas 
         Caption         =   "Taxas de Produção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2115
         TabIndex        =   50
         ToolTipText     =   "Abre Browse das Taxas de Produção para esta Máquina"
         Top             =   4590
         Width           =   1905
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tempos Padrão (Horas)"
         Height          =   690
         Left            =   60
         TabIndex        =   24
         Top             =   1725
         Width           =   9120
         Begin MSMask.MaskEdBox TempoMovimentacao 
            Height          =   315
            Left            =   4590
            TabIndex        =   5
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TempoPreparacao 
            Height          =   315
            Left            =   1605
            TabIndex        =   4
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TempoDescarga 
            Height          =   315
            Left            =   7275
            TabIndex        =   6
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin VB.Label LabelTempoDescarga 
            Caption         =   "Descarga:"
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
            Left            =   6330
            TabIndex        =   31
            Top             =   330
            Width           =   870
         End
         Begin VB.Label LabelTempoPreparacao 
            Caption         =   "Preparação:"
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
            Left            =   450
            TabIndex        =   30
            Top             =   330
            Width           =   1095
         End
         Begin VB.Label LabelTempoMovimentacao 
            Caption         =   "Movimentação:"
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
            Left            =   3210
            TabIndex        =   29
            Top             =   330
            Width           =   1350
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1500
         Left            =   60
         TabIndex        =   23
         Top             =   90
         Width           =   9120
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2565
            Picture         =   "Maquinas.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Numeração Automática"
            Top             =   315
            Width           =   300
         End
         Begin VB.ComboBox Recurso 
            Height          =   315
            ItemData        =   "Maquinas.ctx":00EA
            Left            =   5745
            List            =   "Maquinas.ctx":00EC
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   300
            Width           =   3330
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1635
            TabIndex        =   0
            Top             =   300
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   1635
            TabIndex        =   2
            Top             =   690
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   315
            Left            =   1635
            TabIndex        =   3
            Top             =   1080
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label LabelTipo 
            Caption         =   "Tipo de Recurso:"
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
            Height          =   315
            Left            =   4140
            TabIndex        =   28
            Top             =   345
            Width           =   1680
         End
         Begin VB.Label LabelDescricao 
            Caption         =   "Descrição:"
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
            Height          =   315
            Left            =   645
            TabIndex        =   27
            Top             =   1110
            Width           =   945
         End
         Begin VB.Label LabelNomeReduzido 
            Caption         =   "Nome Reduzido:"
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
            Height          =   315
            Left            =   150
            TabIndex        =   26
            Top             =   720
            Width           =   1410
         End
         Begin VB.Label LabelCodigo 
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
            Height          =   315
            Left            =   885
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   330
            Width           =   690
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5055
      Index           =   3
      Left            =   165
      TabIndex        =   41
      Top             =   795
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Frame Frame4 
         Caption         =   "Tipos de Mão-de-Obra"
         Height          =   4905
         Left            =   75
         TabIndex        =   42
         Top             =   15
         Width           =   9060
         Begin MSMask.MaskEdBox PercentualUso 
            Height          =   315
            Left            =   6660
            TabIndex        =   47
            Top             =   1740
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   6
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeTipoMO 
            Height          =   315
            Left            =   5415
            TabIndex        =   48
            Top             =   1740
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   3
            Format          =   "###"
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoTipoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1995
            TabIndex        =   46
            Top             =   1410
            Width           =   3990
         End
         Begin VB.TextBox CodigoTipoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   525
            MaxLength       =   20
            TabIndex        =   45
            Top             =   1725
            Width           =   1395
         End
         Begin VB.CommandButton BotaoTipoDeMaodeObra 
            Caption         =   "Tipo de Mão-de-Obra"
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
            Left            =   165
            TabIndex        =   43
            ToolTipText     =   "Abre o Browse de Tipos de Mão de Obra"
            Top             =   4425
            Width           =   2145
         End
         Begin MSFlexGridLib.MSFlexGrid GridTiposDeMaodeObra 
            Height          =   1995
            Left            =   150
            TabIndex        =   44
            Top             =   315
            Width           =   8760
            _ExtentX        =   15452
            _ExtentY        =   3519
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1575
         Picture         =   "Maquinas.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Maquinas.ctx":026C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "Maquinas.ctx":079E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "Maquinas.ctx":0928
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5505
      Left            =   75
      TabIndex        =   14
      Top             =   405
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   9710
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Consumos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Operadores"
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
Attribute VB_Name = "Maquinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim iFrameAtual As Integer

Dim objGrid1 As AdmGrid
Dim iLinhaAntiga As Integer

'Grid de Consumo de Insumos
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_UMProduto_Col As Integer
Dim iGrid_BarraSeparadora_Col As Integer
Dim iGrid_UMTempo_Col As Integer

'Grid de Tipos de Mão-de-Obra
Dim objGridTiposDeMaodeObra As AdmGrid
Dim iGrid_CodigoTipoMO_Col As Integer
Dim iGrid_DescricaoTipoMO_Col As Integer
Dim iGrid_QuantidadeTipoMO_Col As Integer
Dim iGrid_PercentualUso_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoTipoDeMaodeObra As AdmEvento
Attribute objEventoTipoDeMaodeObra.VB_VarHelpID = -1
Private WithEvents objEventoProd As AdmEvento
Attribute objEventoProd.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Máquinas, Habilidades e Processos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Maquinas"

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


Private Sub BotaoCT_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection
Dim objMaquinas As ClassMaquinas
Dim sFiltro As String

On Error GoTo Erro_BotaoTaxas_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 137924
    
    Set objMaquinas = New ClassMaquinas
    
    objMaquinas.iFilialEmpresa = giFilialEmpresa
    objMaquinas.iCodigo = StrParaInt(Codigo.Text)
    
    'Verifica se a Máquina existe, lendo no BD a partir do Código
    lErro = CF("Maquinas_Le", objMaquinas)
    If lErro <> SUCESSO And lErro <> 103090 Then gError 137925
        
    If lErro <> SUCESSO Then gError 137926
    
    sFiltro = "NumIntDoc IN (SELECT NumIntDocCT FROM CTMaquinas WHERE NumIntDocMaq = ?)"
    colSelecao.Add objMaquinas.lNumIntDoc

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, Nothing, sFiltro)

    Exit Sub

Erro_BotaoTaxas_Click:

    Select Case gErr
    
        Case 137924
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 137926
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINAS_NAO_CADASTRADO", gErr, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    
        Case 137925
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162582)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTaxas_Click()

Dim lErro As Long
Dim objTaxaDeProducao As New ClassTaxaDeProducao
Dim colSelecao As New Collection
Dim objMaquinas As ClassMaquinas
Dim sFiltro As String

On Error GoTo Erro_BotaoTaxas_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 137927
    
    Set objMaquinas = New ClassMaquinas
    
    objMaquinas.iFilialEmpresa = giFilialEmpresa
    objMaquinas.iCodigo = StrParaInt(Codigo.Text)
    
    'Verifica se a Máquina existe, lendo no BD a partir do Código
    lErro = CF("Maquinas_Le", objMaquinas)
    If lErro <> SUCESSO And lErro <> 103090 Then gError 137928
        
    If lErro <> SUCESSO Then gError 137929

    sFiltro = "Ativo = ? And (NumIntDocMaq = ? Or NumIntDocMaq = ? )"
    colSelecao.Add TAXA_ATIVA
    colSelecao.Add objMaquinas.lNumIntDoc
    colSelecao.Add 0
    
    Call Chama_Tela("TaxaDeProducaoLista", colSelecao, objTaxaDeProducao, Nothing, sFiltro)

    Exit Sub

Erro_BotaoTaxas_Click:

    Select Case gErr
    
        Case 137927
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 137929
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINAS_NAO_CADASTRADO", gErr, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    
        Case 137928
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162583)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTipoDeMaodeObra_Click()

Dim lErro As Long
Dim objTiposDeMaodeObras As New ClassTiposDeMaodeObra
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    If Me.ActiveControl Is CodigoTipoMO Then
    
        objTiposDeMaodeObras.iCodigo = StrParaInt(CodigoTipoMO.Text)
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridTiposDeMaodeObra.Row = 0 Then gError 134380

        objTiposDeMaodeObras.iCodigo = StrParaInt(GridTiposDeMaodeObra.TextMatrix(GridTiposDeMaodeObra.Row, iGrid_CodigoTipoMO_Col))
        
    End If

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTiposDeMaodeObras, objEventoTipoDeMaodeObra)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
        
        Case 134380
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162584)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
        
End Sub

Private Sub GridItens_LostFocus()

    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub objEventoTipoDeMaodeObra_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTiposDeMaodeObra = obj1

    'Verifica se há algum produto repetido no grid
    For iLinha = 1 To objGridTiposDeMaodeObra.iLinhasExistentes
        
        If iLinha < GridTiposDeMaodeObra.Row Then
                                                
            If GridTiposDeMaodeObra.TextMatrix(iLinha, iGrid_CodigoTipoMO_Col) = objTiposDeMaodeObra.iCodigo Then
                CodigoTipoMO.Text = ""
                gError 134121
                
            End If
                
        End If
                       
    Next
    
    CodigoTipoMO.Text = CStr(objTiposDeMaodeObra.iCodigo)
    
    If Not (Me.ActiveControl Is CodigoTipoMO) Then
    
        GridTiposDeMaodeObra.TextMatrix(GridTiposDeMaodeObra.Row, iGrid_CodigoTipoMO_Col) = CStr(objTiposDeMaodeObra.iCodigo)
        GridTiposDeMaodeObra.TextMatrix(GridTiposDeMaodeObra.Row, iGrid_DescricaoTipoMO_Col) = objTiposDeMaodeObra.sDescricao
        GridTiposDeMaodeObra.TextMatrix(GridTiposDeMaodeObra.Row, iGrid_PercentualUso_Col) = Format(1, "Percent")
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridTiposDeMaodeObra.Row - GridTiposDeMaodeObra.FixedRows = objGridTiposDeMaodeObra.iLinhasExistentes Then
            objGridTiposDeMaodeObra.iLinhasExistentes = objGridTiposDeMaodeObra.iLinhasExistentes + 1
        End If
        
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 134121
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, objTiposDeMaodeObra.iCodigo, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162585)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Sub Recurso_Click()

    iAlterado = REGISTRO_ALTERADO
    Call Seleciona_Recurso(Codigo_Extrai(Recurso.Text))

End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
    
        If Me.ActiveControl Is Produto Then Call BotaoProdutos_Click
    
        If Me.ActiveControl Is CodigoTipoMO Then Call BotaoTipoDeMaodeObra_Click
        
        If Me.ActiveControl Is Prod Then Call LabelProduto_Click
    
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
        
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoTipoDeMaodeObra = Nothing
    Set objEventoProd = Nothing
    
    Set objGrid1 = Nothing
    Set objGridTiposDeMaodeObra = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162586)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoTipoDeMaodeObra = New AdmEvento
    Set objEventoProd = New AdmEvento
    
    iFrameAtual = 1
    
    'Grid Itens
    Set objGrid1 = New AdmGrid
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    lErro = Inicializa_GridItens(objGrid1)
    If lErro <> SUCESSO Then gError 134076
    
    'Grid Tipos de MO
    Set objGridTiposDeMaodeObra = New AdmGrid
    
    'tela em questão
    Set objGridTiposDeMaodeObra.objForm = Me
    
    lErro = Inicializa_GridTiposDeMaodeObra(objGridTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 134340
    
    lErro = CarregaComboRecursos(Recurso)
    If lErro <> SUCESSO Then gError 134077
    
    Recurso.ListIndex = 0
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 134078
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Prod)
    If lErro <> SUCESSO Then gError 134078
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 134076, 134077, 134078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162587)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objMaquinas As ClassMaquinas) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objMaquinas Is Nothing) Then

        lErro = Traz_Maquinas_Tela(objMaquinas)
        If lErro <> SUCESSO And lErro <> 134095 And lErro <> 134097 Then gError 134079
                    
        If lErro <> SUCESSO Then
                
            If objMaquinas.iCodigo > 0 Then
                    
                'Coloca o código da Maquina na tela
                Codigo.Text = objMaquinas.iCodigo
                        
            ElseIf Len(Trim(objMaquinas.sNomeReduzido)) > 0 Then
                    
                'Coloca o NomeReduzido da Maquina na tela
                NomeReduzido.Text = objMaquinas.sNomeReduzido
                    
            End If
    
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 134079

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162588)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objMaquinasInsumos As ClassMaquinasInsumos
Dim objMaquinaOperadores As ClassMaquinaOperadores
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objMaquinas.iCodigo = StrParaInt(Codigo.Text)
    objMaquinas.iFilialEmpresa = giFilialEmpresa
    objMaquinas.sNomeReduzido = NomeReduzido.Text
    objMaquinas.sDescricao = Descricao.Text
    objMaquinas.dTempoMovimentacao = StrParaDbl(TempoMovimentacao.Text)
    objMaquinas.dTempoPreparacao = StrParaDbl(TempoPreparacao.Text)
    objMaquinas.dTempoDescarga = StrParaDbl(TempoDescarga.Text)
    objMaquinas.dCustoHora = StrParaDbl(CustoHora.Text)
    objMaquinas.iRecurso = Codigo_Extrai(Recurso.Text)
    
    objMaquinas.dPeso = StrParaDbl(Peso.Text)
    objMaquinas.dComprimento = StrParaDbl(Comprimento.Text)
    objMaquinas.dLargura = StrParaDbl(Largura.Text)
    objMaquinas.dEspessura = StrParaDbl(Espessura.Text)
    
    lErro = CF("Produto_Formata", Prod.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134080

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objMaquinas.sProduto = sProdutoFormatado
    
    'Ir preenchendo a colecao no objMaquina com todas as linhas "existentes" do grid Insumos
    For iIndice = 1 To objGrid1.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then Exit For
        
        Set objProdutos = New ClassProduto
        
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 134080
        
        objProdutos.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134081
        
        Set objMaquinasInsumos = New ClassMaquinasInsumos
        
        objMaquinasInsumos.lNumIntDocMaq = objMaquinas.lNumIntDoc
        objMaquinasInsumos.sProduto = objProdutos.sCodigo
        objMaquinasInsumos.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objMaquinasInsumos.sUMProduto = GridItens.TextMatrix(iIndice, iGrid_UMProduto_Col)
        objMaquinasInsumos.sUMTempo = GridItens.TextMatrix(iIndice, iGrid_UMTempo_Col)
    
        objMaquinas.colProdutos.Add objMaquinasInsumos
    
    Next
    
    'Ir preenchendo a colecao no objMaquina com todas as linhas "existentes" do grid Operadores
    For iIndice = 1 To objGridTiposDeMaodeObra.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_CodigoTipoMO_Col))) = 0 Then Exit For
                
        Set objMaquinaOperadores = New ClassMaquinaOperadores
        
        objMaquinaOperadores.iTipoMaoDeObra = StrParaInt(GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_CodigoTipoMO_Col))
        objMaquinaOperadores.lNumIntDocMaq = objMaquinas.lNumIntDoc
        objMaquinaOperadores.iQuantidade = StrParaInt(GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_QuantidadeTipoMO_Col))
        objMaquinaOperadores.dPercentualUso = StrParaDbl(Val(GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_PercentualUso_Col)) / 100)
    
        objMaquinas.colTipoOperadores.Add objMaquinaOperadores
    
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 134080, 134081
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162589)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Maquinas"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objMaquinas)
    If lErro <> SUCESSO Then gError 134082

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objMaquinas.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objMaquinas.sDescricao, STRING_MAQUINA_DESCRICAO, "Descricao"
    colCampoValor.Add "NomeReduzido", objMaquinas.sNomeReduzido, STRING_MAQUINA_NOMEREDUZIDO, "NomeReduzido"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 134082
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162590)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Tela_Preenche

    objMaquinas.iCodigo = colCampoValor.Item("Codigo").vValor
    objMaquinas.sDescricao = colCampoValor.Item("Descricao").vValor
    objMaquinas.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
    
    objMaquinas.iFilialEmpresa = giFilialEmpresa

    If objMaquinas.iCodigo <> 0 And objMaquinas.iFilialEmpresa <> 0 Then
        lErro = Traz_Maquinas_Tela(objMaquinas)
        If lErro <> SUCESSO Then gError 134083
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 134083

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162591)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 134084
        
    'Verifica se o NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 134085
    
    'Verifica se a Descrição está preenchido
    If Len(Trim(Descricao.Text)) = 0 Then gError 134086

    'Verifica se o Recurso está preenchido
    If Len(Trim(Recurso.Text)) = 0 Then gError 134087
    
    'Para cada MaquinasInsumos
    For iIndice = 1 To objGrid1.iLinhasExistentes
        
        'Verifica se a Quantidade foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 134088

        'Verifica se a Unidade de Medida do Produto foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UMProduto_Col))) = 0 Then gError 134089
        
        'Verifica se a Unidade de Medida de Tempo foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UMTempo_Col))) = 0 Then gError 134090

    Next
            
    'Para cada MaquinaOperadores
    For iIndice = 1 To objGridTiposDeMaodeObra.iLinhasExistentes
        
        'Verifica se a Quantidade foi informada
        If Len(Trim(GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_QuantidadeTipoMO_Col))) = 0 Then gError 134653

        'Verifica se o Percentual de Uso foi preenchido
        If Len(Trim(GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_PercentualUso_Col))) = 0 Then gError 134654
                
    Next
                        
    'Preenche o objMaquinas
    lErro = Move_Tela_Memoria(objMaquinas)
    If lErro <> SUCESSO Then gError 134091

    lErro = Trata_Alteracao(objMaquinas, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 134092

    'Grava o/a Maquinas no Banco de Dados
    lErro = CF("Maquina_Grava", objMaquinas)
    If lErro <> SUCESSO Then gError 134093
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134084
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 134085
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
        
        Case 134086
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
                                    
        Case 134087
            Call Rotina_Erro(vbOKOnly, "ERRO_RECURSO_NAO_PREENCHIDO", gErr)
        
        Case 134088, 134653
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 134654
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUALUSO_NAO_PREENCHIDO", gErr)
        
        Case 134089
            Call Rotina_Erro(vbOKOnly, "ERRO_UMEDIDA_PRODUTO_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 134090
            Call Rotina_Erro(vbOKOnly, "ERRO_UMEDIDA_TEMPO_NAO_PREENCHIDA", gErr, iIndice)
                
        Case 134091, 134092, 134093
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162592)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Maquinas() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Maquinas
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    Recurso.ListIndex = 0
    
    DescProd.Caption = ""

    Call Grid_Limpa(objGrid1)
    Call Grid_Limpa(objGridTiposDeMaodeObra)
    
    iAlterado = 0

    Limpa_Tela_Maquinas = SUCESSO

    Exit Function

Erro_Limpa_Tela_Maquinas:

    Limpa_Tela_Maquinas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162593)

    End Select

    Exit Function

End Function

Function Traz_Maquinas_Tela(objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutos As ClassProduto
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim sProdutoMascarado As String

On Error GoTo Erro_Traz_Maquinas_Tela

    If objMaquinas.iCodigo > 0 Then
        
        objMaquinas.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se a Máquina existe, lendo no BD a partir do Código
        lErro = CF("Maquinas_Le", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103090 Then gError 134094
            
        If lErro <> SUCESSO Then gError 134095
            
    ElseIf Len(Trim(objMaquinas.sNomeReduzido)) > 0 Then
                
        'Verifica se a Máquina existe, lendo no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 134096
    
        If lErro <> SUCESSO Then gError 134097
    
    End If

    'Limpa a tela
    Call Limpa_Tela_Maquinas
        
    If objMaquinas.iCodigo <> 0 Then Codigo.Text = CStr(objMaquinas.iCodigo)
    NomeReduzido.Text = objMaquinas.sNomeReduzido
    Descricao.Text = objMaquinas.sDescricao
    If objMaquinas.dTempoMovimentacao <> 0 Then TempoMovimentacao.Text = Formata_Estoque(objMaquinas.dTempoMovimentacao)
    If objMaquinas.dTempoPreparacao <> 0 Then TempoPreparacao.Text = Formata_Estoque(objMaquinas.dTempoPreparacao)
    If objMaquinas.dTempoDescarga <> 0 Then TempoDescarga.Text = Formata_Estoque(objMaquinas.dTempoDescarga)
    If objMaquinas.dCustoHora <> 0 Then CustoHora.Text = Format(objMaquinas.dCustoHora, "STANDARD")

    If objMaquinas.dPeso > 0 Then Peso.Text = Format(objMaquinas.dPeso, Peso.Format)
    If objMaquinas.dComprimento > 0 Then Comprimento.Text = Format(objMaquinas.dComprimento, Comprimento.Format)
    If objMaquinas.dLargura > 0 Then Largura.Text = Format(objMaquinas.dLargura, Largura.Format)
    If objMaquinas.dEspessura > 0 Then Espessura.Text = Format(objMaquinas.dEspessura, Espessura.Format)

    If Len(Trim(objMaquinas.sProduto)) > 0 Then
        lErro = Mascara_RetornaProdutoEnxuto(objMaquinas.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134100
        
        Prod.PromptInclude = False
        Prod.Text = sProdutoMascarado
        Prod.PromptInclude = True
        Call Prod_Validate(bSGECancelDummy)
    End If

    Call Combo_Seleciona_ItemData(Recurso, objMaquinas.iRecurso)
            
    Call Seleciona_Recurso(objMaquinas.iRecurso)
    
    'Limpa os Grids antes de colocar algo neles
    Call Grid_Limpa(objGrid1)
    Call Grid_Limpa(objGridTiposDeMaodeObra)
    
    'Lê o MaquinasInsumos que está sendo Passado
    lErro = CF("Maquinas_Le_Itens", objMaquinas)
    If lErro <> SUCESSO Then gError 134098
    
    'Exibe os dados da coleção de Produtos na tela (GridItens)
    For iIndice = 1 To objMaquinas.colProdutos.Count
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objMaquinas.colProdutos.Item(iIndice).sProduto
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134099
        
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134100
                                
        'Insere no Grid MaquinasInsumos
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridItens.TextMatrix(iIndice, iGrid_DescricaoProduto_Col) = objProdutos.sDescricao
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objMaquinas.colProdutos.Item(iIndice).dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_UMProduto_Col) = objMaquinas.colProdutos.Item(iIndice).sUMProduto
        GridItens.TextMatrix(iIndice, iGrid_BarraSeparadora_Col) = STRING_BARRA_SEPARADORA
        GridItens.TextMatrix(iIndice, iGrid_UMTempo_Col) = objMaquinas.colProdutos.Item(iIndice).sUMTempo

    Next

    objGrid1.iLinhasExistentes = objMaquinas.colProdutos.Count

    'Exibe os dados da coleção de Operadores na tela (GridTiposDeMaodeObra)
    For iIndice = 1 To objMaquinas.colTipoOperadores.Count
                
        Set objTiposDeMaodeObra = New ClassTiposDeMaodeObra
        
        objTiposDeMaodeObra.iCodigo = objMaquinas.colTipoOperadores.Item(iIndice).iTipoMaoDeObra
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 134655
    
        'Insere no GridTiposDeMaodeObra
        GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_CodigoTipoMO_Col) = CStr(objMaquinas.colTipoOperadores.Item(iIndice).iTipoMaoDeObra)
        GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_DescricaoTipoMO_Col) = objTiposDeMaodeObra.sDescricao
        GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_QuantidadeTipoMO_Col) = CStr(objMaquinas.colTipoOperadores.Item(iIndice).iQuantidade)
        GridTiposDeMaodeObra.TextMatrix(iIndice, iGrid_PercentualUso_Col) = Format(objMaquinas.colTipoOperadores.Item(iIndice).dPercentualUso, "Percent")
        
    Next

    objGridTiposDeMaodeObra.iLinhasExistentes = objMaquinas.colTipoOperadores.Count
        
    iAlterado = 0
    
    Traz_Maquinas_Tela = SUCESSO

    Exit Function

Erro_Traz_Maquinas_Tela:

    Traz_Maquinas_Tela = gErr

    Select Case gErr

        Case 134094, 134096, 134098, 134099, 134100, 134655
            'Erros tratados nas rotinas chamadas
        
        Case 134095, 134097 '134095 = Não encontrou por código; 134097 = Não encontrou por NomeReduzido
            'Erros tratados na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162594)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Grava a Máquina
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 134101

    'Limpa Tela
    Call Limpa_Tela_Maquinas
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 134101

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162595)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162596)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134102

    Call Limpa_Tela_Maquinas
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 134102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162597)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 134103

    objMaquinas.iCodigo = StrParaInt(Codigo.Text)
    objMaquinas.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_MAQUINAS", objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a requisição de consumo
    lErro = CF("Maquina_Exclui", objMaquinas)
    If lErro <> SUCESSO And lErro <> 137181 Then gError 134104

    If lErro = SUCESSO Then
    
        'Limpa Tela
        Call Limpa_Tela_Maquinas
        
    End If

    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134103
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 134104

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162598)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

        'Critica a Codigo
        lErro = Inteiro_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 134105

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 134105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162599)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_NomeReduzido_Validate

    'Veifica se NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) <> 0 Then

        objMaquinas.sNomeReduzido = NomeReduzido.Text

        lErro = Traz_Maquinas_Tela(objMaquinas)
        If lErro <> SUCESSO And lErro <> 134095 And lErro <> 134097 Then gError 137129

    End If

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137129
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162600)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Veifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162601)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoMovimentacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoMovimentacao_Validate

    'Verifica se TempoMovimentacao está preenchida
    If Len(Trim(TempoMovimentacao.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TempoMovimentacao.Text)
        If lErro <> SUCESSO Then gError 135045
        
        TempoMovimentacao.Text = Formata_Estoque(TempoMovimentacao.Text)

    End If

    Exit Sub

Erro_TempoMovimentacao_Validate:

    Cancel = True

    Select Case gErr

        Case 135045
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162602)

    End Select

    Exit Sub

End Sub

Private Sub TempoMovimentacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoMovimentacao, iAlterado)
    
End Sub

Private Sub TempoMovimentacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoPreparacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoPreparacao_Validate

    'Verifica se TempoPreparacao está preenchida
    If Len(Trim(TempoPreparacao.Text)) > 0 Then
        
        lErro = Valor_NaoNegativo_Critica(TempoPreparacao.Text)
        If lErro <> SUCESSO Then gError 134107
        
        TempoPreparacao.Text = Formata_Estoque(TempoPreparacao.Text)

    End If

    Exit Sub

Erro_TempoPreparacao_Validate:

    Cancel = True

    Select Case gErr

        Case 134107
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162603)

    End Select

    Exit Sub

End Sub

Private Sub TempoPreparacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoPreparacao, iAlterado)
    
End Sub

Private Sub TempoPreparacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoDescarga_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoDescarga_Validate

    'Verifica se TempoDescarga está preenchida
    If Len(Trim(TempoDescarga.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TempoDescarga.Text)
        If lErro <> SUCESSO Then gError 134108
        
        TempoDescarga.Text = Formata_Estoque(TempoDescarga.Text)

    End If

    Exit Sub

Erro_TempoDescarga_Validate:

    Cancel = True

    Select Case gErr

        Case 134108
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162604)

    End Select

    Exit Sub

End Sub

Private Sub TempoDescarga_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoDescarga, iAlterado)
    
End Sub

Private Sub TempoDescarga_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Recurso_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Recurso_Validate

    'Verifica se Recurso está preenchida
    If Len(Trim(Recurso.Text)) = 0 Then gError 137136

    Exit Sub

Erro_Recurso_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137136
            Call Rotina_Erro(vbOKOnly, "ERRO_RECURSO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162605)

    End Select

    Exit Sub

End Sub

Private Sub Recurso_Change()

    iAlterado = REGISTRO_ALTERADO
    Call Seleciona_Recurso(Codigo_Extrai(Recurso.Text))

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objMaquinas = obj1

    'Mostra os dados do Maquinas na tela
    lErro = Traz_Maquinas_Tela(objMaquinas)
    If lErro <> SUCESSO And lErro <> 134095 And lErro <> 134097 Then gError 134109

    'se não foi encontrado o código
    If lErro = 134095 Then gError 134110
    
    'se não foi encontrado o NomeReduzido
    If lErro = 134097 Then gError 134111

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 134109
        
        Case 134110
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_CODIGO_INEXISTENTE", gErr, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
        
        Case 134111
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NOME_INEXISTENTE", gErr, objMaquinas.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162606)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objMaquinas.iCodigo = Codigo.Text
        objMaquinas.sNomeReduzido = NomeReduzido.Text
        objMaquinas.sDescricao = Descricao.Text

    End If

    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162607)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("UM Prod.")
    objGrid.colColuna.Add ("/")
    objGrid.colColuna.Add ("UM Tempo")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (DescricaoProduto.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (UMProduto.Name)
    objGrid.colCampo.Add (BarraSeparadora.Name)
    objGrid.colCampo.Add (UMTempo.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescricaoProduto_Col = 2
    iGrid_Quantidade_Col = 3
    iGrid_UMProduto_Col = 4
    iGrid_BarraSeparadora_Col = 5
    iGrid_UMTempo_Col = 6

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()
    
    Call Saida_Celula(objGrid1)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid1)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGrid1)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then

            Select Case GridItens.Col
    
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 134112
                    
                Case iGrid_DescricaoProduto_Col
    
                    lErro = Saida_Celula_DescricaoProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 134113
                
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 134114
                
                Case iGrid_UMProduto_Col
    
                    lErro = Saida_Celula_UMProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 134115
                    
                Case iGrid_BarraSeparadora_Col
    
                    lErro = Saida_Celula_BarraSeparadora(objGridInt)
                    If lErro <> SUCESSO Then gError 134116
                
                Case iGrid_UMTempo_Col
    
                    lErro = Saida_Celula_UMTempo(objGridInt)
                    If lErro <> SUCESSO Then gError 134117
    
            End Select

        'Tipos de MO
        ElseIf objGridInt.objGrid.Name = GridTiposDeMaodeObra.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CodigoTipoMO_Col
                
                    lErro = Saida_Celula_CodigoTipoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 134370
    
                Case iGrid_DescricaoTipoMO_Col
    
                    lErro = Saida_Celula_DescricaoTipoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 134371
        
                Case iGrid_QuantidadeTipoMO_Col
    
                    lErro = Saida_Celula_QuantidadeTipoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
                Case iGrid_PercentualUso_Col
    
                    lErro = Saida_Celula_PercentualUso(objGridInt)
                    If lErro <> SUCESSO Then gError 137066
        
            End Select

        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 134118

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 134112 To 134117, 134370 To 134372, 137066

        Case 134118
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162608)

    End Select

    Exit Function

End Function

Private Function CarregaComboRecursos(objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CarregaComboRecursos

    
    objCombo.AddItem ITEMCT_RECURSO_MAQUINA & SEPARADOR & STRING_ITEMCT_RECURSO_MAQUINA
    objCombo.ItemData(objCombo.NewIndex) = ITEMCT_RECURSO_MAQUINA
    
    objCombo.AddItem ITEMCT_RECURSO_HABILIDADE & SEPARADOR & STRING_ITEMCT_RECURSO_HABILIDADE
    objCombo.ItemData(objCombo.NewIndex) = ITEMCT_RECURSO_HABILIDADE
    
    objCombo.AddItem ITEMCT_RECURSO_PROCESSO & SEPARADOR & STRING_ITEMCT_RECURSO_PROCESSO
    objCombo.ItemData(objCombo.NewIndex) = ITEMCT_RECURSO_PROCESSO
    
    CarregaComboRecursos = SUCESSO

    Exit Function

Erro_CarregaComboRecursos:

    CarregaComboRecursos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162609)

    End Select

    Exit Function

End Function

Private Sub Seleciona_Recurso(ByVal iRecurso As Integer)
        
    Select Case iRecurso
    
        Case ITEMCT_RECURSO_MAQUINA
        
            TempoDescarga.Text = ""
            TempoDescarga.Enabled = False
            TempoMovimentacao.Enabled = True
            TempoPreparacao.Enabled = True
            LabelTempoDescarga.Enabled = False
            LabelTempoMovimentacao.Enabled = True
            LabelTempoPreparacao.Enabled = True
        
        Case ITEMCT_RECURSO_HABILIDADE
        
            TempoDescarga.Text = ""
            TempoPreparacao.Text = ""
            TempoDescarga.Enabled = False
            TempoMovimentacao.Enabled = True
            TempoPreparacao.Enabled = False
            LabelTempoDescarga.Enabled = False
            LabelTempoMovimentacao.Enabled = True
            LabelTempoPreparacao.Enabled = False
        
        Case ITEMCT_RECURSO_PROCESSO
        
            TempoMovimentacao.Text = ""
            TempoDescarga.Enabled = True
            TempoMovimentacao.Enabled = False
            TempoPreparacao.Enabled = True
            LabelTempoDescarga.Enabled = True
            LabelTempoMovimentacao.Enabled = False
            LabelTempoPreparacao.Enabled = True
            
    End Select

End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProduto1 As String

On Error GoTo Erro_LabelProduto_Click
    
    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134338

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134119
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 134119

        Case 134338
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162610)

    End Select

    Exit Sub

End Sub
Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim sUnidadeMed As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 134120
    
    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True

    'Verifica se há algum produto repetido no grid
    For iLinha = 1 To objGrid1.iLinhasExistentes
        
        If iLinha <> GridItens.Row Then
                                                
            If GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text Then
                Produto.PromptInclude = False
                Produto.Text = ""
                Produto.PromptInclude = True
                gError 134121
                
            End If
                
        End If
                       
    Next
    
    If Not (Me.ActiveControl Is Produto) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = sProdutoMascarado
        GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
        GridItens.TextMatrix(GridItens.Row, iGrid_UMProduto_Col) = objProduto.sSiglaUMEstoque
        GridItens.TextMatrix(GridItens.Row, iGrid_BarraSeparadora_Col) = STRING_BARRA_SEPARADORA
        
        Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)
        GridItens.TextMatrix(GridItens.Row, iGrid_UMTempo_Col) = sUnidadeMed
        'GridItens.TextMatrix(GridItens.Row, iGrid_UMTempo_Col) = TAXA_CONSUMO_TEMPO_PADRAO
          
        Call Produto_Validate(bSGECancelDummy)
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If
        
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 134120
        
        Case 134121
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162611)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para uma Máquina
    lErro = CF("Maquina_Automatico", iCodigo)
    If lErro <> SUCESSO Then gError 134122
    
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 134122
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162612)
    
    End Select

    Exit Sub

End Sub

Private Sub UMProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub UMProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub UMProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = UMProduto
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMTempo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMTempo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub UMTempo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub UMTempo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = UMTempo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub BarraSeparadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BarraSeparadora_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub BarraSeparadora_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub BarraSeparadora_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = BarraSeparadora
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Function Saida_Celula_DescricaoProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoProduto

    Set objGridInt.objControle = DescricaoProduto
    
    'Se o campo foi preenchido
    If Len(Trim(DescricaoProduto.Text)) > 0 Then
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134123

    Saida_Celula_DescricaoProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoProduto:

    Saida_Celula_DescricaoProduto = gErr

    Select Case gErr

        Case 134123
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162613)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UMProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UMProduto

    Set objGridInt.objControle = UMProduto
    
    'Se o campo foi preenchido
    If Len(Trim(UMProduto.Text)) > 0 Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_UMProduto_Col) = UMProduto.Text
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134124

    Saida_Celula_UMProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_UMProduto:

    Saida_Celula_UMProduto = gErr

    Select Case gErr

        Case 134124
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162614)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UMTempo(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UMTempo

    Set objGridInt.objControle = UMTempo
    
    'Se o campo foi preenchido
    If Len(Trim(UMTempo.Text)) > 0 Then
                         
        GridItens.TextMatrix(GridItens.Row, iGrid_UMTempo_Col) = UMTempo.Text
        
        
                          
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134125

    Saida_Celula_UMTempo = SUCESSO

    Exit Function

Erro_Saida_Celula_UMTempo:

    Saida_Celula_UMTempo = gErr

    Select Case gErr

        Case 134125
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162615)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_BarraSeparadora(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_BarraSeparadora

    Set objGridInt.objControle = BarraSeparadora
    
    'Se o campo foi preenchido
    If Len(Trim(BarraSeparadora.Text)) > 0 Then
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134126

    Saida_Celula_BarraSeparadora = SUCESSO

    Exit Function

Erro_Saida_Celula_BarraSeparadora:

    Saida_Celula_BarraSeparadora = gErr

    Select Case gErr

        Case 134126
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162616)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    
    'Se o campo foi preenchido
    If Len(Trim(Quantidade.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 134337
        
        Quantidade.Text = Formata_Estoque(StrParaDbl(Quantidade.Text))

        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134127

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 134127, 134337
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162617)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodProduto As String
Dim iLinha As Integer
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
                
    sCodProduto = Produto.Text
        
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134128
    
    'Se o campo foi preenchido
    If Len(sProdutoFormatado) > 0 Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134129
                
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
                
        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridItens.Row Then
                                                    
                If GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text Then
                    Produto.PromptInclude = False
                    Produto.Text = ""
                    Produto.PromptInclude = True
                    gError 134130
                    
                End If
                    
            End If
                           
        Next
        
        Set objProdutos = New ClassProduto

        objProdutos.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134131
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoProduto_Col) = objProdutos.sDescricao
        GridItens.TextMatrix(GridItens.Row, iGrid_UMProduto_Col) = objProdutos.sSiglaUMEstoque
        GridItens.TextMatrix(GridItens.Row, iGrid_BarraSeparadora_Col) = STRING_BARRA_SEPARADORA
        'GridItens.TextMatrix(GridItens.Row, iGrid_UMTempo_Col) = TAXA_CONSUMO_TEMPO_PADRAO
        Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)
        GridItens.TextMatrix(GridItens.Row, iGrid_UMTempo_Col) = sUnidadeMed

        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134132

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 134130
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134128, 134129, 134131, 134132
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162618)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sCodProduto As String
Dim objProdutos As ClassProduto
Dim objClasseUM As ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sCodTipo As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Guardo o valor do Codigo do Item
    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    'Guardo o valor do Codigo do Tipo de MO
    sCodTipo = GridTiposDeMaodeObra.TextMatrix(GridTiposDeMaodeObra.Row, iGrid_CodigoTipoMO_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134133

    'Grid Itens
    If objControl.Name = "Produto" Then
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objControl.Enabled = False
            
        Else
            objControl.Enabled = True
        
        End If
        
    ElseIf objControl.Name = "DescricaoProduto" Then

        objControl.Enabled = False
                            
    ElseIf objControl.Name = "Quantidade" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "UMProduto" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objControl.Enabled = True

            Set objProdutos = New ClassProduto

            objProdutos.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProdutos)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 134134

            Set objClasseUM = New ClassClasseUM
            
            objClasseUM.iClasse = objProdutos.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO And lErro <> 22539 Then gError 134135

            'Se tem algum valor para UMProduto do Grid
            If Len(GridItens.TextMatrix(GridItens.Row, iGrid_UMProduto_Col)) > 0 Then
                'Guardo o valor da UMProduto da Linha
                sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UMProduto_Col)
            Else
                'Senão coloco o do Produto em estoque
                sUnidadeMed = objProdutos.sSiglaUMEstoque
            End If
            
            'Limpar as Unidades utilizadas anteriormente
            UMProduto.Clear

            For Each objUnidadeDeMedida In colSiglas
                UMProduto.AddItem objUnidadeDeMedida.sSigla
            Next

            UMProduto.AddItem ""

            'Tento selecionar na Combo a Unidade anterior
            If UMProduto.ListCount <> 0 Then

                For iIndice = 0 To UMProduto.ListCount - 1

                    If UMProduto.List(iIndice) = sUnidadeMed Then
                        UMProduto.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "BarraSeparadora" Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_BarraSeparadora_Col) = STRING_BARRA_SEPARADORA
    
        objControl.Enabled = False
    
    ElseIf objControl.Name = "UMTempo" Then
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objControl.Enabled = True
            
            Set objClasseUM = New ClassClasseUM
            
            objClasseUM.iClasse = gobjEST.iClasseUMTempo

            'Preenche a List da Combo UnidadeMed com as UM's de Tempo
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO And lErro <> 22539 Then gError 134136

            'Se tem algum valor para UMTempo do Grid
            If Len(GridItens.TextMatrix(GridItens.Row, iGrid_UMTempo_Col)) > 0 Then
                'Guardo o valor da UMTempo da Linha
                sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UMTempo_Col)
            Else
                'Senão coloco o Padrão UMTempo
                'sUnidadeMed = TAXA_CONSUMO_TEMPO_PADRAO
                Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)

            End If
            
            'Limpar as Unidades utilizadas anteriormente
            UMTempo.Clear

            For Each objUnidadeDeMedida In colSiglas
                UMTempo.AddItem objUnidadeDeMedida.sSigla
            Next

            UMTempo.AddItem ""

            'Tento selecionar na Combo a Unidade anterior
            If UMTempo.ListCount <> 0 Then

                For iIndice = 0 To UMTempo.ListCount - 1

                    If UMTempo.List(iIndice) = sUnidadeMed Then
                        UMTempo.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If
        
        Else
            objControl.Enabled = False
        
        End If
        
    'Grid Tipos de MO
    ElseIf objControl.Name = "CodigoTipoMO" Then
        
        If Len(sCodTipo) > 0 Then
        
            objControl.Enabled = False
            
        Else
        
            objControl.Enabled = True
        
        End If

    ElseIf objControl.Name = "DescricaoTipoMO" Then

        objControl.Enabled = False
                            
    ElseIf objControl.Name = "QuantidadeTipoMO" Then
        
        If Len(sCodTipo) > 0 Then
        
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
    
    ElseIf objControl.Name = "PercentualUso" Then
        
        If Len(sCodTipo) > 0 Then
            
            If Len(Trim(PercentualUso.Text)) = 0 Then
            
                PercentualUso.Text = 100
            
            End If
            
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 134133 To 134136
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162619)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridTiposDeMaodeObra(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Código")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("% Uso")

    'Controles que participam do Grid
    objGrid.colCampo.Add (CodigoTipoMO.Name)
    objGrid.colCampo.Add (DescricaoTipoMO.Name)
    objGrid.colCampo.Add (QuantidadeTipoMO.Name)
    objGrid.colCampo.Add (PercentualUso.Name)

    'Colunas do Grid
    iGrid_CodigoTipoMO_Col = 1
    iGrid_DescricaoTipoMO_Col = 2
    iGrid_QuantidadeTipoMO_Col = 3
    iGrid_PercentualUso_Col = 4

    objGrid.objGrid = GridTiposDeMaodeObra

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 11

    'Largura da primeira coluna
    GridTiposDeMaodeObra.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridTiposDeMaodeObra = SUCESSO

End Function

Private Sub GridTiposDeMaodeObra_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridTiposDeMaodeObra, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTiposDeMaodeObra, iAlterado)
    End If

End Sub

Private Sub GridTiposDeMaodeObra_GotFocus()
    
    Call Grid_Recebe_Foco(objGridTiposDeMaodeObra)

End Sub

Private Sub GridTiposDeMaodeObra_EnterCell()

    Call Grid_Entrada_Celula(objGridTiposDeMaodeObra, iAlterado)

End Sub

Private Sub GridTiposDeMaodeObra_LeaveCell()
    
    Call Saida_Celula(objGridTiposDeMaodeObra)

End Sub

Private Sub GridTiposDeMaodeObra_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridTiposDeMaodeObra, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTiposDeMaodeObra, iAlterado)
    End If

End Sub

Private Sub GridTiposDeMaodeObra_RowColChange()

    Call Grid_RowColChange(objGridTiposDeMaodeObra)

End Sub

Private Sub GridTiposDeMaodeObra_Scroll()

    Call Grid_Scroll(objGridTiposDeMaodeObra)

End Sub

Private Sub GridTiposDeMaodeObra_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridTiposDeMaodeObra)
        
End Sub

Private Sub GridTiposDeMaodeObra_LostFocus()

    Call Grid_Libera_Foco(objGridTiposDeMaodeObra)

End Sub


Private Sub CodigoTipoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoTipoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTiposDeMaodeObra)

End Sub

Private Sub CodigoTipoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTiposDeMaodeObra)

End Sub

Private Sub CodigoTipoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTiposDeMaodeObra.objControle = CodigoTipoMO
    lErro = Grid_Campo_Libera_Foco(objGridTiposDeMaodeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub DescricaoTipoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoTipoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTiposDeMaodeObra)

End Sub

Private Sub DescricaoTipoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTiposDeMaodeObra)

End Sub

Private Sub DescricaoTipoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTiposDeMaodeObra.objControle = DescricaoTipoMO
    lErro = Grid_Campo_Libera_Foco(objGridTiposDeMaodeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub QuantidadeTipoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeTipoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTiposDeMaodeObra)

End Sub

Private Sub QuantidadeTipoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTiposDeMaodeObra)

End Sub

Private Sub QuantidadeTipoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTiposDeMaodeObra.objControle = QuantidadeTipoMO
    lErro = Grid_Campo_Libera_Foco(objGridTiposDeMaodeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub PercentualUso_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentualUso_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTiposDeMaodeObra)

End Sub

Private Sub PercentualUso_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTiposDeMaodeObra)

End Sub

Private Sub PercentualUso_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTiposDeMaodeObra.objControle = PercentualUso
    lErro = Grid_Campo_Libera_Foco(objGridTiposDeMaodeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CodigoTipoMO(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodTipoMO As String
Dim iLinha As Integer
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_CodigoTipoMO

    Set objGridInt.objControle = CodigoTipoMO
                    
    'Se o campo foi preenchido
    If Len(CodigoTipoMO.Text) > 0 Then

        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridTiposDeMaodeObra.Row Then
                                                    
                If GridTiposDeMaodeObra.TextMatrix(iLinha, iGrid_CodigoTipoMO_Col) = CodigoTipoMO.Text Then
                    sCodTipoMO = CodigoTipoMO.Text
                    CodigoTipoMO.Text = ""
                    gError 134130
                    
                End If
                    
            End If
                           
        Next
        
        Set objTiposDeMaodeObra = New ClassTiposDeMaodeObra
        
        objTiposDeMaodeObra.iCodigo = StrParaInt(CodigoTipoMO.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 135029
    
        If lErro = SUCESSO Then

            GridTiposDeMaodeObra.TextMatrix(GridTiposDeMaodeObra.Row, iGrid_DescricaoTipoMO_Col) = objTiposDeMaodeObra.sDescricao
            GridTiposDeMaodeObra.TextMatrix(GridTiposDeMaodeObra.Row, iGrid_PercentualUso_Col) = Format(1, "Percent")
            
            'verifica se precisa preencher o grid com uma nova linha
            If GridTiposDeMaodeObra.Row - GridTiposDeMaodeObra.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
        
        Else
        
            CodigoTipoMO.Text = ""
            gError 137935
        
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134132

    Saida_Celula_CodigoTipoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigoTipoMO:

    Saida_Celula_CodigoTipoMO = gErr

    Select Case gErr

        Case 134130
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, sCodTipoMO, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 137935
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134128, 134129, 134131, 134132
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162620)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoTipoMO(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoTipoMO

    Set objGridInt.objControle = DescricaoTipoMO
                    
    'Se o campo foi preenchido
    If Len(DescricaoTipoMO.Text) > 0 Then

        'verifica se precisa preencher o grid com uma nova linha
        If GridTiposDeMaodeObra.Row - GridTiposDeMaodeObra.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134132

    Saida_Celula_DescricaoTipoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoTipoMO:

    Saida_Celula_DescricaoTipoMO = gErr

    Select Case gErr
        
        Case 134132
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162621)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeTipoMO(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantidadeTipoMO

    Set objGridInt.objControle = QuantidadeTipoMO
                    
    'Se o campo foi preenchido
    If Len(QuantidadeTipoMO.Text) > 0 Then

        'verifica se precisa preencher o grid com uma nova linha
        If GridTiposDeMaodeObra.Row - GridTiposDeMaodeObra.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134132

    Saida_Celula_QuantidadeTipoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeTipoMO:

    Saida_Celula_QuantidadeTipoMO = gErr

    Select Case gErr
        
        Case 134132
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162622)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualUso(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PercentualUso

    Set objGridInt.objControle = PercentualUso
                    
    'Se o campo foi preenchido
    If Len(PercentualUso.Text) > 0 Then

        'Critica o valor
        lErro = Porcentagem_Critica(PercentualUso.Text)
        If lErro <> SUCESSO Then gError 134337

        'verifica se precisa preencher o grid com uma nova linha
        If GridTiposDeMaodeObra.Row - GridTiposDeMaodeObra.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134132

    Saida_Celula_PercentualUso = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualUso:

    Saida_Celula_PercentualUso = gErr

    Select Case gErr
        
        Case 134132, 134337
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162623)

    End Select

    Exit Function

End Function

'----------------------------------------

Private Sub CustoHora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoHora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CustoHora_Validate

    'Veifica se CustoHora está preenchida
    If Len(Trim(CustoHora.Text)) <> 0 Then

       'Critica a CustoHora
       lErro = Valor_Positivo_Critica(CustoHora.Text)
       If lErro <> SUCESSO Then gError 137574

    End If

    Exit Sub

Erro_CustoHora_Validate:

    Cancel = True

    Select Case gErr

        Case 137574

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174977)

    End Select

    Exit Sub

End Sub

Private Sub CustoHora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CustoHora, iAlterado)
    
End Sub

Private Sub Prod_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Prod_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_Prod_Validate

    lErro = CF("Produto_Perde_Foco", Prod, DescProd)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 202440
    
    If lErro <> SUCESSO Then gError 202441

    Exit Sub

Erro_Prod_Validate:

    Cancel = True

    Select Case gErr

        Case 202440

        Case 202441
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202442)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", Prod.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134507

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    sFiltro = "Ativo = ? "
    
    colSelecao.Add PRODUTO_ATIVO
        
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProd, sFiltro)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 134507

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174528)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProd_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProd_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 134508

    Prod.PromptInclude = False
    Prod.Text = sProdutoMascarado
    Prod.PromptInclude = True
    
    Call Prod_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProd_evSelecao:

    Select Case gErr

        Case 134508
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174529)

    End Select

    Exit Sub

End Sub

