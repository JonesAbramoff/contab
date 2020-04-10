VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RequisicaoModeloOcx 
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   9450
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4320
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   750
      Width           =   9165
      Begin VB.Frame Frame3 
         Caption         =   "Cabeçalho"
         ClipControls    =   0   'False
         Height          =   4215
         Left            =   45
         TabIndex        =   2
         Top             =   75
         Width           =   9075
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2205
            Picture         =   "RequisicaoModeloOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Numeração Automática"
            Top             =   510
            Width           =   300
         End
         Begin VB.CheckBox Urgente 
            Caption         =   "Urgente"
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
            Left            =   5775
            TabIndex        =   18
            Top             =   2130
            Width           =   1005
         End
         Begin VB.ComboBox FilialCompra 
            Height          =   315
            Left            =   5760
            TabIndex        =   11
            Top             =   975
            Width           =   2520
         End
         Begin VB.TextBox Observacao 
            Height          =   315
            Left            =   1410
            MaxLength       =   255
            TabIndex        =   17
            Top             =   2190
            Width           =   3615
         End
         Begin VB.TextBox Descricao 
            Height          =   315
            Left            =   5760
            MaxLength       =   30
            TabIndex        =   7
            Top             =   420
            Width           =   3120
         End
         Begin VB.ComboBox TipoTributacao 
            Height          =   315
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1530
            Width           =   2520
         End
         Begin VB.Frame Frame6 
            Caption         =   "Local de Entrega"
            Height          =   1050
            Left            =   570
            TabIndex        =   19
            Top             =   2970
            Width           =   8070
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   750
               Index           =   0
               Left            =   4335
               TabIndex        =   23
               Top             =   195
               Width           =   3645
               Begin VB.ComboBox FilialEmpresa 
                  Height          =   315
                  Left            =   1110
                  TabIndex        =   25
                  Top             =   210
                  Width           =   2160
               End
               Begin VB.Label Label6 
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
                  Height          =   195
                  Left            =   600
                  TabIndex        =   24
                  Top             =   270
                  Width           =   465
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Tipo"
               Height          =   600
               Left            =   225
               TabIndex        =   20
               Top             =   240
               Width           =   3615
               Begin VB.OptionButton TipoDestino 
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
                  Height          =   255
                  Index           =   1
                  Left            =   2010
                  TabIndex        =   22
                  Top             =   255
                  Width           =   1335
               End
               Begin VB.OptionButton TipoDestino 
                  Caption         =   "Filial Empresa"
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
                  Index           =   0
                  Left            =   120
                  TabIndex        =   21
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   1515
               End
            End
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Height          =   675
               Index           =   1
               Left            =   4365
               TabIndex        =   26
               Top             =   225
               Visible         =   0   'False
               Width           =   3645
               Begin MSMask.MaskEdBox Fornecedor 
                  Height          =   300
                  Left            =   1110
                  TabIndex        =   28
                  Top             =   0
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   20
                  PromptChar      =   " "
               End
               Begin VB.ComboBox FilialFornecedor 
                  Height          =   315
                  Left            =   1125
                  TabIndex        =   30
                  Top             =   360
                  Width           =   2160
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
                  Height          =   195
                  Left            =   0
                  MousePointer    =   14  'Arrow and Question
                  TabIndex        =   27
                  Top             =   60
                  Width           =   1035
               End
               Begin VB.Label Label21 
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
                  Height          =   195
                  Left            =   570
                  TabIndex        =   29
                  Top             =   405
                  Width           =   465
               End
            End
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1395
            TabIndex        =   4
            Top             =   480
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   300
            Left            =   1425
            TabIndex        =   9
            Top             =   1005
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   315
            Left            =   2190
            TabIndex        =   13
            Top             =   1545
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label CodigoLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   645
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   3
            Top             =   540
            Width           =   690
         End
         Begin VB.Label LabelRequisitante 
            AutoSize        =   -1  'True
            Caption         =   "Requisitante:"
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
            Left            =   195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   8
            Top             =   1065
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Filial Compra:"
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
            Left            =   4455
            TabIndex        =   10
            Top             =   1020
            Width           =   1155
         End
         Begin VB.Label ObservacaoLabel 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   225
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   16
            Top             =   2220
            Width           =   1095
         End
         Begin VB.Label CclPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo/Lucro:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   12
            Top             =   1635
            Width           =   2010
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   4680
            TabIndex        =   6
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tributação:"
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
            Left            =   3915
            TabIndex        =   14
            Top             =   1590
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4290
      Index           =   2
      Left            =   105
      TabIndex        =   31
      Top             =   750
      Visible         =   0   'False
      Width           =   9150
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
         Height          =   285
         Left            =   135
         TabIndex        =   46
         Top             =   3855
         Width           =   1005
      End
      Begin VB.CommandButton BotaoAlmoxarifados 
         Caption         =   "Almoxarifados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5046
         TabIndex        =   49
         Top             =   3855
         Width           =   1485
      End
      Begin VB.CommandButton BotaoFiliaisFornProd 
         Caption         =   "Fornecedores do Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6690
         TabIndex        =   50
         Top             =   3855
         Width           =   2370
      End
      Begin VB.CommandButton BotaoCcl 
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
         Height          =   285
         Left            =   1312
         TabIndex        =   47
         Top             =   3855
         Width           =   1710
      End
      Begin VB.CommandButton BotaoPlanoConta 
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
         Height          =   285
         Left            =   3179
         TabIndex        =   48
         Top             =   3855
         Width           =   1710
      End
      Begin VB.Frame Frame4 
         Caption         =   "Itens"
         Height          =   3390
         Left            =   120
         TabIndex        =   32
         Top             =   210
         Width           =   8955
         Begin VB.ComboBox TipoTribItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   2670
            Width           =   2520
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   240
            Left            =   7155
            TabIndex        =   40
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
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
         Begin MSMask.MaskEdBox CentroCusto 
            Height          =   225
            Left            =   6405
            TabIndex        =   39
            Top             =   375
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   5040
            TabIndex        =   38
            Top             =   420
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   3990
            TabIndex        =   37
            Top             =   375
            Width           =   990
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecGrid 
            Height          =   225
            Left            =   1785
            TabIndex        =   42
            Top             =   2835
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox Exclusivo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "RequisicaoModeloOcx.ctx":00EA
            Left            =   5505
            List            =   "RequisicaoModeloOcx.ctx":00F4
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2685
            Width           =   1305
         End
         Begin VB.ComboBox FilialFornecGrid 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3735
            TabIndex        =   43
            Top             =   2730
            Width           =   1770
         End
         Begin VB.TextBox ObservacaoGrid 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6405
            MaxLength       =   255
            TabIndex        =   45
            Top             =   2715
            Width           =   2355
         End
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2865
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   345
            Width           =   1065
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   35
            Top             =   420
            Width           =   1455
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   285
            TabIndex        =   34
            Top             =   390
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2595
            Left            =   180
            TabIndex        =   33
            Top             =   375
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   4577
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            Redraw          =   0   'False
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7185
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "RequisicaoModeloOcx.ctx":0111
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RequisicaoModeloOcx.ctx":028F
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RequisicaoModeloOcx.ctx":07C1
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RequisicaoModeloOcx.ctx":094B
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4725
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8334
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
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
Attribute VB_Name = "RequisicaoModeloOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variaveis Globais
Dim gColItemReqModelo As Collection
Dim giTipoTributacao As Integer

'EVENTOS DOS BROWSERS
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoRequisitante As AdmEvento
Attribute objEventoRequisitante.VB_VarHelpID = -1
Private WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Private WithEvents objEventoObservacao As AdmEvento
Attribute objEventoObservacao.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoBotaoCcl As AdmEvento
Attribute objEventoBotaoCcl.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxarifados As AdmEvento
Attribute objEventoAlmoxarifados.VB_VarHelpID = -1
Private WithEvents objEventoFiliaisFornProduto As AdmEvento
Attribute objEventoFiliaisFornProduto.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer

'GridItens
Dim objGridItens As AdmGrid
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_CCL_Col As Integer
Dim iGrid_ContaContabil_Col As Integer
Dim iGrid_TipoTributacao_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialFornecedor_Col As Integer
Dim iGrid_Exclusivo_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iFrameAtual As Integer

Function Trata_Parametros(Optional objRequisicaoModelo As ClassRequisicaoModelo) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma Requisição foi passada por parâmetro
    If Not (objRequisicaoModelo Is Nothing) Then

        'Se o número interno estiver preenchido
        If objRequisicaoModelo.lNumIntDoc > 0 Then

            'Lê a Requisição Modelo a partir de seu número interno
            lErro = CF("RequisicaoModelo_Le", objRequisicaoModelo)
            If lErro = SUCESSO Then

                'Traz os dados da Requisição para a tela
                lErro = Traz_RequisicaoModelo_Tela(objRequisicaoModelo)
                If lErro <> SUCESSO Then Error 61579
            
            'Se não encontrou a Requisição
            Else
                
                'Exibe apenas o código passado por objRequisicaoModelo
                Codigo.Text = CStr(objRequisicaoModelo.lCodigo)
            
            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 61579

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174010)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("RequisicaoModelo_Codigo_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 61816

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 61816
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174011)
    
    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim sMascaraCclPadrao As String
Dim bCancel As Boolean

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    'Inicializa os ObjEventos
    Set objEventoCodigo = New AdmEvento
    Set objEventoRequisitante = New AdmEvento
    Set objEventoCclPadrao = New AdmEvento
    Set objEventoObservacao = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoBotaoCcl = New AdmEvento
    Set objEventoAlmoxarifados = New AdmEvento
    Set objEventoFiliaisFornProduto = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    
    'Inicializa coleção de Itens de Requisição
    Set gColItemReqModelo = New Collection
    
    'Atualiza a variável global para controle de frames e seta um tipo Padrao
    iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True

    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 61517

    'Preenche a Combo FilialCompra com as filiais Empresas
    lErro = Carrega_ComboFiliais(colCodigoDescricao)
    If lErro <> SUCESSO Then Error 61518

    'Carrega Tipos de Tributação
    lErro = Carrega_TipoTributacao()
    If lErro <> SUCESSO Then gError 66605

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 61519

    'Inicializa Máscara de CentroCusto e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then Error 61520

    Ccl.Mask = sMascaraCclPadrao
    CentroCusto.Mask = sMascaraCclPadrao

    'Inicializa mascara de ContaContabil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then Error 61521

    'Inicializa o GridItens
    Set objGridItens = New AdmGrid

    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then Error 61522

    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then Error 61706

    'Coloca FilialCompra Default na tela
    If objConfiguraCOM.iFilialCompra > 0 Then
        FilialCompra.Text = objConfiguraCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If
    Call FilialCompra_Validate(bCancel)
    
    
    
    'Coloca FiliaEmpresa Default na Tela
    FilialEmpresa.Text = giFilialEmpresa
    Call FilialEmpresa_Validate(bCancel)
    
    FilialEmpresa.ListIndex = 0
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_LIGHT Then
        
        FilialCompra.left = POSICAO_FORA_TELA
        FilialCompra.TabStop = False
        Label3.left = POSICAO_FORA_TELA
        Label3.Visible = False
        FilialEmpresa.left = POSICAO_FORA_TELA
        FilialEmpresa.TabStop = False
        Label6.left = POSICAO_FORA_TELA
        Label6.Visible = False
        
    End If
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 61517, 61518, 61519, 61520, 61521, 61522, 61706, 66605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174012)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Carrega_ComboFiliais(colCodigoDescricao As AdmColCodigoNome) As Long
'Carrega as Combos (FilialEmpresa e FilialCompra com as Filiais Empresa passada na colecao

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Carrega_ComboFiliais

    'Preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao

        If objCodigoNome.iCodigo <> 0 Then

            FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo

            FilialCompra.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialCompra.ItemData(FilialCompra.NewIndex) = objCodigoNome.iCodigo

        End If

    Next

    Carrega_ComboFiliais = SUCESSO

    Exit Function

Erro_Carrega_ComboFiliais:

    Carrega_ComboFiliais = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174013)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil")
    objGridInt.colColuna.Add ("Tipo de Tributação")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusividade")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)
    objGridInt.colCampo.Add (TipoTribItem.Name)
    objGridInt.colCampo.Add (FornecGrid.Name)
    objGridInt.colCampo.Add (FilialFornecGrid.Name)
    objGridInt.colCampo.Add (Exclusivo.Name)
    objGridInt.colCampo.Add (ObservacaoGrid.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UM_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_Almoxarifado_Col = 5
    iGrid_CCL_Col = 6
    iGrid_ContaContabil_Col = 7
    iGrid_TipoTributacao_Col = 8
    iGrid_Fornecedor_Col = 9
    iGrid_FilialFornecedor_Col = 10
    iGrid_Exclusivo_Col = 11
    iGrid_Observacao_Col = 12

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridItens.Row = 0 Then gError 66600

    If GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = "" Then gError 66601

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 66602

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 66600
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 66601
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 66602

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174014)

    End Select

    Exit Sub

End Sub

Private Sub DescProduto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta <> "" Then
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 66599
            
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
        
        GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 66599
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174015)

    End Select

    Exit Sub

End Sub

Function Carrega_TipoTributacao() As Long
'Carrega Tipos de Tributação

Dim lErro As Long
Dim colTributacao As New AdmColCodigoNome
Dim iIndice As Integer
Dim iTipoTrib As Integer

On Error GoTo Erro_Carrega_TipoTributacao

    'Lê os Tipos de Tributação associadas a Compras
    lErro = CF("TiposTributacaoCompras_Le", colTributacao)
    If lErro <> SUCESSO Then gError 66603
        
    'Lê o Tipo de Tributação Padrão
    lErro = CF("TipoTributacaoPadrao_Le", iTipoTrib)
    If lErro <> SUCESSO And lErro <> 66597 Then gError 66604
    
    'Carrega Tipos de Tributação
    For iIndice = 1 To colTributacao.Count
        TipoTributacao.AddItem colTributacao(iIndice).iCodigo & SEPARADOR & colTributacao(iIndice).sNome
        TipoTribItem.AddItem colTributacao(iIndice).iCodigo & SEPARADOR & colTributacao(iIndice).sNome
    Next
    
    'Seleciona Tipo de Tributação default
    For iIndice = 0 To TipoTributacao.ListCount - 1
        If Codigo_Extrai(TipoTributacao.List(iIndice)) = iTipoTrib Then
            TipoTributacao.ListIndex = iIndice
            TipoTribItem.ListIndex = iIndice
            Exit For
        End If
    Next
    
    giTipoTributacao = iTipoTrib
    
    Carrega_TipoTributacao = SUCESSO
    
    Exit Function
    
Erro_Carrega_TipoTributacao:

    Carrega_TipoTributacao = gErr
    
    Select Case gErr
        
        Case 66603, 66604
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174016)
        
    End Select
    
    Exit Function
    
End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)
'Rotina que habilita a entrada na celula

Dim lErro As Long
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim objProduto As New ClassProduto
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim iIndice As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor
Dim iCodigo As Integer
Dim sTipoTrib As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError (61523)

    'Passa o produto controle
    If objControl.Name = Produto.Name Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
           objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If

    ElseIf objControl.Name = UM.Name Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objControl.Enabled = True

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError (61524)

            If lErro = 28030 Then gError (61525)

            objClasseUM.iClasse = objProduto.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError (61526)

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col)

            'Limpar as Unidades utilizadas anteriormente
            UM.Clear

            For Each objUnidadeDeMedida In colSiglas
                UM.AddItem objUnidadeDeMedida.sSigla
            Next

            'Tento selecionar na Combo a Unidade anterior
            If UM.ListCount <> 0 Then

                For iIndice = 0 To UM.ListCount - 1
                    If UM.List(iIndice) = sUnidadeMed Then
                        UM.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = Almoxarifado.Name Or objControl.Name = CentroCusto.Name Then

            'Verifica se o detino é a empresa
            If iFrameTipoDestinoAtual <> TIPO_DESTINO_EMPRESA Or Len(Trim(FilialEmpresa.Text)) = 0 Then
                
                Almoxarifado.Text = ""
                CentroCusto.PromptInclude = False
                CentroCusto.Text = ""
                CentroCusto.PromptInclude = True
                
                objControl.Enabled = False
            Else

                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If

            End If

    ElseIf objControl.Name = Exclusivo.Name Then
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col))) > 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
    
    ElseIf objControl.Name = Quantidade.Name Or objControl.Name = ContaContabil.Name Or objControl.Name = FornecGrid.Name Or objControl.Name = ObservacaoGrid.Name Or objControl.Name = DescProduto.Name Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = TipoTribItem.Name Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objControl.Enabled = True
            
            'Guardo o valor do Tipo de tributação
            sTipoTrib = GridItens.TextMatrix(GridItens.Row, iGrid_TipoTributacao_Col)

            'Limpar os Tipos de tributação
            TipoTribItem.Clear

            For iIndice = 0 To TipoTributacao.ListCount - 1
                TipoTribItem.AddItem TipoTributacao.List(iIndice)
            Next

            'Tento selecionar na Combo o Tipo anterior
            If TipoTribItem.ListCount <> 0 Then

                For iIndice = 0 To TipoTribItem.ListCount - 1
                    If TipoTribItem.List(iIndice) = sTipoTrib Then
                        TipoTribItem.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If
        
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = FilialFornecGrid.Name Then
    
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
            objControl.Enabled = False
        Else
            
            objControl.Enabled = True
            
            'Se o Fornecedor não está preenchido
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col))) = 0 Then
                
                'Desabilita combo de Filiais
                objControl.Enabled = False
                
            Else
                
                objFornecedor.sNomeReduzido = GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col)
                
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError (65594)
                If lErro = 6681 Then gError (65595)
                
                lErro = CF("FornecedorProdutoFF_Le_FilialForn", sProdutoFormatado, objFornecedor.lCodigo, Codigo_Extrai(FilialCompra.Text), colCodigoNome)
                If lErro <> SUCESSO Then gError (65596)
                
                If colCodigoNome.Count = 0 Then gError (65597)
                    
                If Len(Trim(FilialFornecGrid.Text)) = 0 Then
                    iCodigo = colCodigoNome.Item(1).iCodigo
                Else
                    iCodigo = Codigo_Extrai(FilialFornecGrid.Text)
                End If

                FilialFornecGrid.Clear
                    
                Call CF("Filial_Preenche", FilialFornecGrid, colCodigoNome)
                Call CF("Filial_Seleciona", FilialFornecGrid, iCodigo)
            
            End If
            
        End If
        
    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 61523, 61524, 61526, 65594, 65596

        Case 61525
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 65595
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case 65597
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR", gErr, objFornecedor.sNomeReduzido, sProdutoFormatado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174017)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'Produto
            Case iGrid_Produto_Col
                lErro = Saida_Celula_Produto(objGridInt)
                If lErro <> SUCESSO Then gError 61787
            
            Case iGrid_Descricao_Col
                lErro = Saida_Celula_Descricao(objGridInt)
                If lErro <> SUCESSO Then gError 86177
            
            'Unidade de Medida
            Case iGrid_UM_Col
                lErro = Saida_Celula_UnidadeMed(objGridInt)
                If lErro <> SUCESSO Then gError 61788
            
            'Quantidade
            Case iGrid_Quantidade_Col
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 61789
            
            'Almoxarifado
            Case iGrid_Almoxarifado_Col
                lErro = Saida_Celula_Almoxarifado(objGridInt)
                If lErro <> SUCESSO Then gError 61790
            
            'Ccl
            Case iGrid_CCL_Col
                lErro = Saida_Celula_Ccl(objGridInt)
                If lErro <> SUCESSO Then gError 61791
                    
            'ContaContabil
            Case iGrid_ContaContabil_Col
                lErro = Saida_Celula_ContaContabil(objGridInt)
                If lErro <> SUCESSO Then gError 61792

            'Fornecedor
            Case iGrid_Fornecedor_Col
                lErro = Saida_Celula_Fornecedor(objGridInt)
                If lErro <> SUCESSO Then gError 61793
                
            'Filial Fornecedor
            Case iGrid_FilialFornecedor_Col
                lErro = Saida_Celula_FilialForn(objGridInt)
                If lErro <> SUCESSO Then gError 61794
                
            'Exclusivo
            Case iGrid_Exclusivo_Col
                lErro = Saida_Celula_Exclusivo(objGridInt)
                If lErro <> SUCESSO Then gError 61795
            
            'Observação
            Case iGrid_Observacao_Col
                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 61796
        
            'Tipo de Tributação
            Case iGrid_TipoTributacao_Col
                lErro = Saida_Celula_TipoTributacao(objGridInt)
                If lErro <> SUCESSO Then gError 65501
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 61797

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 61787 To 61797, 65501, 86177
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174018)

    End Select

    Exit Function

End Function

'SISTEMA DE SETAS
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objRequisicaoModelo As New ClassRequisicaoModelo
Dim sNomeRed As String

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ReqModRequisitante"
    
    ' Move todos os dados Presentes da Tela para objRequisaoModelo
    lErro = Move_Tela_Memoria(objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61527

    sNomeRed = Requisitante.Text

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objRequisicaoModelo.lCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objRequisicaoModelo.sDescricao, STRING_BUFFER_MAX_TEXTO, "Descricao"
    colCampoValor.Add "Ccl", objRequisicaoModelo.sCcl, STRING_BUFFER_MAX_TEXTO, "Ccl"
    colCampoValor.Add "Requisitante", objRequisicaoModelo.lRequisitante, 0, "Requisitante"
    colCampoValor.Add "FilialCompra", objRequisicaoModelo.iFilialCompra, 0, "FilialCompra"
    colCampoValor.Add "Observacao", objRequisicaoModelo.sObservacao, STRING_BUFFER_MAX_TEXTO, "Observacao"
    colCampoValor.Add "TipoDestino", objRequisicaoModelo.iTipoDestido, 0, "TipoDestino"
    colCampoValor.Add "FornCliDestino", objRequisicaoModelo.lFornCliDestino, 0, "FornCliDestino"
    colCampoValor.Add "FilialDestino", objRequisicaoModelo.iFilialDestino, 0, "FilialDestino"
    colCampoValor.Add "Urgente", objRequisicaoModelo.iUrgente, 0, "Urgente"
'    colCampoValor.Add "TipoTributacao", objRequisicaoModelo.iTipoTributacao, 0, "TipoTributacao"
    colCampoValor.Add "NomeReduzido", sNomeRed, STRING_BUFFER_MAX_TEXTO, "NomeReduzido"
    colCampoValor.Add "lObservacao", objRequisicaoModelo.lObservacao, 0, "lObservacao"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 61527

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174019)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objRequisicaoModelo As New ClassRequisicaoModelo

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objRequisicaoModelo.lCodigo = colCampoValor.Item("Codigo").vValor
    objRequisicaoModelo.sDescricao = colCampoValor.Item("Descricao").vValor
    objRequisicaoModelo.sCcl = colCampoValor.Item("Ccl").vValor
    objRequisicaoModelo.lRequisitante = colCampoValor.Item("Requisitante").vValor
    objRequisicaoModelo.iFilialCompra = colCampoValor.Item("FilialCompra").vValor
    objRequisicaoModelo.sObservacao = colCampoValor.Item("Observacao").vValor
    objRequisicaoModelo.iTipoDestido = colCampoValor.Item("TipoDestino").vValor
    objRequisicaoModelo.lFornCliDestino = colCampoValor.Item("FornCliDestino").vValor
    objRequisicaoModelo.iFilialDestino = colCampoValor.Item("FilialDestino").vValor
    objRequisicaoModelo.iUrgente = colCampoValor.Item("Urgente").vValor
'    objRequisicaoModelo.iTipoTributacao = colCampoValor.Item("TipoTributacao").vValor
    objRequisicaoModelo.iFilialEmpresa = giFilialEmpresa
    objRequisicaoModelo.lObservacao = colCampoValor.Item("lObservacao").vValor

    'Traz os dados da Requisição Modelo para tela
    lErro = Traz_RequisicaoModelo_Tela(objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61528

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 61528

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174020)

    End Select

    Exit Sub

End Sub

Private Sub TipoTributacao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Function Move_Tela_Memoria(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Move os dados da tela para o objRequisicaoModelo

Dim objRequisitante As New ClassRequisitante
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    'Move o codigo e a descricao
    objRequisicaoModelo.lCodigo = StrParaLong(Codigo.Text)
    objRequisicaoModelo.sDescricao = Descricao.Text

    'Move Urgente
    objRequisicaoModelo.iUrgente = Urgente.Value

    If Len(Trim(Requisitante.Text)) > 0 Then
    
        'Move o requisitante
        objRequisitante.sNomeReduzido = Requisitante.Text
    
        'Lê os dados do Requisitante a partir do seu Nome Reduzido
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then Error 61717
    
        'Se não encontrou o Requisitante, erro
        If lErro = 51152 Then Error 61718
    
    End If
    
    objRequisicaoModelo.lRequisitante = objRequisitante.lCodigo
    objRequisicaoModelo.iTipoTributacao = Codigo_Extrai(TipoTributacao.Text)
    
    'Move CCL
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 61719

    objRequisicaoModelo.sCcl = sCclFormatada

    'Move a Filial Compra
    objRequisicaoModelo.iFilialCompra = Codigo_Extrai(FilialCompra.Text)

    'Move a Observacao
    objRequisicaoModelo.sObservacao = Observacao.Text

    'Move a FilialEmpresa
    objRequisicaoModelo.iFilialEmpresa = giFilialEmpresa

    'Move o Frame local de entrega
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then

        objRequisicaoModelo.iTipoDestido = TIPO_DESTINO_EMPRESA
        objRequisicaoModelo.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)

    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then

        objRequisicaoModelo.iTipoDestido = TIPO_DESTINO_FORNECEDOR
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 61805
        If lErro = 6681 Then Error 61806
                
        objRequisicaoModelo.lFornCliDestino = objFornecedor.lCodigo
        objRequisicaoModelo.iFilialDestino = Codigo_Extrai(FilialFornecedor.Text)

    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err

        Case 61717, 61719, 61805
        
        Case 61718
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO1", Err, objRequisitante.sNomeReduzido)
        
        Case 61806
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174021)

    End Select

    Exit Function

End Function

Function Move_GridItens_Memoria(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Move itens do Grid para objRequisicaoModelo

Dim lErro As Long
Dim iIndice As Integer, iCount As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCcl As String, sCclFormatada As String, iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemReqModelo As ClassItemReqModelo
Dim sContaContabil As String
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim objFornecedor As New ClassFornecedor
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemReqModelo = New ClassItemReqModelo
        
        sProduto = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError (61657)

        objItemReqModelo.sProduto = sProdutoFormatado
        objItemReqModelo.sDescProduto = GridItens.TextMatrix(iIndice, iGrid_Descricao_Col)
        objItemReqModelo.sUM = GridItens.TextMatrix(iIndice, iGrid_UM_Col)
                
        objItemReqModelo.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) > 0 Then
        
            objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
    
            'Lê dados do almoxarifado a partir do Nome Reduzido
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25060 Then gError (61658)
    
            'Se não econtrou o almoxarifado, erro
            If lErro = 25060 Then gError (61659)
    
            objItemReqModelo.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        End If
        
        sCcl = GridItens.TextMatrix(iIndice, iGrid_CCL_Col)

        If Len(Trim(sCcl)) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError (61660)
        Else
            sCclFormatada = ""
        End If

        objItemReqModelo.sCcl = sCclFormatada

        sContaContabil = GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col)
        
        If Len(Trim(sContaContabil)) > 0 Then
            
            'Formata ContaContábil para BD
            lErro = CF("Conta_Formata", sContaContabil, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError (61661)
        
        Else
            sContaFormatada = ""
        End If
        
        objItemReqModelo.sContaContabil = sContaFormatada
        
        objItemReqModelo.iTipoTributacao = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col))
        
        'Move o Código do Fornecedor
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col))) > 0 Then
            
            objFornecedor.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col)
            
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError (61662)
            If lErro = 6681 Then gError (61663)
            objItemReqModelo.lFornecedor = objFornecedor.lCodigo
            
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col))) > 0 Then
                objItemReqModelo.iFilial = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col))
            End If
            
            objFornecedorProdutoFF.lFornecedor = objItemReqModelo.lFornecedor
            objFornecedorProdutoFF.iFilialForn = objItemReqModelo.iFilial
            objFornecedorProdutoFF.sProduto = objItemReqModelo.sProduto
            objFornecedorProdutoFF.iFilialEmpresa = objRequisicaoModelo.iFilialCompra
            
            lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
            If lErro <> SUCESSO And lErro <> 54217 Then gError 86097
            If lErro <> SUCESSO Then gError 86098
            
        End If
        
        If GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) = "Exclusivo" Then
            objItemReqModelo.iExclusivo = 1
        Else
            objItemReqModelo.iExclusivo = 0
        End If
        
        objItemReqModelo.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)
        objItemReqModelo.iTipoTributacao = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col))
        
        objItemReqModelo.lNumIntDoc = gColItemReqModelo.Item(iIndice)
        objRequisicaoModelo.colItensReqModelo.Add objItemReqModelo
    
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 61657, 61658, 61660, 61661, 61662, 86097
        
        Case 61659
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 61663
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
                            
        Case 86098
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174022)
    
    End Select

    Exit Function

End Function

Function Traz_RequisicaoModelo_Tela(objRequisicaoModelo As ClassRequisicaoModelo) As Long

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim objObservacao As New ClassObservacao
Dim bCancel As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Traz_RequisicaoModelo_Tela

    lErro = CF("RequisicaoModelo_Le_Codigo", objRequisicaoModelo)
    If lErro <> SUCESSO And lErro <> 61508 Then Error 61701
    
    'Limpa tela Requisicao Modelo
    Call Limpa_Tela_RequisicaoModelo

    'Coloca os dados na tela
    Codigo.PromptInclude = False
    Codigo.Text = objRequisicaoModelo.lCodigo
    Codigo.PromptInclude = True
    
    Descricao.Text = objRequisicaoModelo.sDescricao
    Urgente.Value = objRequisicaoModelo.iUrgente

    'Verifica se Observacao esta preenchido
    If objRequisicaoModelo.lObservacao <> 0 Then

        objObservacao.lNumInt = objRequisicaoModelo.lObservacao
        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then Error 61701
        If lErro = 53827 Then Error 61702
        Observacao.Text = objObservacao.sObservacao

    End If
    
    If objRequisicaoModelo.lRequisitante <> 0 Then
    
        objRequisitante.lCodigo = objRequisicaoModelo.lRequisitante
    
        'Le o requisitante para colocar o NomeReduzido na tela
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then Error 61535
        If lErro = 49084 Then Error 61536
    
        Requisitante.Text = objRequisitante.sNomeReduzido
    
    End If
    
    'Tipo de Tributação
    For iIndice = 0 To TipoTributacao.ListCount - 1
        If Codigo_Extrai(TipoTributacao.List(iIndice)) = objRequisicaoModelo.iTipoTributacao Then
            TipoTributacao.ListIndex = iIndice
            Exit For
        End If
    Next

    If Len(Trim(objRequisicaoModelo.sCcl)) > 0 Then
    
        'Preenche a CCL
        sCclMascarado = String(STRING_CCL, 0)
    
        lErro = Mascara_MascararCcl(objRequisicaoModelo.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 61537
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
    
    End If
    
    FilialCompra.Text = objRequisicaoModelo.iFilialCompra
    Call FilialCompra_Validate(bCancel)

    'Preenche TipoDestino e suas Caracteristicas
    lErro = Preenche_TipoDestino(objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61539

    'Lê os itens da Requisicao Modelo
    lErro = CF("ItensReqModelo_Le", objRequisicaoModelo)
    If lErro <> SUCESSO And lErro <> 61533 Then Error 61529

    'Se não encontrou itens, erro
    If objRequisicaoModelo.colItensReqModelo.Count = 0 Then Error 61534

    'Preenche o grid com os Itens da requição modelo
    lErro = Preenche_GridItens(objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61540

    iAlterado = 0

    Traz_RequisicaoModelo_Tela = SUCESSO

    Exit Function

Erro_Traz_RequisicaoModelo_Tela:

    Traz_RequisicaoModelo_Tela = Err

    Select Case Err

        Case 61534
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_MODELO_AUSENCIA_ITENS", Err, objRequisicaoModelo.lCodigo)

        Case 61529, 61535, 61537, 61539, 61540, 61701

        Case 61536
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", Err, objRequisitante.lCodigo)

        Case 61702
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", Err, objObservacao.lNumInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174023)

    End Select

    Exit Function

End Function

Function Preenche_TipoDestino(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Preenche o Tipo destino

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

On Error GoTo Erro_Preenche_TipoDestino

    TipoDestino.Item(objRequisicaoModelo.iTipoDestido).Value = True

    Select Case objRequisicaoModelo.iTipoDestido

        Case TIPO_DESTINO_EMPRESA

            FilialEmpresa.Text = objRequisicaoModelo.iFilialDestino
            Call FilialEmpresa_Validate(bCancel)

        Case TIPO_DESTINO_FORNECEDOR
            
            objFornecedor.lCodigo = objRequisicaoModelo.lFornCliDestino

            'Lê o fornecedor, seu nome reduzido
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then Error 61541
            
            'Se não encontrou o Forncedor, Erro
            If lErro = 12729 Then Error 61799
            
            Fornecedor.Text = objFornecedor.sNomeReduzido

            FilialFornecedor.Text = objRequisicaoModelo.iFilialDestino
            Call FilialFornecedor_Validate(bCancel)
            
        Case Else
            Error 61543

    End Select

    Preenche_TipoDestino = SUCESSO

    Exit Function

Erro_Preenche_TipoDestino:

    Preenche_TipoDestino = Err

    Select Case Err

        Case 61541, 61543

        Case 61799
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174024)

    End Select

    Exit Function

End Function

Function Preenche_GridItens(objRequicaoModelo As ClassRequisicaoModelo) As Long

Dim lErro As Long
Dim objItemReqModelo As ClassItemReqModelo
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim objObservacao As New ClassObservacao
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iIndice As Integer
Dim iItem As Integer
Dim sProdutoMascarado As String
Dim sCclMascarado As String
Dim sContaEnxuta As String
Dim iCont As Integer

On Error GoTo Erro_Preenche_GridItens
    
    Set gColItemReqModelo = New Collection
    
    'Preenche GridItens
    For Each objItemReqModelo In objRequicaoModelo.colItensReqModelo

        iIndice = iIndice + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Colococa o Produto Mascarado no Grid
        lErro = Mascara_RetornaProdutoEnxuto(objItemReqModelo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 61544

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        
        GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemReqModelo.sDescProduto
        GridItens.TextMatrix(iIndice, iGrid_UM_Col) = objItemReqModelo.sUM
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Format(objItemReqModelo.dQuantidade, "Standard")
        
        If objRequicaoModelo.iTipoDestido = TIPO_DESTINO_EMPRESA Then
        
            If objItemReqModelo.iAlmoxarifado <> 0 Then
                
                'Lê o Almoxarifado e coloca seu nome Reduzido no Grid
                objAlmoxarifado.iCodigo = objItemReqModelo.iAlmoxarifado
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then Error 61548
                If lErro = 25056 Then Error 61549
            
                GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            
            End If
                        
            'Coloca o Ccl mascarado no Grid
            If Len(Trim(objItemReqModelo.sCcl)) > 0 Then
                lErro = Mascara_RetornaCclEnxuta(objItemReqModelo.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 61550
            Else
                sCclMascarado = ""
            End If
            
            'Preenche o campo ccl
            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True
            
            'Preenche o Ccl no grid
            GridItens.TextMatrix(iIndice, iGrid_CCL_Col) = Ccl.Text
                    
        End If
        
        'Coloca Conta Contábil no Grid
        If Len(Trim(objItemReqModelo.sContaContabil)) > 0 Then

            lErro = Mascara_RetornaContaEnxuta(objItemReqModelo.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then Error 61545

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True
            GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col) = ContaContabil.Text

        End If
        
        'Tipo de Tributação
        For iCont = 0 To TipoTribItem.ListCount - 1
            If Codigo_Extrai(TipoTribItem.List(iCont)) = objItemReqModelo.iTipoTributacao Then
                GridItens.TextMatrix(iIndice, iGrid_TipoTributacao_Col) = TipoTribItem.List(iCont)
                Exit For
            End If
        Next
        
        'Coloca Nome Reduzido do Fornecedor no Grid
        If objItemReqModelo.lFornecedor > 0 Then

            objFornecedor.lCodigo = objItemReqModelo.lFornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then Error 61546
            If lErro = 12729 Then Error 61547

            GridItens.TextMatrix(iIndice, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido

        End If

        'Coloca Filial do Fornecedor no Grid
        If objItemReqModelo.iFilial > 0 Then
            
            objFilialFornecedor.iCodFilial = objItemReqModelo.iFilial
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then Error 61704
            If lErro = 12929 Then Error 61705
            
            GridItens.TextMatrix(iIndice, iGrid_FilialFornecedor_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        End If
        
        If objItemReqModelo.lFornecedor <> 0 Then
        
            'Preenche Exclusivo
            For iItem = 0 To Exclusivo.ListCount - 1
                If objItemReqModelo.iExclusivo = Exclusivo.ItemData(iItem) Then
                    GridItens.TextMatrix(iIndice, iGrid_Exclusivo_Col) = Exclusivo.List(iItem)
                    Exit For
                End If
            Next
        
        End If
        
        'Se possui observação
        If objItemReqModelo.lObservacao <> 0 Then

            objObservacao.lNumInt = objItemReqModelo.lObservacao

            'Lê a observação a partir do número interno
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then Error 61801
            If lErro <> SUCESSO Then Error 61802

            GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao

        End If
          
        gColItemReqModelo.Add objItemReqModelo.lNumIntDoc
    
    Next
    
    lErro = Grid_Refresh_Checkbox(objGridItens)
    If lErro <> SUCESSO Then Error 61703

    objGridItens.iLinhasExistentes = gColItemReqModelo.Count

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = Err

    Select Case Err

        Case 61544, 61545, 61546, 61548, 61550, 61703, 61704, 61801
        
        Case 61549
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)
                    
        Case 61547
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)
        
        Case 61705
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
            
        Case 61802
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", Err, objObservacao.lNumInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174025)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRequisicaoModelo As New ClassRequisicaoModelo
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o código não estiver preenchido, Erro
    If Len(Trim(Codigo.Text)) = 0 Then Error 61683
     
    objRequisicaoModelo.iFilialEmpresa = giFilialEmpresa
    objRequisicaoModelo.lCodigo = StrParaLong(Codigo.Text)
    
    'Lê a Requisição Modelo a partir do código Passado em objRequisicaoModelo
    lErro = CF("RequisicaoModelo_Le_Codigo", objRequisicaoModelo)
    If lErro <> SUCESSO And lErro <> 61508 Then Error 61684
    If lErro = 61508 Then Error 61685
    
    'Envia aviso perguntando se realmente deseja excluir a Requisição Modelo
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_REQUISICAOMODELO", objRequisicaoModelo.lCodigo)

    'Se a resposta for positiva
    If vbMsgRes = vbYes Then
    
        'Exclui a Requisição Modelo
        lErro = CF("RequisicaoModelo_Exclui", objRequisicaoModelo)
        If lErro <> SUCESSO Then Error 61686
        
        'Limpa a Tela
        Call Limpa_Tela_RequisicaoModelo
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        iAlterado = 0
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case Err
        
        Case 61683
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
        
        Case 61684, 61686
            
        Case 61685
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOMODELO_NAO_CADASTRADA", Err, objRequisicaoModelo.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174026)
    
    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 61647

    'Limpa a tela
    Call Limpa_Tela_RequisicaoModelo

    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 61647

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174027)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro()

Dim lErro As Long
Dim objRequisicaoModelo As New ClassRequisicaoModelo
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica o preenchimento dos campos obrigatórios
    If Len(Trim(Codigo.Text)) = 0 Then gError 61648
    If Len(Trim(FilialCompra.Text)) = 0 Then gError 61538
    
    'Verifica se o Grid foi preenchido
    If objGridItens.iLinhasExistentes = 0 Then gError 61650
    
    'Se o tipo destino for empresa
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then
    
        'Se a FilialEmpresa não estiver preenchida, erro
        If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 61651
    
    'Se o tipo destino for Fornecedor
    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then
    
        'Se o Fornecedor não estiver preenchido, erro
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 61652
        
        'Se a Filial do Fornecedor não estiver preenchida, erro
        If Len(Trim(FilialFornecedor.Text)) = 0 Then gError 61653
            
    End If
                
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objRequisicaoModelo)
    If lErro <> SUCESSO Then gError 61656
    
    'Se foi escolhido o Requisitante automático, erro
    If objRequisicaoModelo.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO Then gError 67304
    
    'Recolhe os dados do Grid
    lErro = Move_GridItens_Memoria(objRequisicaoModelo)
    If lErro <> SUCESSO Then gError 61665
    
    lErro = Trata_Alteracao(objRequisicaoModelo, objRequisicaoModelo.iFilialEmpresa, objRequisicaoModelo.lCodigo)
    If lErro <> SUCESSO Then gError 32294

    'Grava a Requisição Modelo
    lErro = CF("RequisicaoModelo_Grava", objRequisicaoModelo)
    If lErro <> SUCESSO Then gError 61666
    
    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32294

        Case 61538
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", gErr)

        Case 61648
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
                    
        Case 61650
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDITENS_VAZIO", gErr)
         
        Case 61651
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)
        
        Case 61652
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO", gErr)
        
        Case 61653
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA", gErr)
                        
        Case 61656, 61665, 61666
                
        Case 67304
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_REQUISITANTE_AUTOMATICO", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174028)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 61700
    
    'Limpa a Tela
    Call Limpa_Tela_RequisicaoModelo
   
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
   
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 61700
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174029)

    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_RequisicaoModelo()

Dim iIndice As Integer
Dim bCancel  As Boolean
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RequisicaoModelo

    'Função genérica que limpa a tela
    Call Limpa_Tela(Me)

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
    
    'Limpa o GridItens
    Call Grid_Limpa(objGridItens)

    Set gColItemReqModelo = New Collection
    
    Urgente.Value = vbUnchecked
    TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True
    
    'Limpa as combos
    FilialFornecedor.Clear
    
    'Coloca Filiais default para FilialCompra e FilialEmpresa
    For iIndice = 0 To FilialEmpresa.ListCount - 1
        If giFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text) Then
            FilialEmpresa.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then Error 61812

    'Coloca FilialCompra Default na tela
    If objConfiguraCOM.iFilialCompra > 0 Then
        FilialCompra.Text = objConfiguraCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If
    Call FilialCompra_Validate(bCancel)
        
            
    'Coloca Tipo de tributação Default
    For iIndice = 0 To TipoTributacao.ListCount - 1
        If Codigo_Extrai(TipoTributacao.List(iIndice)) = giTipoTributacao Then
            TipoTributacao.ListIndex = iIndice
            Exit For
        End If
    Next
        
    Exit Sub
    
Erro_Limpa_Tela_RequisicaoModelo:

    Select Case Err
        
        Case 61812
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174030)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Finaliza os objEventos
    Set objEventoCodigo = Nothing
    Set objEventoRequisitante = Nothing
    Set objEventoCclPadrao = Nothing
    Set objEventoObservacao = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoProduto = Nothing
    Set objEventoBotaoCcl = Nothing
    Set objEventoAlmoxarifados = Nothing
    Set objEventoFiliaisFornProduto = Nothing
    Set objEventoContaContabil = Nothing
    
    'Libera variáveis globais
    Set gColItemReqModelo = Nothing
    Set objGridItens = Nothing
    
    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Ccl_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Descricao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCompra_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Urgente_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Click()
    
Dim lErro As Long
Dim iCodFilial As Integer

On Error GoTo Erro_FilialEmpresa_Click
    
    'Se nenhuma FilialEmpresa foi selecionada, sai da rotina
    If FilialEmpresa.ListIndex = -1 Then Exit Sub
    
    'Guarda o código da Filial
    iCodFilial = Codigo_Extrai(FilialEmpresa.Text)
    
    lErro = AlmoxarifadoPadrao_Preenche(iCodFilial)
    If lErro <> SUCESSO Then Error 61707
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub
    
Erro_FilialEmpresa_Click:
    
    Select Case Err
        
        Case 61707 'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174031)
        
    End Select
    
    Exit Sub

End Sub

Function AlmoxarifadoPadrao_Preenche(iCodFilial As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim iAlmoxarifadoPadrao As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoPadrao_Preenche
    
    'Para cada linha do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    
            objProduto.sCodigo = sProdutoFormatado
            
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 61708
    
            'Se não encontrou o Produto --> Erro
            If lErro = 28030 Then Error 61709
            
            'Se o Produto possui Estoque
            If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                'Lê dados do seu almoxarifado Padrão
                lErro = CF("AlmoxarifadoPadrao_Le", iCodFilial, objProduto.sCodigo, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then Error 61710
    
                'Se encontrou
                If lErro = SUCESSO And iAlmoxarifadoPadrao <> 0 Then
    
                    objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
    
                    'Lê os dados do Almoxarifado a partir do código passado
                    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25056 Then Error 61711
            
                    'Se não encontrou, erro
                    If lErro = 25056 Then Error 61712
            
                    'Coloca o Nome Reduzido na Coluna Almoxarifado
                    GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
                Else
                    GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = ""
                End If
                
            
            End If
        
        End If
        
    Next
    
    AlmoxarifadoPadrao_Preenche = SUCESSO
    
    Exit Function
    
Erro_AlmoxarifadoPadrao_Preenche:

    AlmoxarifadoPadrao_Preenche = Err
    
    Select Case Err
        
        Case 61708, 61711
        
        Case 61709
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
            
        Case 61712
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174032)
    
    End Select
    
    Exit Function
    
End Function

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim vbMsgRes As VbMsgBoxResult
Dim iCodFilial As Integer

On Error GoTo Erro_FilialEmpresa_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialEmpresa.Text)) = 0 Then Exit Sub

    'Verifica se é uma FilialEmpresa selecionada
    If FilialEmpresa.Text = FilialEmpresa.List(FilialEmpresa.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61713

    'Se não encontrou o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 61714

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then Error 61715

        'coloca na tela
        FilialEmpresa.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 61716
        
    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True
    
    Select Case Err

        Case 61713, 61714

        Case 61715
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, FilialEmpresa.Text)
            
        Case 61716
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", Err, FilialEmpresa.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174033)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornecedor_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornecedor_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Requisitante_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoDestino_Click(Index As Integer)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TipoDestino_Click

    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame correspondente a Index visivel
    FrameTipoDestino(Index).Visible = True

    'Torna Frame atual invisivel
    FrameTipoDestino(iFrameTipoDestinoAtual).Visible = False

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

    'Se o Destino da não é a própria empresa
    If Index <> TIPO_DESTINO_EMPRESA Then

        'Limpa os almoxarifados e os Ccls do GridItens
        For iIndice = 1 To objGridItens.iLinhasExistentes
            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = ""
            GridItens.TextMatrix(iIndice, iGrid_CCL_Col) = ""
        Next
    Else

        'Seleciona a Filial Empresa na combo
        Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)
        Call FilialEmpresa_Click
    
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_TipoDestino_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174034)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        'Verifica preenchimento de Fornecedor
        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 61597

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then Error 61598

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornecedor, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornecedor, iCodFilial)

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            'Se Fornecedor não foi preenchido limpa a combo de Filiais
            FilialFornecedor.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 61597, 61598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174035)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FilialFornecedor_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornecedor.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornecedor.Text = FilialFornecedor.List(FilialFornecedor.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFornecedor, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61599

    'Se não encontrar o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 61600

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 61601

        'Se não achou a Filial Fornecedor --> erro
        If lErro = 18272 Then Error 61602

        'coloca na tela
        FilialFornecedor.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 61603

    Exit Sub

Erro_FilialFornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 61600
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
    
        Case 61599, 61601
    
        Case 61602
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then

                objFornecedor.sNomeReduzido = Fornecedor.Text

                'Lê Fornecedor no BD
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)

                'Se achou o Fornecedor --> coloca o codigo em objFilialFornecedor
                If lErro = SUCESSO Then objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            
            End If

        Case 61603
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, FilialFornecedor.Text)
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174036)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'Critica o produto passado
        lErro = CF("Produto_Critica_Compra", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25605 Then gError 61610

        'Produto não cadastrado
        If lErro = 25605 Then gError 61611
            
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Verifica se o Produto já está em outra linha do Grid
            For iIndice = 1 To objGridItens.iLinhasExistentes
                If iIndice <> GridItens.Row And Produto.Text = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) Then gError 61810
            Next
            
            'Preenche a UM, a Descrição e o Almoxarifado Padrão do Produto
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 61612

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 61614

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 61613

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 61610, 61612, 61613, 61614
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61611
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 61810
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174037)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UM

    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UM_Col) = UM.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61615

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = Err

    Select Case Err

        Case 61615
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174038)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoTributacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo de Tributação que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TipoTributacao

    Set objGridInt.objControle = TipoTribItem

    GridItens.TextMatrix(GridItens.Row, iGrid_TipoTributacao_Col) = TipoTribItem.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 66606

    Saida_Celula_TipoTributacao = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoTributacao:

    Saida_Celula_TipoTributacao = gErr

    Select Case gErr

        Case 66606
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174039)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then Error 61616

        dQuantidade = CDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61617

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = Err

    Select Case Err

        Case 61616, 61617
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174040)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Almoxarifado do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sContaEnxuta As String

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado foi preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        'Formata o Produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 61618
        
        'Lê o Almoxarifado
        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 61619
        If lErro = 25130 Then gError 61620
        If lErro = 25135 Then gError 61621
        If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 86099

        'Coloca Conta Contábil no GridItens
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 67866
        
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                
        'Se a Conta contábil não foi preenchida
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col))) = 0 Then
        
            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 49991 Then gError 67867
    
            If lErro <> 49991 Then
    
                lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
                If lErro <> SUCESSO Then gError 67868
        
                ContaContabil.PromptInclude = False
                ContaContabil.Text = sContaEnxuta
                ContaContabil.PromptInclude = True
        
                GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
        
            Else
        
                'Preenche em branco a conta de estoque no grid
                GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ""
        
            End If
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 61622

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 61618, 61619, 61622, 67866, 67867, 67868
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61620
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 61621
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 86099
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174041)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Ccl do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = CentroCusto

    'Verifica se Ccl foi preenchido
    If Len(Trim(CentroCusto.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", CentroCusto, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then Error 61623

        If lErro = 5703 Then Error 61624

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61625

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err

    Select Case Err

        Case 61623
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61624
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CentroCusto)
            If vbMsgRes = vbYes Then

                objCcl.sCcl = sCclFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("CclTela", objCcl)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61625
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174042)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaEnxuta As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabil

    Set objGridItens.objControle = ContaContabil

    'Se a Conta Contábil foi preenchida
    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabil.ClipText, objPlanoConta, MODULO_COMPRAS)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 61626

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaEnxuta = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 61627

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_COMPRAS)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 61628

            'conta não cadastrada
            If lErro = 5700 Then Error 61629

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61630

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = Err

    Select Case Err

        Case 61626, 61628, 61630
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61627
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61629
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174043)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Fornecedor(objGridInt As AdmGrid) As Long
'faz a critica da celula fornecedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iFilialEmpresa As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sFornecedor As String

On Error GoTo Erro_Saida_Celula_Fornecedor

    Set objGridInt.objControle = FornecGrid

    sFornecedor = FornecGrid.Text
    
    'Se o fornecedor foi preenchido
    If Len(Trim(FornecGrid.ClipText)) > 0 Then

        'Verifica se o fornecedor está cadastrado
        lErro = TP_Fornecedor_Grid(FornecGrid, objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 25611 And lErro <> 25613 And lErro <> 25616 And lErro <> 25619 Then Error 61631

        'Fornecedor não cadastrado
        'Nome Reduzido
        If lErro = 25611 Then Error 61632

        'Codigo
        If lErro = 25613 Then Error 61633

        'CGC/CPF
        If lErro = 25616 Or lErro = 25619 Then Error 61634

        If sFornecedor <> objFornecedor.sNomeReduzido Then
        
            'Formata o Produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 61814
            
            iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
            
            'Lê coleção de códigos e nomes da Filial do Fornecedor
            lErro = CF("FornecedorProdutoFF_Le_FilialForn", sProdutoFormatado, objFornecedor.lCodigo, Codigo_Extrai(FilialCompra.Text), colCodigoNome)
            If lErro <> SUCESSO Then Error 61635
    
            'Se não encontrou nenhuma Filial, erro
            If colCodigoNome.Count = 0 Then Error 61636
            
            If iCodFilial > 0 Then
    
                For iIndice = 1 To colCodigoNome.Count
                    If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                        Exit For
                    End If
                Next
    
                If iIndice = colCodigoNome.Count Then Error 61637
    
            ElseIf iCodFilial = 0 Then
                iCodFilial = colCodigoNome.Item(1).iCodigo
            End If
            
            For iIndice = 1 To colCodigoNome.Count
                If colCodigoNome.Item(iIndice).iCodigo = iCodFilial Then
                    GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
                    Exit For
                End If
            Next
            
        End If
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col))) = 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = "Preferencial"
        End If
    
    Else
        
        'Limpa a Filial e Exclusividade Correspondente
        GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61638

    Saida_Celula_Fornecedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Fornecedor:

    Saida_Celula_Fornecedor = Err

    Select Case Err

        Case 61631, 61638, 61814
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61632 'Fornecedor com Nome Reduzido %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", FornecGrid.Text)
            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61633 'Fornecedor com código %s não encontrado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", FornecGrid.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61634 'Fornecedor com CGC/CPF %s não encontado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_3", FornecGrid.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Fornecedores", objFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61636
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_FIL_PROD_FORN_FILIALCOMPRA", Err, objFornecedor.sNomeReduzido, sProdutoFormatado)
            FornecGrid.Text = ""
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61637
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", Err, iCodFilial, objFornecedor.sNomeReduzido, sProdutoFormatado)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174044)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialForn(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sNomeFilial As String

On Error GoTo Erro_Saida_Celula_FilialForn

    Set objGridInt.objControle = FilialFornecGrid
    
    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornecGrid.Text)) > 0 Then

        'Verifica se é uma filial selecionada
        If Not FilialFornecGrid.Text = FilialFornecGrid.List(FilialFornecGrid.ListIndex) Then

            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialFornecGrid, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 61639
    
            'Se nao encontra o ítem com o código informado
            If lErro = 6730 Then
    
                'Verifica se o Fornecedor foi preenchido
                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col))) = 0 Then gError 61640
    
                lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 61815
                
                sFornecedor = GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col)
                objFornecedorProdutoFF.iFilialForn = iCodigo
                objFornecedorProdutoFF.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
                objFornecedorProdutoFF.sProduto = sProdutoFormatado
    
                'Pesquisa se existe filial com o codigo extraido
                lErro = CF("FornecedorProdutoFF_Le_NomeRed", sFornecedor, sNomeFilial, objFornecedorProdutoFF)
                If lErro <> SUCESSO And lErro <> 61780 Then gError 61641
    
                'Se não encontrou a Filial do Fornecedor
                If lErro = 61780 Then
    
                    'Lê FilialFornecedor do BD
                    objFilialFornecedor.iCodFilial = iCodigo
                    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
                    If lErro <> SUCESSO And lErro <> 18272 Then gError 61642
    
                    'Se não encontrou, pergunta se deseja criar
                    If lErro = 18272 Then
                        gError 61643
                    
                    'Se encontrou, erro
                    Else
                        gError 61804
                    End If
                
                'Se encontrou a Filial do Fornecedor
                Else
    
                    'coloca na tela
                    FilialFornecGrid.Text = iCodigo & SEPARADOR & sNomeFilial
    
                End If
    
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 61644

        End If
    Else
        gError 86096
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 61813

    Saida_Celula_FilialForn = SUCESSO
    
    Exit Function

Erro_Saida_Celula_FilialForn:

    Saida_Celula_FilialForn = gErr
    
    Select Case gErr

        Case 61639, 61641, 61642, 61813, 61815
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61640
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_FORNECEDOR_NAO_PREENCHIDO", gErr, GridItens.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61643

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then

                objFornecedor.sNomeReduzido = Fornecedor.Text

                'Lê Fornecedor no BD
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)

                'Se achou o Fornecedor --> coloca o codigo em objFilialFornecedor
                If lErro = SUCESSO Then objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 61644
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_NAO_ENCONTRADA_ASSOCIADA", gErr, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 61804
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS", gErr, objFilialFornecedor.iCodFilial, sFornecedor, objFornecedorProdutoFF.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 86096
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174045)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Exclusivo(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Exclusivo do grid que está deixando de ser a corrente
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Exclusivo

    Set objGridInt.objControle = Exclusivo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61645

    Saida_Celula_Exclusivo = SUCESSO

    Exit Function

Erro_Saida_Celula_Exclusivo:

    Saida_Celula_Exclusivo = Err

    Select Case Err

        Case 61645
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174046)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Observacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = ObservacaoGrid

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61646

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = Err

    Select Case Err

        Case 61646
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174047)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AO SISTEMA DE BROWSE "'
'""""""""""""""""""""""""""""""""""""""""""""""
Private Sub CodigoLabel_Click()

Dim colSelecao As New Collection
Dim objRequisicaoModelo As New ClassRequisicaoModelo

    'Verifica se o código do Modelo foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then objRequisicaoModelo.lCodigo = CLng(Codigo.Text)

    Call Chama_Tela("RequisicaoModeloLista", colSelecao, objRequisicaoModelo, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisicaoModelo As New ClassRequisicaoModelo

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRequisicaoModelo = obj1

    'Traz dados da Requisição Modelo para a tela
    lErro = Traz_RequisicaoModelo_Tela(objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61551

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 61551

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174048)

    End Select

    Exit Sub

End Sub

Private Sub LabelRequisitante_Click()

Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

    'Se o Requisitante estiver preenchido
    If Len(Trim(Requisitante.Text)) > 0 Then objRequisitante.sNomeReduzido = Requisitante.Text

    'Chama o Browser que Lista os Requisitantes
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoRequisitante)

End Sub

Private Sub objEventoRequisitante_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim bCancel As Boolean

    Set objRequisitante = obj1

    'Colocao Nome Reduzido do Requisitante na tela
    Requisitante.Text = objRequisitante.sNomeReduzido

    'Dispara o Validate de Requisitante
    Call Requisitante_Validate(bCancel)

    Me.Show

End Sub

Private Sub CclPadraoLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_LabelCcl_Click

    'Critica o formato do centro de custo
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 61552

    objCcl.sCcl = sCclFormatada

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

    Exit Sub

Erro_LabelCcl_Click:

    Select Case Err

        Case 61552

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174049)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    'Coloca a conta no formato conta enxuta
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 61553

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case Err

        Case 61553

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174050)

        End Select

    Exit Sub

End Sub

Private Sub ObservacaoLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objObservacao As New ClassObservacao

    objObservacao.sObservacao = Observacao.Text

    Call Chama_Tela("ObservacaoLista", colSelecao, objObservacao, objEventoObservacao)

End Sub

Private Sub objEventoObservacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objObservacao As New ClassObservacao

    Set objObservacao = obj1

    Observacao.Text = objObservacao.sObservacao

    Me.Show

End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Public Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo
    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 61554

'    'Verifica se o Produto está preenchido
'    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then
'
'        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
'        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
'
'    End If

    '###############################################
    'Inserido por Wagner 05/05/06
    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 177417
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
    '###############################################

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoLista_Consulta
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 61554
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 177417

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174051)

    End Select

    Exit Sub

End Sub

Public Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoEnxuto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 61555

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoEnxuto = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then Error 61556

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 61557

            'Se não encontrou o Produto --> Erro
            If lErro = 28030 Then Error 61558

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is Produto) Then

                GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then Error 61870

            End If

        End If

    End If

    'Se necessário incrementa-se o número de linha existentes
    If GridItens.Row > objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        gColItemReqModelo.Add (0) 'Inserido por Wagner 09/05/2006
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 61555, 61557, 61870

        Case 61556
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 61558
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174052)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sContaEnxuta As String
Dim objProdutoFilial As New ClassProdutoFilial
Dim objFornecedor As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_ProdutoLinha_Preenche

    'Preenche no Grid a Descrição do Produto e a Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = objProduto.sSiglaUMCompra
    GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao

    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True Then
    
        If Len(Trim(FilialEmpresa.Text)) > 0 Then
            lErro = CF("AlmoxarifadoPadrao_Le", Codigo_Extrai(FilialEmpresa.Text), objProduto.sCodigo, iAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 61559
        
            If lErro = SUCESSO And iAlmoxarifado <> 0 Then
        
                objAlmoxarifado.iCodigo = iAlmoxarifado
        
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 61560
        
                If lErro = 25056 Then gError 61561
        
                'Coloca o Nome Reduzido na Coluna Almoxarifado
                GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
        
                'Coloca Conta Contábil no GridItens
                objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
                objEstoqueProduto.sProduto = objProduto.sCodigo
                lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                If lErro <> SUCESSO And lErro <> 49991 Then gError 66379
    
                If lErro <> 49991 Then
    
                    lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
                    If lErro <> SUCESSO Then gError 65598
            
                    ContaContabil.PromptInclude = False
                    ContaContabil.Text = sContaEnxuta
                    ContaContabil.PromptInclude = True
            
                    GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
            
                Else
            
                    'Preenche em branco a conta de estoque no grid
                    GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ""
            
                End If
            End If
        End If
        
        'Preenche Ccl
        If Len(Trim(Ccl.ClipText)) > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col) = Ccl.Text
        End If

    End If
            
    'Tipo de Tributação
    If Len(Trim(TipoTributacao.Text)) > 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_TipoTributacao_Col) = TipoTributacao.Text
    End If
        
    objProdutoFilial.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
    objProdutoFilial.sProduto = objProduto.sCodigo

    'Busca a filial fornecedor padrão da filialempresa
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 62666
    If lErro = SUCESSO Then
        'Se há filial fornecedor padrão
        If (objProdutoFilial.lFornecedor > 0) And (objProdutoFilial.iFilialForn > 0) Then
            
            objFornecedor.lCodigo = objProdutoFilial.lFornecedor
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 62667
            If lErro <> SUCESSO Then gError 62668
            
            objFilialForn.lCodFornecedor = objFornecedor.lCodigo
            objFilialForn.iCodFilial = objProdutoFilial.iFilialForn
            'Lê a filial do fornecedor
            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 62669
            If lErro <> SUCESSO Then gError 62670
                
            'Coloca no grid a filial fornecedor
            GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
            GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = objFilialForn.iCodFilial & SEPARADOR & objFilialForn.sNome
            GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = "Preferencial"
                
        End If
    End If
    
    'Se necessário cria uma nova linha no Grid
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        gColItemReqModelo.Add (0)
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 61559, 61560, 66379, 65598, 62666, 62667, 62669

        Case 61561
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case 62670
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialForn.iCodFilial, objFilialForn.lCodFornecedor)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174053)

    End Select

    Exit Function

End Function

Private Sub BotaoCcl_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_BotaoCcls_Click

    'Se nenhuma linha foi selecionada do Grid, Erro
    If GridItens.Row = 0 Then gError 61562

    'Se o campo está desabilitado, sai da rotina
    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = False Then gError 76126

    'Verifica se o Produto foi preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then gError 61563

    'Verifica se o Ccl Foi preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col))) > 0 Then

        sCclFormatada = String(STRING_CCL, 0)

        lErro = CF("Ccl_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col), sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 61564

        objCcl.sCcl = sCclFormatada

    End If

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoBotaoCcl)

    Exit Sub

Erro_BotaoCcls_Click:

    Select Case gErr

        Case 61562
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 61563
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 61564

        Case 76126
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCALENTREGA_DIFERENTE_FILIALEMPRESA", gErr)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174054)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se nenhuma foi selecionada, erro
    If GridItens.Row = 0 Then Error 61565

    sContaEnxuta = String(STRING_CCL, 0)

    'Coloca a conta no formato conta enxuta
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 61566

    GridItens.TextMatrix(GridItens.Row, iGrid_CCL_Col) = sCclMascarado
    CentroCusto.PromptInclude = False
    CentroCusto.Text = sCclMascarado
    CentroCusto.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case Err

        Case 61565
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 61566

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174055)

        End Select

    Exit Sub

End Sub

Private Sub BotaoAlmoxarifados_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iFilialEntrega As Integer

On Error GoTo Erro_BotaoALmoxarifados_Click

    If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = False Then gError 76125

    If GridItens.Row = 0 Then gError 61567

    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    sProdutoFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 61569


    'Carrega a Variável com os dados do frame visível
    iFilialEntrega = Codigo_Extrai(FilialEmpresa.Text)
    If Len(Trim(iFilialEntrega)) = 0 Then gError 84602
    
    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        
        colSelecao.Add iFilialEntrega
        
        Call Chama_Tela("AlmoxarifadoFilialLista", colSelecao, objEstoqueProduto, objEventoAlmoxarifados)

    Else
        gError 61568
    End If

    Exit Sub

Erro_BotaoALmoxarifados_Click:

    Select Case gErr

        Case 84602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_ENTREGA_NAO_PREENCHIDA", gErr)
        
        Case 61567
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 61568
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 61569

        Case 76125
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCALENTREGA_DIFERENTE_FILIALEMPRESA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174056)

    End Select

    Exit Sub

End Sub

Public Sub objEventoAlmoxarifados_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto
Dim bCancel As Boolean

    Set objEstoqueProduto = obj1

    'Preenche campo Almoxarifado
    GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    Me.Show

    Exit Sub

End Sub

Private Sub BotaoFiliaisFornProd_Click()

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoFiliaisFornProd_Click

    'Se nenhuma linha foi preenchida, erro
    If GridItens.Row = 0 Then Error 61570

    'Se a FilialCompra não foi preenchida, erro
    If Len(Trim(FilialCompra.Text)) = 0 Then Error 61345
    
    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 61571
    
    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        colSelecao.Add Codigo_Extrai(FilialCompra.Text)
        
        Call Chama_Tela("FiliaisFornProdutoLista", colSelecao, objFornecedorProdutoFF, objEventoFiliaisFornProduto)
    Else
        Error 61572
    End If

    Exit Sub

Erro_BotaoFiliaisFornProd_Click:

    Select Case Err

        Case 61345
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", Err)
            
        Case 61571

        Case 61570
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 61572
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174057)

    End Select

    Exit Sub

End Sub

Public Sub objEventoFiliaisFornProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim objFornecedor As New ClassFornecedor
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_objEventoFiliaisFornProd_evSelecao

    Set objFornecedorProdutoFF = obj1

    'Lê o Nome Reduzido do Fornecedor
    objFornecedor.lCodigo = objFornecedorProdutoFF.lFornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO Then gError 61573

    'Preenche campo Fornecedor
    GridItens.TextMatrix(GridItens.Row, iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido

    'Lê coleção de códigos e nomes da Filial do Fornecedor
    lErro = CF("FornecedorProdutoFF_Le_FilialForn", objFornecedorProdutoFF.sProduto, objFornecedor.lCodigo, objFornecedorProdutoFF.iFilialEmpresa, colCodigoNome)
    If lErro <> SUCESSO Then gError 61574
    
    'Se não encontrou nenhuma Filial, erro
    If colCodigoNome.Count = 0 Then gError 66639
    
    'Se foi passada um Filial como parâmetro
    If objFornecedorProdutoFF.iFilialForn > 0 Then

        'Verifica se ela está presente na coleção de filiais
        For iIndice = 1 To colCodigoNome.Count
            If colCodigoNome.Item(iIndice).iCodigo = objFornecedorProdutoFF.iFilialForn Then
                Exit For
            End If
        Next
    
        'Se não encontrou, erro
        If iIndice > colCodigoNome.Count Then gError 66640
    
    'Se não foi passada uma filial como parâmetro
    ElseIf objFornecedorProdutoFF.iFilialForn = 0 Then
        'Coloca como default a primeira filial da coleção
        objFornecedorProdutoFF.iFilialForn = colCodigoNome.Item(1).iCodigo
    End If

    'Coloca no Grid a filial passada
    For iIndice = 1 To colCodigoNome.Count
        If colCodigoNome.Item(iIndice).iCodigo = objFornecedorProdutoFF.iFilialForn Then
            GridItens.TextMatrix(GridItens.Row, iGrid_FilialFornecedor_Col) = CStr(colCodigoNome.Item(iIndice).iCodigo) & SEPARADOR & colCodigoNome.Item(iIndice).sNome
            Exit For
        End If
    Next
            
    'Se não foi preenchida a exclusividade, coloca como default "Preferencial"
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col))) = 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Exclusivo_Col) = "Preferencial"
    End If

    Me.Show

    Exit Sub

Erro_objEventoFiliaisFornProd_evSelecao:

    Select Case gErr

        Case 61573, 61574, 66639, 66640

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174058)

    End Select

    Exit Sub

End Sub

Private Sub Requisitante_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclEnxuta As String
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Requisitante_Validate

    'Se o Requisitante não foi preenchido, sai da rotina
    If Len(Trim(Requisitante.Text)) = 0 Then Exit Sub

    lErro = TP_Requisitante_Le(Requisitante, objRequisitante)
    If lErro <> SUCESSO Then gError 61583

    If Len(Trim(objRequisitante.sCcl)) > 0 Then
        
        lErro = Mascara_RetornaCclEnxuta(objRequisitante.sCcl, sCclEnxuta)
        If lErro <> SUCESSO Then gError 79987
        
        Ccl.PromptInclude = False
        Ccl.Text = sCclEnxuta
        Ccl.PromptInclude = True
    End If

    Exit Sub

Erro_Requisitante_Validate:

    Cancel = True
    
    Select Case gErr

        Case 61583
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174059)

    End Select

    Exit Sub

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Ccl_Validate

    'Se o Ccl não está preenchido, sai da rotina
    If Len(Trim(Ccl.ClipText)) = 0 Then Exit Sub

    'Critica o Ccl
    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then Error 61584

    'Se o Ccl não está cadastrado, erro
    If lErro = 5703 Then Error 61585

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True
    
    Select Case Err

        Case 61584

        Case 61585

            'Pergunta se deseja cadastrar nova Ccl
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsg = vbYes Then
                objCcl.sCcl = sCclFormatada
                Call Chama_Tela("CclTela", objCcl)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174060)

    End Select

    Exit Sub

End Sub

Private Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCompra_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialCompra.Text)) = 0 Then Exit Sub

    'Verifica se é uma FilialEmpresa selecionada
    If FilialCompra.Text = FilialCompra.List(FilialCompra.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCompra, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61589

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.lCodEmpresa = glEmpresa
        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 61586

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then Error 61587

        'coloca na tela
        FilialCompra.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 61588

    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case Err

        Case 61586, 61589
            
        Case 61587
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, FilialCompra.Text)

        Case 61588
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", Err, FilialCompra.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174061)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAtual As Integer
Dim lErro As Long

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iLinhaAtual = GridItens.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        gColItemReqModelo.Remove (iLinhaAtual)

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174062)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CentroCusto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CentroCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CentroCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CentroCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CentroCusto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabil_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ContaContabil_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = ContaContabil
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoTribItem_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoTribItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub TipoTribItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub TipoTribItem_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = TipoTribItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FornecGrid_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub FornecGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub FornecGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub FornecGrid_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = FornecGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialFornecGrid_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub FilialFornecGrid_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornecGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub FilialFornecGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub FilialFornecGrid_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = FilialFornecGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Exclusivo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Exclusivo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Exclusivo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Exclusivo_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = Exclusivo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoGrid_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ObservacaoGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ObservacaoGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ObservacaoGrid_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = ObservacaoGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is Requisitante Then
            Call LabelRequisitante_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is Observacao Then
            Call ObservacaoLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is FornecGrid Then
            Call BotaoFiliaisFornProd_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoAlmoxarifados_Click
        ElseIf Me.ActiveControl Is CentroCusto Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call BotaoPlanoConta_Click
        End If
    End If

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Modelo de Requisição"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RequisicaoModelo"

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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelRequisitante_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRequisitante, Source, X, Y)
End Sub

Private Sub LabelRequisitante_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRequisitante, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub ObservacaoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ObservacaoLabel, Source, X, Y)
End Sub

Private Sub ObservacaoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ObservacaoLabel, Button, Shift, X, Y)
End Sub

Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
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

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = DescProduto

    If Len(Trim(DescProduto.Text)) = 0 Then gError 86173

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53644

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 53644 'Erro tratado na rotina chamada.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 86173
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174063)

    End Select

    Exit Function

End Function

