VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ExportarProjetosOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   2
      Left            =   165
      TabIndex        =   13
      Top             =   690
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame Frame2 
         Caption         =   "Itens de Projeto a Exportar"
         Height          =   5085
         Left            =   15
         TabIndex        =   14
         Top             =   60
         Width           =   9100
         Begin VB.CommandButton BotaoProjeto 
            Caption         =   "Ver Projeto"
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
            Left            =   7845
            TabIndex        =   57
            ToolTipText     =   "Abre a tela de Projetos"
            Top             =   4425
            Width           =   1125
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   315
            Left            =   6915
            TabIndex        =   56
            Top             =   1740
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProdItens 
            Height          =   315
            Left            =   7485
            TabIndex        =   55
            Top             =   2175
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataInicio 
            Height          =   315
            Left            =   3180
            TabIndex        =   50
            Top             =   2625
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataTermino 
            Height          =   315
            Left            =   4305
            TabIndex        =   49
            Top             =   2625
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoItens 
            Height          =   315
            Left            =   555
            TabIndex        =   54
            Top             =   2175
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox VersaoProdItens 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "ExportarProjetos.ctx":0000
            Left            =   2055
            List            =   "ExportarProjetos.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   2175
            Width           =   930
         End
         Begin VB.ComboBox UMProdItens 
            Height          =   315
            Left            =   6600
            TabIndex        =   52
            Top             =   2175
            Width           =   885
         End
         Begin VB.TextBox DescricaoProdItens 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2940
            TabIndex        =   51
            Top             =   2175
            Width           =   3660
         End
         Begin VB.ComboBox Destino 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "ExportarProjetos.ctx":0004
            Left            =   4935
            List            =   "ExportarProjetos.ctx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   1740
            Width           =   1980
         End
         Begin VB.CheckBox Exportar 
            Height          =   315
            Left            =   4200
            TabIndex        =   47
            Top             =   1740
            Width           =   720
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
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
            Left            =   120
            Picture         =   "ExportarProjetos.ctx":0008
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   4425
            Width           =   1665
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
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
            Left            =   1845
            Picture         =   "ExportarProjetos.ctx":1022
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   4425
            Width           =   1665
         End
         Begin VB.CommandButton BotaoRelatorio 
            Caption         =   "Itens Exportados"
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
            Left            =   6660
            TabIndex        =   44
            ToolTipText     =   "Abre a tela do Relatório de Itens Exportados"
            Top             =   4425
            Width           =   1125
         End
         Begin VB.CommandButton BotaoExportar 
            Caption         =   "Exportar"
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
            Left            =   5535
            TabIndex        =   43
            ToolTipText     =   "Faz a Exportação para o Destino selecionado"
            Top             =   4425
            Width           =   1050
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   4095
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   1
      Left            =   165
      TabIndex        =   15
      Top             =   750
      Width           =   9165
      Begin VB.CheckBox RelatorioAutomatico 
         Caption         =   "Gerar Relatório de Itens Exportados após a Exportação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5460
         TabIndex        =   60
         Top             =   4500
         Width           =   2970
      End
      Begin VB.Frame Frame4 
         Caption         =   "Destino"
         Height          =   720
         Left            =   5310
         TabIndex        =   58
         Top             =   3570
         Width           =   3180
         Begin VB.ComboBox DestinoSeleciona 
            Height          =   315
            ItemData        =   "ExportarProjetos.ctx":2204
            Left            =   510
            List            =   "ExportarProjetos.ctx":2206
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   285
            Width           =   2325
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Termino"
         Height          =   720
         Left            =   525
         TabIndex        =   36
         Top             =   4320
         Width           =   4650
         Begin MSMask.MaskEdBox DataTerminoInicial 
            Height          =   300
            Left            =   705
            TabIndex        =   37
            Top             =   270
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataTerminoInicial 
            Height          =   300
            Left            =   1875
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataTerminoFinal 
            Height          =   300
            Left            =   3045
            TabIndex        =   39
            Top             =   270
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataTerminoFinal 
            Height          =   300
            Left            =   4215
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelDataTerminoDe 
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
            Left            =   345
            TabIndex        =   42
            Top             =   285
            Width           =   315
         End
         Begin VB.Label LabelDataTerminoAte 
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
            Left            =   2640
            TabIndex        =   41
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Frame FrameDataInicio 
         Caption         =   "Data Inicio"
         Height          =   720
         Left            =   540
         TabIndex        =   27
         Top             =   3570
         Width           =   4635
         Begin MSMask.MaskEdBox DataIniInicial 
            Height          =   300
            Left            =   705
            TabIndex        =   4
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataIniInicial 
            Height          =   300
            Left            =   1860
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataIniFinal 
            Height          =   300
            Left            =   3045
            TabIndex        =   5
            Top             =   285
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataIniFinal 
            Height          =   300
            Left            =   4215
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelDataIniAte 
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
            Left            =   2625
            TabIndex        =   29
            Top             =   300
            Width           =   360
         End
         Begin VB.Label LabelDataIniDe 
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
            Left            =   345
            TabIndex        =   28
            Top             =   270
            Width           =   315
         End
      End
      Begin VB.Frame FrameProjeto 
         Caption         =   "Projeto"
         Height          =   1155
         Left            =   555
         TabIndex        =   26
         Top             =   1140
         Width           =   7950
         Begin MSMask.MaskEdBox ProjetoInicial 
            Height          =   315
            Left            =   705
            TabIndex        =   30
            Top             =   255
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProjetoFinal 
            Height          =   315
            Left            =   705
            TabIndex        =   31
            Top             =   675
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label DescProjetoFinal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2295
            TabIndex        =   35
            Top             =   675
            Width           =   5535
         End
         Begin VB.Label DescProjetoInicio 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2295
            TabIndex        =   34
            Top             =   255
            Width           =   5535
         End
         Begin VB.Label LabelProjetoDe 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   33
            Top             =   285
            Width           =   360
         End
         Begin VB.Label LabelProjetoAte 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   720
            Width           =   435
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente"
         Height          =   1140
         Left            =   570
         TabIndex        =   21
         Top             =   -30
         Width           =   7935
         Begin MSMask.MaskEdBox ClienteInicial 
            Height          =   315
            Left            =   705
            TabIndex        =   0
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteFinal 
            Height          =   315
            Left            =   705
            TabIndex        =   1
            Top             =   660
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelClienteAte 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   705
            Width           =   435
         End
         Begin VB.Label LabelClienteDe 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   270
            Width           =   360
         End
         Begin VB.Label DescClienteInicial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2295
            TabIndex        =   23
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label DescClienteFinal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   22
            Top             =   660
            Width           =   5535
         End
      End
      Begin VB.Frame FrameProdutos 
         Caption         =   "Produtos"
         Height          =   1170
         Index           =   0
         Left            =   555
         TabIndex        =   16
         Top             =   2355
         Width           =   7935
         Begin MSMask.MaskEdBox ProdutoInicial 
            Height          =   315
            Left            =   705
            TabIndex        =   2
            Top             =   255
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoFinal 
            Height          =   315
            Left            =   705
            TabIndex        =   3
            Top             =   675
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label DescProdFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   20
            Top             =   675
            Width           =   5535
         End
         Begin VB.Label DescProdInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   19
            Top             =   255
            Width           =   5535
         End
         Begin VB.Label LabelProdutoDe 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   18
            Top             =   285
            Width           =   360
         End
         Begin VB.Label LabelProdutoAte 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   17
            Top             =   720
            Width           =   435
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   8280
      ScaleHeight     =   465
      ScaleWidth      =   1065
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   45
      Width           =   1125
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "ExportarProjetos.ctx":2208
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   570
         Picture         =   "ExportarProjetos.ctx":273A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5595
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens de Projetos"
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
Attribute VB_Name = "ExportarProjetosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iTabPrincipalAlterado As Integer

Dim iFrameAtual As Integer

'Grid de Itens
Dim objGridItens As AdmGrid
Dim iGrid_Exportar_Col As Integer
Dim iGrid_Destino_Col As Integer
Dim iGrid_Projeto_Col As Integer
Dim iGrid_ProdutoItens_Col As Integer
Dim iGrid_VersaoProdItens_Col As Integer
Dim iGrid_DescricaoProdItens_Col As Integer
Dim iGrid_UMProdItens_Col As Integer
Dim iGrid_QuantidadeProdItens_Col As Integer
Dim iGrid_DataInicio_Col As Integer
Dim iGrid_DataTermino_Col As Integer

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1

Private WithEvents objEventoProjetoDe As AdmEvento
Attribute objEventoProjetoDe.VB_VarHelpID = -1
Private WithEvents objEventoProjetoAte As AdmEvento
Attribute objEventoProjetoAte.VB_VarHelpID = -1

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Private Const TAB_Selecao = 1
Private Const TAB_ItensProjetos = 2
Private Sub BotaoExportar_Click()

Dim lErro As Long
Dim iCont As Integer
Dim iLinha As Integer
Dim sMsgDestinoNaoPreenchido
Dim iIndice As Integer
Dim colOrcamentoVenda As New Collection
Dim colPedidoVenda As New Collection
Dim colOrdemDeProducao As New Collection
Dim colNFSimples As New Collection
Dim colNFFatura As New Collection
Dim colOrdemServico As New Collection
   
On Error GoTo Erro_BotaoExportar_Click
           
    GL_objMDIForm.MousePointer = vbHourglass
           
    iCont = 0

    'Verifica se os destinos estão preenchidos
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'se marcado para exportar...
        If GridItens.TextMatrix(iLinha, iGrid_Exportar_Col) = MARCADO Then

            iCont = iCont + 1

            'e não tem destino...
            If Len(GridItens.TextMatrix(iLinha, iGrid_Destino_Col)) = 0 Then

                'monta a linha para mensagem
                If Len(sMsgDestinoNaoPreenchido) = 0 Then
                    sMsgDestinoNaoPreenchido = CStr(iLinha)
                Else
                    sMsgDestinoNaoPreenchido = sMsgDestinoNaoPreenchido & ", " & CStr(iLinha)
                End If
            End If
        End If
    Next

    'se tem mensagem de erro... exibi-la
    If Len(sMsgDestinoNaoPreenchido) <> 0 Then

        'colocar um " e" no lugar da última "," na mensagem
        For iIndice = Len(sMsgDestinoNaoPreenchido) To 1 Step -1
    
            If Mid(sMsgDestinoNaoPreenchido, iIndice, 1) = "," Then
                sMsgDestinoNaoPreenchido = Mid(sMsgDestinoNaoPreenchido, 1, iIndice - 1) & " e" & Right(sMsgDestinoNaoPreenchido, Len(sMsgDestinoNaoPreenchido) - iIndice)
                Exit For
            End If
    
        Next iIndice
        
        gError 137429
    
    End If

    'se não tem nenhum marcado... erro
    If iCont = 0 Then gError 137400
    
    'Agrupa os destinos em coleções para exportar...
    lErro = Agrupa_Destinos(colOrcamentoVenda, colPedidoVenda, colOrdemDeProducao, colNFSimples, colNFFatura, colOrdemServico)
    If lErro <> SUCESSO Then gError 137401

    If colOrcamentoVenda.Count > 0 Then

        'Exporta para a tela de orçamento de venda
        lErro = Exporta_OrcamentoVenda(colOrcamentoVenda)
        If lErro <> SUCESSO Then gError 137402

    End If

    If colPedidoVenda.Count > 0 Then

        'Exporta para a tela de pedido de venda
        lErro = Exporta_PedidoVenda(colPedidoVenda)
        If lErro <> SUCESSO Then gError 137403

    End If

    If colOrdemDeProducao.Count > 0 Then

        'Exporta para a tela de ordem de producao
        lErro = Exporta_OrdemDeProducao(colOrdemDeProducao)
        If lErro <> SUCESSO Then gError 137404

    End If

    If colNFSimples.Count > 0 Then

        'Exporta para a tela de nota fiscal simples
        lErro = Exporta_NFSimples(colNFSimples)
        If lErro <> SUCESSO Then gError 137405

    End If

    If colNFFatura.Count > 0 Then

        'Exporta para a tela de nota fiscal fatura
        lErro = Exporta_NFFatura(colNFFatura)
        If lErro <> SUCESSO Then gError 137406

    End If

    If colOrdemServico.Count > 0 Then

        'Exporta para a tela de ordem de servico
'        lErro = Exporta_OrdemServico(colOrdemServico)
        If lErro <> SUCESSO Then gError 137407

    End If
    
    'se é para imprimir o relatório ao final
    If RelatorioAutomatico.Value = vbChecked Then
    
        'chama o relatorio de Itens Exportados
        '----------------------------------------------------
        '----------------------------------------------------
    
    End If
    
    'Limpa a Tela
    lErro = Limpa_Tela_ExportarProjetos
    If lErro <> SUCESSO Then gError 137409
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExportar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 137400
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_ITENS_A_EXPORTAR", gErr)
            
        Case 137401 To 137409
            'erros tratados nas rotinas chamadas

        Case 137429
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_PREENCHIDO", gErr, sMsgDestinoNaoPreenchido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159836)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159837)

    End Select

    Exit Sub

End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Exportar Projetos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ExportarProjetos"

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

Private Sub BotaoProjeto_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjeto
   
On Error GoTo Erro_BotaoCusteio_Click
    
    'Se não tiver linha selecionada => Erro
    If GridItens.Row = 0 Then gError 136395
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Projeto_Col)) = 0 Then gError 137428

    objProjeto.sNomeReduzido = Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Projeto_Col))
       
    'Chama a tela de projeto
    Call Chama_Tela("Projeto", objProjeto)

    Exit Sub

Erro_BotaoCusteio_Click:

    Select Case gErr

        Case 136395
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 137428
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159838)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ClienteInicial_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
    Call Cliente_Preenche(ClienteInicial)

End Sub

Private Sub ClienteInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ClienteInicial, iAlterado)
    
End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    DescClienteInicial.Caption = ""

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 137599
        
        ClienteInicial.Text = objCliente.sNomeReduzido
        DescClienteInicial.Caption = objCliente.sRazaoSocial

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137599
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159839)

    End Select

End Sub

Private Sub ClienteFinal_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
    'por Jorge Specian
    Call Cliente_Preenche(ClienteFinal)

End Sub

Private Sub ClienteFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ClienteFinal, iAlterado)
    
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    DescClienteFinal.Caption = ""
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 137600
        
        ClienteFinal.Text = objCliente.sNomeReduzido
        DescClienteFinal.Caption = objCliente.sRazaoSocial

    End If
        
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137600
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159840)

    End Select

End Sub

Private Sub DestinoSeleciona_Click()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click
    
    'Verifica se Cliente Final foi preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then
    
        If IsNumeric(ClienteFinal.Text) Then

            objCliente.lCodigo = StrParaLong(Trim(ClienteFinal.Text))
            
        Else
        
            objCliente.sNomeReduzido = Trim(ClienteFinal.Text)
        
        End If

    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159841)

    End Select

    Exit Sub


End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click
        
    'Verifica se Cliente Inicial foi preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        If IsNumeric(ClienteInicial.Text) Then

            objCliente.lCodigo = StrParaLong(Trim(ClienteInicial.Text))

        Else

            objCliente.sNomeReduzido = Trim(ClienteInicial.Text)

        End If

    End If
        
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159842)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137601

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 137601
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159843)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137602

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 137602
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159844)

    End Select

    Exit Sub

End Sub

Private Sub LabelProjetoAte_Click()

Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o projeto foi preenchido
    If Len(ProjetoFinal.Text) <> 0 Then

        Set objProjeto = New ClassProjeto
        
        'Verifica sua existencia
        lErro = CF("TP_Projeto_Le", ProjetoFinal, objProjeto)
        If lErro <> SUCESSO Then gError 134467
        
    End If

    Call Chama_Tela("ProjetoLista", colSelecao, objProjeto, objEventoProjetoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 137602
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159845)

    End Select

    Exit Sub

End Sub

Private Sub LabelProjetoDe_Click()

Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o projeto foi preenchido
    If Len(ProjetoInicial.Text) <> 0 Then

        Set objProjeto = New ClassProjeto
        
        'Verifica sua existencia
        lErro = CF("TP_Projeto_Le", ProjetoInicial, objProjeto)
        If lErro <> SUCESSO Then gError 134467
        
    End If

    Call Chama_Tela("ProjetoLista", colSelecao, objProjeto, objEventoProjetoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 137602
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159846)

    End Select

    Exit Sub

End Sub


Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteInicial.Text = CStr(objCliente.lCodigo)
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137604

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 137605

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 137606

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 137604, 137606
            'erros tratados nas rotinas chamadas

        Case 137605
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159847)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137607

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 137608

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 137609

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 137607, 137609
            'erro tratado na rotina chamada

        Case 137608
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159848)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProjetoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As New ClassProjeto

On Error GoTo Erro_objEventoProjetoAte_evSelecao

    Set objProjeto = obj1
    
    ProjetoFinal.Text = objProjeto.sNomeReduzido
    
    Call ProjetoFinal_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProjetoAte_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159849)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProjetoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As New ClassProjeto

On Error GoTo Erro_objEventoProjetoDe_evSelecao

    Set objProjeto = obj1
    
    ProjetoInicial.Text = objProjeto.sNomeReduzido
    
    Call ProjetoInicial_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProjetoDe_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159850)

    End Select

    Exit Sub

End Sub

Private Sub ProjetoFinal_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProjetoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProjetoInicial, iAlterado)
    
End Sub


Private Sub ProjetoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProjeto As ClassProjeto

On Error GoTo Erro_Projeto_Validate

    DescProjetoFinal.Caption = ""

    'Verifica se Projeto está preenchido
    If Len(Trim(ProjetoFinal.Text)) > 0 Then
    
        Set objProjeto = New ClassProjeto
        
        'Verifica sua existencia
        lErro = CF("TP_Projeto_Le", ProjetoFinal, objProjeto)
        If lErro <> SUCESSO Then gError 134467
        
        'Coloca a descrição na tela
        DescProjetoFinal.Caption = objProjeto.sDescricao
        
    End If
    
    Exit Sub

Erro_Projeto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 134467
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159851)

    End Select

    Exit Sub

End Sub

Private Sub ProjetoInicial_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProjetoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProjetoInicial, iAlterado)
    
End Sub


Private Sub ProjetoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProjeto As ClassProjeto

On Error GoTo Erro_Projeto_Validate

    DescProjetoInicio.Caption = ""

    'Verifica se Projeto está preenchido
    If Len(Trim(ProjetoInicial.Text)) > 0 Then
    
        Set objProjeto = New ClassProjeto
        
        'Verifica sua existencia
        lErro = CF("TP_Projeto_Le", ProjetoInicial, objProjeto)
        If lErro <> SUCESSO Then gError 134467
        
        'Coloca a descrição na tela
        DescProjetoInicio.Caption = objProjeto.sDescricao
        
    End If
    
    Exit Sub

Erro_Projeto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 134467
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159852)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataIniFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataIniFinal_DownClick

    DataIniFinal.SetFocus

    If Len(DataIniFinal.ClipText) > 0 Then

        sData = DataIniFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137610

        DataIniFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataIniFinal_DownClick:

    Select Case gErr

        Case 137610

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159853)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataIniFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataIniFinal_UpClick

    DataIniFinal.SetFocus

    If Len(Trim(DataIniFinal.ClipText)) > 0 Then

        sData = DataIniFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137611

        DataIniFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataIniFinal_UpClick:

    Select Case gErr

        Case 137611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159854)

    End Select

    Exit Sub

End Sub

Private Sub DataIniInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataIniInicial, iAlterado)
    
End Sub

Private Sub DataIniInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataIniInicial_Validate

    If Len(Trim(DataIniInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataIniInicial.Text)
        If lErro <> SUCESSO Then gError 137612

    End If

    Exit Sub

Erro_DataIniInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137612

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159855)

    End Select

    Exit Sub

End Sub

Private Sub DataIniInicial_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataIniInicial_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataIniInicial_DownClick

    DataIniInicial.SetFocus

    If Len(DataIniInicial.ClipText) > 0 Then

        sData = DataIniInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137613

        DataIniInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataIniInicial_DownClick:

    Select Case gErr

        Case 137613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159856)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataIniInicial_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataIniInicial_UpClick

    DataIniInicial.SetFocus

    If Len(Trim(DataIniInicial.ClipText)) > 0 Then

        sData = DataIniInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137614

        DataIniInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataIniInicial_UpClick:

    Select Case gErr

        Case 137614

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159857)

    End Select

    Exit Sub

End Sub

Private Sub DataIniFinal_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataIniFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataIniFinal, iAlterado)
    
End Sub

Private Sub DataIniFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataIniFinal_Validate

    If Len(Trim(DataIniFinal.ClipText)) <> 0 Then

        lErro = Data_Critica(DataIniFinal.Text)
        If lErro <> SUCESSO Then gError 137615

    End If

    Exit Sub

Erro_DataIniFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137615

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159858)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataTerminoFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataTerminoFinal_DownClick

    DataTerminoFinal.SetFocus

    If Len(DataTerminoFinal.ClipText) > 0 Then

        sData = DataTerminoFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137610

        DataTerminoFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataTerminoFinal_DownClick:

    Select Case gErr

        Case 137610

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159859)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataTerminoFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataTerminoFinal_UpClick

    DataTerminoFinal.SetFocus

    If Len(Trim(DataTerminoFinal.ClipText)) > 0 Then

        sData = DataTerminoFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137611

        DataTerminoFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataTerminoFinal_UpClick:

    Select Case gErr

        Case 137611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159860)

    End Select

    Exit Sub

End Sub

Private Sub DataTerminoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataTerminoInicial, iAlterado)
    
End Sub

Private Sub DataTerminoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataTerminoInicial_Validate

    If Len(Trim(DataTerminoInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataTerminoInicial.Text)
        If lErro <> SUCESSO Then gError 137612

    End If

    Exit Sub

Erro_DataTerminoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137612

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159861)

    End Select

    Exit Sub

End Sub

Private Sub DataTerminoInicial_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataTerminoInicial_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataTerminoInicial_DownClick

    DataTerminoInicial.SetFocus

    If Len(DataTerminoInicial.ClipText) > 0 Then

        sData = DataTerminoInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137613

        DataTerminoInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataTerminoInicial_DownClick:

    Select Case gErr

        Case 137613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159862)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataTerminoInicial_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataTerminoInicial_UpClick

    DataTerminoInicial.SetFocus

    If Len(Trim(DataTerminoInicial.ClipText)) > 0 Then

        sData = DataTerminoInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137614

        DataTerminoInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataTerminoInicial_UpClick:

    Select Case gErr

        Case 137614

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159863)

    End Select

    Exit Sub

End Sub

Private Sub DataTerminoFinal_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataTerminoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataTerminoFinal, iAlterado)
    
End Sub

Private Sub DataTerminoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataTerminoFinal_Validate

    If Len(Trim(DataTerminoFinal.ClipText)) <> 0 Then

        lErro = Data_Critica(DataTerminoFinal.Text)
        If lErro <> SUCESSO Then gError 137615

    End If

    Exit Sub

Erro_DataTerminoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137615

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159864)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoFinal, iAlterado)
    
End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 137616
    
    If lErro <> SUCESSO Then gError 137617
  
    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137616
            'erro tratado na rotina chamada

        Case 137617
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159865)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Change()
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoInicial, iAlterado)
    
End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 137618
    
    If lErro <> SUCESSO Then gError 137619

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137618
            'erro tratado na rotina chamada
            
        Case 137619
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159866)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Opcao_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_Selecao Then
    
        'Valida a seleção
        lErro = ValidaSelecao()
        If lErro <> SUCESSO Then gError 136384
    
    End If

    Exit Sub

Erro_Opcao_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 136384

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159867)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Integer

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'Se Frame selecionado foi o de Itens de Projetos
        If TabStrip1.SelectedItem.Index = TAB_ItensProjetos Then
            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
            
                lErro = Trazer_Selecao()
                If lErro <> SUCESSO Then gError 137999
            
            End If
        
        End If
        
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr
    
        Case 137999
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159868)

    End Select

    Exit Sub

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

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0
    
    iFrameAtual = 1
        
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    
    Set objEventoProjetoDe = New AdmEvento
    Set objEventoProjetoAte = New AdmEvento
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 137620

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 137621
    
    lErro = CarregaComboDestino(DestinoSeleciona)
    If lErro <> SUCESSO Then gError 134077
            
    lErro = CarregaComboDestino(Destino)
    If lErro <> SUCESSO Then gError 134077
    
    'Grid Itens
    Set objGridItens = New AdmGrid
    
    'tela em questão
    Set objGridItens.objForm = Me
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 137622
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 137620 To 137622, 136467
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159869)

    End Select

    Exit Sub

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing

    Set objEventoProjetoDe = Nothing
    Set objEventoProjetoAte = Nothing

    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing

    Set objGridItens = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159870)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objProjetoSeleciona As ClassProjetoSeleciona) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objProjetoSeleciona Is Nothing) Then

        ProjetoInicial.Text = objProjetoSeleciona.lProjetoInicial
        Call ProjetoInicial_Validate(bSGECancelDummy)
        
        ProjetoFinal.Text = objProjetoSeleciona.lProjetoFinal
        Call ProjetoFinal_Validate(bSGECancelDummy)
        
        If Len(objProjetoSeleciona.sProdutoInicial) <> 0 Then
        
            ProdutoInicial.Text = objProjetoSeleciona.sProdutoInicial
            Call ProdutoInicial_Validate(bSGECancelDummy)
            
            ProdutoFinal.Text = objProjetoSeleciona.sProdutoFinal
            Call ProdutoFinal_Validate(bSGECancelDummy)
        
        End If
        
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159871)

    End Select
    
    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridItens, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridItens, iAlterado)
        End If

End Sub

Private Sub GridItens_GotFocus()
    
    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()
    
    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

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
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Destino_Col

                    lErro = Saida_Celula_Destino(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 137623

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 137623
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159872)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Destino(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Destino do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim DestinoAnterior As String

On Error GoTo Erro_Saida_Celula_Destino

    Set objGridInt.objControle = Destino
    
    DestinoAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_Destino_Col)

    If DestinoAnterior <> Destino.Text Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Destino_Col) = Destino.Text
        
        'Se o campo foi preenchido
        If Len(Trim(Destino.Text)) > 0 Then
        
            'Marca na tela o item em questão
            GridItens.TextMatrix(GridItens.Row, iGrid_Exportar_Col) = MARCADO
    
            'Atualiza na tela a checkbox marcada
            Call Grid_Refresh_Checkbox(objGridItens)
        
        End If
        
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Destino_Col)) = 0 Then
        
            'Desmarca na tela o item em questão
            GridItens.TextMatrix(GridItens.Row, iGrid_Exportar_Col) = DESMARCADO
    
            'Atualiza na tela a checkbox marcada
            Call Grid_Refresh_Checkbox(objGridItens)
        
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_Destino = SUCESSO

    Exit Function

Erro_Saida_Celula_Destino:

    Saida_Celula_Destino = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159873)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Exportar")
    objGrid.colColuna.Add ("Destino")
    objGrid.colColuna.Add ("Projeto")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Data Início")
    objGrid.colColuna.Add ("Data Término")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Exportar.Name)
    objGrid.colCampo.Add (Destino.Name)
    objGrid.colCampo.Add (Projeto.Name)
    objGrid.colCampo.Add (ProdutoItens.Name)
    objGrid.colCampo.Add (VersaoProdItens.Name)
    objGrid.colCampo.Add (DescricaoProdItens.Name)
    objGrid.colCampo.Add (UMProdItens.Name)
    objGrid.colCampo.Add (QuantidadeProdItens.Name)
    objGrid.colCampo.Add (DataInicio.Name)
    objGrid.colCampo.Add (DataTermino.Name)

    'Colunas do Grid
    iGrid_Exportar_Col = 1
    iGrid_Destino_Col = 2
    iGrid_Projeto_Col = 3
    iGrid_ProdutoItens_Col = 4
    iGrid_VersaoProdItens_Col = 5
    iGrid_DescricaoProdItens_Col = 6
    iGrid_UMProdItens_Col = 7
    iGrid_QuantidadeProdItens_Col = 8
    iGrid_DataInicio_Col = 9
    iGrid_DataTermino_Col = 10

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
                
    If objControl.Name = "Exportar" Then
        
        If objGridItens.iLinhasExistentes = 0 Then
            objControl.Enabled = False
        Else
           objControl.Enabled = True
        End If
        
    ElseIf objControl.Name = "Destino" Then
        
        If objGridItens.iLinhasExistentes = 0 Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If
        
    Else
        objControl.Enabled = False
    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159874)

    End Select

    Exit Sub

End Sub

Function ValidaSelecao() As Long

Dim objClienteInicial As ClassCliente
Dim objClienteFinal As ClassCliente
Dim objProjetoInicial As ClassProjeto
Dim objProjetoFinal As ClassProjeto
Dim sProd_I As String
Dim sProd_F As String
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_ValidaSelecao

    'Valida Clientes
    If Len(Trim(ClienteInicial.Text)) <> 0 And Len(Trim(ClienteFinal.Text)) <> 0 Then
    
        Set objClienteInicial = New ClassCliente
    
        objClienteInicial.sNomeReduzido = Trim(ClienteInicial.Text)
        
        'Lê Cliente Inicial pelo NomeReduzido
        lErro = CF("Cliente_Le_NomeReduzido", objClienteInicial)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 137624
                
        Set objClienteFinal = New ClassCliente
        
        objClienteFinal.sNomeReduzido = Trim(ClienteFinal.Text)
        
        'Lê Cliente Final pelo NomeReduzido
        lErro = CF("Cliente_Le_NomeReduzido", objClienteFinal)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 137625
                
        'codigo do cliente inicial não pode ser maior que o final
        If objClienteInicial.lCodigo > objClienteFinal.lCodigo Then gError 137626
        
    End If
    
    'Valida Produtos
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 137628
    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 137629
    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 137630
    
    End If
    
    'Valida Projeto
    If Len(Trim(ProjetoInicial.Text)) <> 0 And Len(Trim(ProjetoFinal.Text)) <> 0 Then
    
        Set objProjetoInicial = New ClassProjeto
            
        'Verifica sua existencia
        lErro = CF("TP_Projeto_Le", ProjetoInicial, objProjetoInicial)
        If lErro <> SUCESSO Then gError 134467
                
        Set objProjetoFinal = New ClassProjeto
        
        'Verifica sua existencia
        lErro = CF("TP_Projeto_Le", ProjetoFinal, objProjetoFinal)
        If lErro <> SUCESSO Then gError 134467
                
        'codigo do Projeto inicial não pode ser maior que o final
        If objProjetoInicial.lCodigo > objProjetoFinal.lCodigo Then gError 137998
        
    End If
        
    'Valida Data Inicio
    'data inicio não pode ser maior que a final
    If Len(Trim(DataIniInicial.Text)) <> 0 And Len(Trim(DataIniFinal.Text)) <> 0 Then
        
        If StrParaDate(DataIniInicial.Text) > StrParaDate(DataIniFinal.Text) Then gError 137631
    
    End If
    
    'Valida Data Termino
    'data inicio não pode ser maior que a final
    If Len(Trim(DataTerminoInicial.Text)) <> 0 And Len(Trim(DataTerminoFinal.Text)) <> 0 Then
        
        If StrParaDate(DataTerminoInicial.Text) > StrParaDate(DataTerminoFinal.Text) Then gError 137632
    
    End If
    
    'Valida Data Inicio X Termino
    'data inicio não pode ser maior que a final
    If Len(Trim(DataIniFinal.Text)) <> 0 And Len(Trim(DataTerminoInicial.Text)) <> 0 Then
        
        If StrParaDate(DataIniFinal.Text) > StrParaDate(DataTerminoInicial.Text) Then gError 137633
    
    ElseIf Len(Trim(DataIniInicial.Text)) <> 0 And Len(Trim(DataTerminoInicial.Text)) <> 0 Then
        
        If StrParaDate(DataIniInicial.Text) > StrParaDate(DataTerminoInicial.Text) Then gError 137634
    
    ElseIf Len(Trim(DataIniFinal.Text)) <> 0 And Len(Trim(DataTerminoFinal.Text)) <> 0 Then
        
        If StrParaDate(DataIniFinal.Text) > StrParaDate(DataTerminoFinal.Text) Then gError 137635
    
    ElseIf Len(Trim(DataIniInicial.Text)) <> 0 And Len(Trim(DataTerminoFinal.Text)) <> 0 Then
        
        If StrParaDate(DataIniInicial.Text) > StrParaDate(DataTerminoFinal.Text) Then gError 137636
    
    End If
        
    ValidaSelecao = SUCESSO
    
    Exit Function
    
Erro_ValidaSelecao:

    ValidaSelecao = gErr

    Select Case gErr
    
        Case 137624, 137625
            'erros tratados nas rotinas chamadas
                
        Case 137626
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
                
        Case 137628
            ProdutoInicial.SetFocus

        Case 137629
            ProdutoFinal.SetFocus

        Case 137630
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
                        
        Case 137631
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_INICIAL_MAIOR", gErr)
            
        Case 137632
            Call Rotina_Erro(vbOKOnly, "ERRO_DATATERMINO_INICIAL_MAIOR", gErr)
        
        Case 137633
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_FIM_MAIOR_DATATERM_INI", gErr)
            
        Case 137634
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_INI_MAIOR_DATATERM_INI", gErr)
            
        Case 137635
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_FIM_MAIOR_DATATERM_FIM", gErr)
            
        Case 137636
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_INI_MAIOR_DATATERM_FIM", gErr)
        
        Case 137998
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETO_INICIAL_MAIOR", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159875)
    
    End Select

    Exit Function

End Function

'por Jorge Specian - chamada por Cliente_Change para localizar pela parte digitada do Nome
'Reduzido do Cliente através da CF Cliente_Pesquisa_NomeReduzido em RotinasCRFAT.ClassCRFATSelect
Private Sub Cliente_Preenche(objControle As Object)

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
    
On Error GoTo Erro_Cliente_Preenche
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objControle, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 137632

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 137632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159876)

    End Select
    
    Exit Sub

End Sub

Function Limpa_Tela_ExportarProjetos() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ExportarProjetos
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    DescClienteFinal.Caption = ""
    DescClienteInicial.Caption = ""
    DescProdFim.Caption = ""
    DescProdInic.Caption = ""
    DescProjetoFinal.Caption = ""
    DescProjetoInicio.Caption = ""
    
    DestinoSeleciona.ListIndex = 0
    
    RelatorioAutomatico.Value = vbUnchecked
    
    Call Grid_Limpa(objGridItens)
    
    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = 1
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    iAlterado = 0
    iTabPrincipalAlterado = REGISTRO_ALTERADO

    Limpa_Tela_ExportarProjetos = SUCESSO

    Exit Function

Erro_Limpa_Tela_ExportarProjetos:

    Limpa_Tela_ExportarProjetos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159877)

    End Select

    Exit Function

End Function

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_ExportarProjetos
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 136396

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159878)

    End Select

    Exit Sub

End Sub

Function Move_TabSelecao_Memoria(ByVal objProjetoSeleciona As ClassProjetoSeleciona) As Long

Dim lErro As Long
Dim objCliente As ClassCliente
Dim sProduto As String
Dim iProdPreenchido As Integer
Dim objProjeto As ClassProjeto

On Error GoTo Erro_Move_TabSelecao_Memoria

    Set objCliente = New ClassCliente

    objCliente.sNomeReduzido = Trim(ClienteInicial.Text)
    
    'Lê Cliente Inicial pelo NomeReduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 136399
    
    objProjetoSeleciona.lClienteInicial = objCliente.lCodigo
            
    Set objCliente = New ClassCliente

    objCliente.sNomeReduzido = Trim(ClienteFinal.Text)
    
    'Lê Cliente Final pelo NomeReduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 136400
    
    objProjetoSeleciona.lClienteFinal = objCliente.lCodigo
    
    Set objProjeto = New ClassProjeto

    objProjeto.sNomeReduzido = Trim(ProjetoInicial.Text)
    
    'Lê Projeto Inicial pelo NomeReduzido
    lErro = CF("Projeto_Le_NomeReduzido", objProjeto)
    If lErro <> SUCESSO And lErro <> 139161 Then gError 136399
    
    objProjetoSeleciona.lProjetoInicial = objProjeto.lCodigo
            
    Set objProjeto = New ClassProjeto

    objProjeto.sNomeReduzido = Trim(ProjetoFinal.Text)
    
    'Lê Projeto Final pelo NomeReduzido
    lErro = CF("Projeto_Le_NomeReduzido", objProjeto)
    If lErro <> SUCESSO And lErro <> 139161 Then gError 136400
    
    objProjetoSeleciona.lProjetoFinal = objProjeto.lCodigo
    
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProduto, iProdPreenchido)
    If lErro <> SUCESSO Then gError 136401
         
    objProjetoSeleciona.sProdutoInicial = sProduto
    
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProduto, iProdPreenchido)
    If lErro <> SUCESSO Then gError 136402
         
    objProjetoSeleciona.sProdutoFinal = sProduto
    
    objProjetoSeleciona.dtDataIniFinal = StrParaDate(DataIniFinal.Text)
    objProjetoSeleciona.dtDataIniInicial = StrParaDate(DataIniInicial.Text)
    objProjetoSeleciona.dtDataTerminoFinal = StrParaDate(DataTerminoFinal.Text)
    objProjetoSeleciona.dtDataTerminoInicial = StrParaDate(DataTerminoInicial.Text)
    
    If Len(DestinoSeleciona.Text) <> 0 Then
        objProjetoSeleciona.iDestino = Codigo_Extrai(DestinoSeleciona.Text)
    End If
         
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr
    
        Case 136399 To 136402

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159879)

    End Select

    Exit Function

End Function

Private Function Trazer_Selecao() As Long

Dim lErro As Long
Dim objProjetoSeleciona As New ClassProjetoSeleciona

On Error GoTo Erro_Trazer_Selecao

    GL_objMDIForm.MousePointer = vbHourglass

    'Limpa o Grid antes de tudo
    Call Grid_Limpa(objGridItens)

    'Move a seleção para a memória
    lErro = Move_TabSelecao_Memoria(objProjetoSeleciona)
    If lErro <> SUCESSO Then gError 136385

    'Le os itens de projeto que podem ser exportados
    lErro = CF("Projeto_Le_ItensExportaveis", objProjetoSeleciona)
    If lErro <> SUCESSO And lErro <> 139131 Then gError 136386
    
    If lErro = SUCESSO Then
        
        'Preenche GridItens com os dados retornados
        lErro = Preenche_GridItens(objProjetoSeleciona)
        If lErro <> SUCESSO Then gError 136387
    
    End If
    
    iTabPrincipalAlterado = 0
        
    GL_objMDIForm.MousePointer = vbDefault
        
    Trazer_Selecao = SUCESSO
            
    Exit Function

Erro_Trazer_Selecao:

    GL_objMDIForm.MousePointer = vbDefault

    Trazer_Selecao = gErr
    
    Select Case gErr
    
        Case 136385, 136386, 136387
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159880)

    End Select
    
End Function

Private Function CarregaComboDestino(objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CarregaComboDestino

    objCombo.AddItem ""
    objCombo.ItemData(objCombo.NewIndex) = 0
    
    objCombo.AddItem ITEMDEST_ORCAMENTO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_ORCAMENTO_DE_VENDA
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_ORCAMENTO_DE_VENDA
    
    objCombo.AddItem ITEMDEST_PEDIDO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_PEDIDO_DE_VENDA
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_PEDIDO_DE_VENDA
    
    objCombo.AddItem ITEMDEST_ORDEM_DE_PRODUCAO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_PRODUCAO
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_ORDEM_DE_PRODUCAO
    
    objCombo.AddItem ITEMDEST_NFISCAL_SIMPLES & SEPARADOR & STRING_ITEMDEST_NFISCAL_SIMPLES
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_NFISCAL_SIMPLES
    
    objCombo.AddItem ITEMDEST_NFISCAL_FATURA & SEPARADOR & STRING_ITEMDEST_NFISCAL_FATURA
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_NFISCAL_FATURA
    
    objCombo.AddItem ITEMDEST_ORDEM_DE_SERVICO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_SERVICO
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_ORDEM_DE_SERVICO
        
    CarregaComboDestino = SUCESSO

    Exit Function

Erro_CarregaComboDestino:

    CarregaComboDestino = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159881)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os Itens do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'Desmarca na tela o Item em questão
        GridItens.TextMatrix(iLinha, iGrid_Exportar_Col) = DESMARCADO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridItens)

End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os itens do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'Marca na tela o item em questão
        GridItens.TextMatrix(iLinha, iGrid_Exportar_Col) = MARCADO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridItens)

End Sub

Private Sub ProdutoItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ProdutoItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ProdutoItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ProdutoItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescricaoProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescricaoProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VersaoProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VersaoProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub VersaoProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub VersaoProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = VersaoProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UMProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UMProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UMProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantidadeProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub QuantidadeProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantidadeProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub DataInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataInicio
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataTermino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataTermino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataTermino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataTermino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataTermino
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub





Private Sub Exportar_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Exportar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Exportar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Exportar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Exportar
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Destino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Destino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Destino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Destino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Destino
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Preenche_GridItens(objProjetoSeleciona As ClassProjetoSeleciona) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objProjeto As New ClassProjeto
Dim objProjetoItens As New ClassProjetoItens
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_Preenche_GridItens

    iLinha = 0
    
    'para cada projeto na coleção passada ...
    For Each objProjeto In objProjetoSeleciona.colProjetos
    
        'Lê o Projeto
        lErro = CF("Projeto_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
    
        'para cada item do projeto encontrado ...
        For Each objProjetoItens In objProjeto.colProjetoItens
        
            Set objProdutos = New ClassProduto
            
            objProdutos.sCodigo = objProjetoItens.sProduto
            
            'Le o Produto para pegar sua Descrição
            lErro = CF("Produto_Le", objProdutos)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 134768
            
            'Mascara o Código do Produto para por na Tela
            lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 134100
            
            iLinha = iLinha + 1
            
            'Preenche as colunas do Grid
            GridItens.TextMatrix(iLinha, iGrid_Destino_Col) = Seleciona_Destino(objProjetoItens.iDestino)
            GridItens.TextMatrix(iLinha, iGrid_Projeto_Col) = objProjeto.sNomeReduzido
            GridItens.TextMatrix(iLinha, iGrid_ProdutoItens_Col) = sProdutoMascarado
            GridItens.TextMatrix(iLinha, iGrid_DescricaoProdItens_Col) = objProdutos.sDescricao
            GridItens.TextMatrix(iLinha, iGrid_VersaoProdItens_Col) = objProjetoItens.sVersao
            GridItens.TextMatrix(iLinha, iGrid_UMProdItens_Col) = objProjetoItens.sUMedida
            GridItens.TextMatrix(iLinha, iGrid_QuantidadeProdItens_Col) = Formata_Estoque(objProjetoItens.dQuantidade)
            If objProjetoItens.dtDataInicioPrev <> DATA_NULA Then
                GridItens.TextMatrix(iLinha, iGrid_DataInicio_Col) = Format(objProjetoItens.dtDataInicioPrev, "dd/mm/yyyy")
            End If
            If objProjetoItens.dtDataTerminoPrev <> DATA_NULA Then
                GridItens.TextMatrix(iLinha, iGrid_DataTermino_Col) = Format(objProjetoItens.dtDataTerminoPrev, "dd/mm/yyyy")
            End If
            
        Next
    
    Next
    
    'seta a quantidade de linhas do grid
    objGridItens.iLinhasExistentes = iLinha
    
    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr
    
        Case 134768
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159882)

    End Select

    Exit Function

End Function

Public Function Verifica_Relacionamento(objProjetoItensRegGerados As ClassProjetoItensRegGerados, iDestino As Integer) As Long

Dim lErro As Long
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProjeto As New ClassProjeto
Dim objProjetoItens As New ClassProjetoItens

On Error GoTo Erro_Verifica_Relacionamento

    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134702

    'se o produto não existe cadastrado ...
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 134200
    
    objProjeto.sNomeReduzido = Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Projeto_Col))
    
    'Verifica se o Projeto existe, lendo no BD a partir do Código
    lErro = CF("Projeto_Le_NomeReduzido", objProjeto)
    If lErro <> SUCESSO And lErro <> 139161 Then gError 134094
        
    If lErro <> SUCESSO Then gError 134095
    
    'Le os Itens do Projeto
    lErro = CF("Projeto_Le_Itens", objProjeto)
    If lErro <> SUCESSO And lErro <> 139126 Then gError 134096
    
    'Percorre os Itens do Projeto para achar o Produto e Versão
    For Each objProjetoItens In objProjeto.colProjetoItens
    
        If objProjetoItens.sProduto = sProdutoFormatado And objProjetoItens.sVersao = Trim(sVersao) Then
    
            Exit For
    
        End If
        
    Next
    
    objProjetoItensRegGerados.lNumIntDocItemProj = objProjetoItens.lNumIntDoc
    objProjetoItensRegGerados.iDestino = iDestino
    
    'Le a tabela de relacionamento
    lErro = CF("Projeto_Le_ItensRegGerados", objProjetoItensRegGerados)
    If lErro <> SUCESSO And lErro <> 139157 Then gError 137101
    
    'se não tem relacionamento ... erro
    If lErro <> SUCESSO Then gError 137102
    
    Verifica_Relacionamento = SUCESSO

    Exit Function

Erro_Verifica_Relacionamento:

    Verifica_Relacionamento = gErr

    Select Case gErr
        
        Case 137102 'Não ha relacionamentos
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159883)

    
    End Select

    Exit Function

End Function

Private Function Seleciona_Destino(iDestino As Integer) As String

Dim sTexto As String

    Select Case iDestino
    
        Case Is = ITEMDEST_ORCAMENTO_DE_VENDA
            sTexto = ITEMDEST_ORCAMENTO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_ORCAMENTO_DE_VENDA
    
        Case Is = ITEMDEST_PEDIDO_DE_VENDA
            sTexto = ITEMDEST_PEDIDO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_PEDIDO_DE_VENDA
    
        Case Is = ITEMDEST_ORDEM_DE_PRODUCAO
            sTexto = ITEMDEST_ORDEM_DE_PRODUCAO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_PRODUCAO
    
        Case Is = ITEMDEST_NFISCAL_SIMPLES
            sTexto = ITEMDEST_NFISCAL_SIMPLES & SEPARADOR & STRING_ITEMDEST_NFISCAL_SIMPLES
    
        Case Is = ITEMDEST_NFISCAL_FATURA
            sTexto = ITEMDEST_NFISCAL_FATURA & SEPARADOR & STRING_ITEMDEST_NFISCAL_FATURA
    
        Case Is = ITEMDEST_ORDEM_DE_SERVICO
            sTexto = ITEMDEST_ORDEM_DE_SERVICO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_SERVICO
    
    End Select
    
    Seleciona_Destino = sTexto
    
End Function

Public Function Agrupa_Destinos(colOrcamentoVenda As Collection, colPedidoVenda As Collection, colOrdemProducao As Collection, colNFSimples As Collection, colNFFatura As Collection, colOrdemServico As Collection) As Long
    
Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim iLinha As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProjetoItens As New ClassProjetoItens
Dim iDestino As Integer

On Error GoTo Erro_Agrupa_Destinos

    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        'se marcado para exportar...
        If GridItens.TextMatrix(iLinha, iGrid_Exportar_Col) = MARCADO Then
        
            Set objProjeto = New ClassProjeto
            
            objProjeto.sNomeReduzido = GridItens.TextMatrix(iLinha, iGrid_Projeto_Col)
            
            'Le o Projeto
            lErro = CF("Projeto_Le_NomeReduzido", objProjeto)
            If lErro <> SUCESSO And lErro <> 139161 Then gError 134094
                
            If lErro <> SUCESSO Then gError 134095
            
            'Le os Itens do Projeto
            lErro = CF("Projeto_Le_Itens", objProjeto)
            If lErro <> SUCESSO And lErro <> 139126 Then gError 134096
            
            'Formata o Código do Produto
            lErro = CF("Produto_Formata", Trim(GridItens.TextMatrix(iLinha, iGrid_ProdutoItens_Col)), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 134080
            
            'Percorre os Itens do Projeto para achar o Produto e Versão
            For Each objProjetoItens In objProjeto.colProjetoItens
            
                If objProjetoItens.sProduto = sProdutoFormatado And objProjetoItens.sVersao = Trim(GridItens.TextMatrix(iLinha, iGrid_VersaoProdItens_Col)) Then
                    Exit For
                End If
                
            Next
            
            iDestino = Codigo_Extrai(GridItens.TextMatrix(iLinha, iGrid_Destino_Col))
            
            'altera o destino para o utilizado no grid
            objProjetoItens.iDestino = iDestino
            
            'Classifica e agrupa em coleções por Destino a Exportar
            Select Case iDestino
            
                Case Is = ITEMDEST_ORCAMENTO_DE_VENDA
                
                    colOrcamentoVenda.Add objProjetoItens
                
                Case Is = ITEMDEST_PEDIDO_DE_VENDA
                
                    colPedidoVenda.Add objProjetoItens
                
                Case Is = ITEMDEST_ORDEM_DE_PRODUCAO
                
                    colOrdemProducao.Add objProjetoItens
                
                Case Is = ITEMDEST_NFISCAL_SIMPLES
                
                    colNFSimples.Add objProjetoItens
                
                Case Is = ITEMDEST_NFISCAL_FATURA
                
                    colNFFatura.Add objProjetoItens
                
                Case Is = ITEMDEST_ORDEM_DE_SERVICO
                
                    colOrdemServico.Add objProjetoItens
                
            End Select
            
        End If
    
    Next

    Agrupa_Destinos = SUCESSO

    Exit Function

Erro_Agrupa_Destinos:

    Agrupa_Destinos = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159884)

    End Select
    
    Exit Function
    
End Function

Public Function Exporta_OrcamentoVenda(colOrcamentoVenda As Collection) As Long

Dim lErro As Long
Dim lNumIntDocProj As Long
Dim objProjetoItens As New ClassProjetoItens
Dim colOrcamentoVenda2 As Collection
Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_Exporta_OrcamentoVenda

    'Inicializar com o primeiro projeto
    lNumIntDocProj = colOrcamentoVenda.Item(1).lNumIntDocProj
    Set colOrcamentoVenda2 = New Collection
    
    'para cada item de projeto da coleção passada
    For Each objProjetoItens In colOrcamentoVenda
    
        'verificar se é do mesmo projeto
        If lNumIntDocProj = objProjetoItens.lNumIntDocProj Then
        
            'sendo do mesmo projeto inclui na segunda coleção
            colOrcamentoVenda2.Add objProjetoItens
            
        Else
            
            'sendo de outro projeto... gerar o anterior
            Set objOrcamentoVenda = New ClassOrcamentoVenda
            
            'Gera o OrcamentoVenda a partir dos Itens do Projeto
            lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda, colOrcamentoVenda2)
            If lErro <> SUCESSO Then gError 134200
            
            'Grava OrcamentoVenda Gerado no BD
            lErro = CF("Projeto_Grava_OrcamentoVenda", objOrcamentoVenda, colOrcamentoVenda2)
            If lErro <> SUCESSO Then gError 134201
            
            'reinicializa com o novo projeto
            lNumIntDocProj = objProjetoItens.lNumIntDocProj
            Set colOrcamentoVenda2 = New Collection
            
            'inclui na nova coleção
            colOrcamentoVenda2.Add objProjetoItens
            
        End If
    
    Next
    
    'sendo o último projeto ou projeto único... gerar este
    Set objOrcamentoVenda = New ClassOrcamentoVenda
    
    'Gera o OrcamentoVenda a partir dos Itens do Projeto
    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda, colOrcamentoVenda2)
    If lErro <> SUCESSO Then gError 134202
    
    'Grava OrcamentoVenda Gerado no BD
    lErro = CF("Projeto_Grava_OrcamentoVenda", objOrcamentoVenda, colOrcamentoVenda2)
    If lErro <> SUCESSO Then gError 134203
        
    Exporta_OrcamentoVenda = SUCESSO

    Exit Function

Erro_Exporta_OrcamentoVenda:

    Exporta_OrcamentoVenda = gErr

    Select Case gErr
    
        Case 134200 To 134203
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159885)

    End Select
    
    Exit Function
    
End Function

Private Function Move_OrcamentoVenda_Memoria(objOrcamentoVenda As ClassOrcamentoVenda, colOrcamentoVenda As Collection) As Long

Dim lErro As Long
Dim lOrcamentoVenda As Long
Dim objProjeto As ClassProjeto
Dim objCliente As ClassCliente
Dim objTipoCliente As New ClassTipoCliente
Dim objFilialCliente As ClassFilialCliente
Dim dDesconto As Double
Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim objProjetoItens As New ClassProjetoItens
Dim objItemOrcamento As ClassItemOV
Dim objProdutos As ClassProduto
Dim iCont As Integer
Dim objCondicaoPagto As ClassCondicaoPagto
Dim iIndice As Integer
Dim objParcelaOV As ClassParcelaOV

On Error GoTo Erro_Move_OrcamentoVenda_Memoria

    Set objProjeto = New ClassProjeto
    
    objProjeto.lNumIntDoc = colOrcamentoVenda.Item(1).lNumIntDocProj
    
    'Le o Projeto
    lErro = CF("Projeto_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
    
    Set objCliente = New ClassCliente

    objCliente.lCodigo = objProjeto.lCodCliente

    'Busca o Cliente no BD
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 84037
                        
    'se encontrou...
    If lErro = SUCESSO Then
    
        objOrcamentoVenda.lCliente = objCliente.lCodigo
        objOrcamentoVenda.sNomeCli = objCliente.sNomeReduzido
            
        'Se o Tipo estiver preenchido
        If objCliente.iTipo > 0 Then
            objTipoCliente.iCodigo = objCliente.iTipo
            'Lê o Tipo de Cliente
            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 126956
        End If

        'Guarda o Desconto do cliente
        If objCliente.dDesconto > 0 Then
            dDesconto = objCliente.dDesconto
        ElseIf objTipoCliente.dDesconto > 0 Then
            dDesconto = objTipoCliente.dDesconto
        Else
            dDesconto = 0
        End If

        'Preenche Vendedor
        If objCliente.iVendedor > 0 Then
            objOrcamentoVenda.iVendedor = objCliente.iVendedor
        ElseIf objTipoCliente.iVendedor > 0 Then
            objOrcamentoVenda.iVendedor = objTipoCliente.iVendedor
        End If

        'Preenche Tabela de Preço
        If objCliente.iTabelaPreco > 0 Then
            objOrcamentoVenda.iTabelaPreco = objCliente.iTabelaPreco
        ElseIf objTipoCliente.iTabelaPreco > 0 Then
            objOrcamentoVenda.iTabelaPreco = objTipoCliente.iTabelaPreco
        End If
        
        'Preenche Condição de Pagamento
        If objCliente.iCondicaoPagto > 0 Then
            objOrcamentoVenda.iCondicaoPagto = objCliente.iCondicaoPagto
        ElseIf objTipoCliente.iCondicaoPagto > 0 Then
            objOrcamentoVenda.iCondicaoPagto = objTipoCliente.iCondicaoPagto
        End If
  
    End If
                       
    objOrcamentoVenda.iFilial = objProjeto.iCodFilial
    
    Set objFilialCliente = New ClassFilialCliente
    
    objFilialCliente.lCodCliente = objProjeto.lCodCliente
    objFilialCliente.iCodFilial = objProjeto.iCodFilial
    
    'Busca a Filial do Cliente no BD
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 84037

    If lErro = SUCESSO Then
        
        objOrcamentoVenda.sNomeFilialCli = objFilialCliente.sNome
        
        If objFilialCliente.iVendedor <> 0 Then
            If objOrcamentoVenda.iVendedor <> objFilialCliente.iVendedor Then
                objOrcamentoVenda.iVendedor = objFilialCliente.iVendedor
            End If
        End If
        
    End If
            
    'Preenche outros dados do OV
    objOrcamentoVenda.dtDataEmissao = gdtDataAtual
    objOrcamentoVenda.dtDataReferencia = gdtDataAtual
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    objOrcamentoVenda.sUsuario = gsUsuario
    objOrcamentoVenda.sNaturezaOp = "5101"
    
    'Obtem o Código do Novo Orcamento de Venda
    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_CODIGO_ORCAMENTOVENDA", "OrcamentoVenda", "Codigo", lOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94422
    
    objOrcamentoVenda.lCodigo = lOrcamentoVenda
    
    dValorProdutos = 0
    iCont = 0
    
    'Para cada OV a ser gerado
    For Each objProjetoItens In colOrcamentoVenda
    
        iCont = iCont + 1
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objProjetoItens.sProduto
        
        'Le o Produto para pegar sua Descrição
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134768
    
        Set objItemOrcamento = New ClassItemOV
        
        'Coloca os dados no Item do Orçamento
        objItemOrcamento.sProduto = objProjetoItens.sProduto
        objItemOrcamento.sDescricao = objProdutos.sDescricao
        objItemOrcamento.sUnidadeMed = objProjetoItens.sUMedida
        objItemOrcamento.dQuantidade = objProjetoItens.dQuantidade
        objItemOrcamento.dPrecoUnitario = objProjetoItens.dPrecoTotalItem
        objItemOrcamento.dPrecoTotal = objProjetoItens.dQuantidade * objProjetoItens.dPrecoTotalItem
        objItemOrcamento.iFilialEmpresa = giFilialEmpresa
        objItemOrcamento.dtDataEntrega = objOrcamentoVenda.dtDataEmissao
        
        'Se controla versão do Kit
        If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
            objItemOrcamento.sVersaoKit = lOrcamentoVenda & SEPARADOR & iCont
            objItemOrcamento.sVersaoKitBase = objProjetoItens.sVersao
        End If
        
        dValorProdutos = dValorProdutos + objProjetoItens.dQuantidade * objProjetoItens.dPrecoTotalItem
        
        'Adiciona o item na colecao de itens do orçamento de venda
        objOrcamentoVenda.colItens.Add objItemOrcamento
    
    Next
    
    objOrcamentoVenda.dValorProdutos = dValorProdutos
    
    'Se o cliente possui desconto
    If dDesconto > 0 Then

        'Calcula o valor do desconto para o cliente
        dValorDesconto = dDesconto * dValorProdutos

        'Para tributação
        objOrcamentoVenda.dValorDesconto = dValorDesconto

    End If
    
    objOrcamentoVenda.dValorTotal = dValorProdutos - dValorDesconto
    
    'Trata a Condição de Pagamento X Parcelas
    Set objCondicaoPagto = New ClassCondicaoPagto
    
    objCondicaoPagto.iCodigo = objOrcamentoVenda.iCondicaoPagto
    
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 84046
    
    'Se encontrou a Condição de Pagamento...
    If lErro <> SUCESSO Then
            
        objCondicaoPagto.dValorTotal = objOrcamentoVenda.dValorTotal
        objCondicaoPagto.dtDataRef = objOrcamentoVenda.dtDataReferencia

        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True)
        If lErro <> SUCESSO Then gError 84076

        'Coloca os valores das Parcelas na coleção de Parcelas
        For iIndice = 1 To objCondicaoPagto.colParcelas.Count
            
            Set objParcelaOV = New ClassParcelaOV

            objParcelaOV.iNumParcela = iIndice
            objParcelaOV.dValor = objCondicaoPagto.colParcelas(iIndice).dValor
            objParcelaOV.dtDataVencimento = objCondicaoPagto.colParcelas(iIndice).dtVencimento
            objParcelaOV.dtDesconto1Ate = DATA_NULA
            objParcelaOV.dtDesconto2Ate = DATA_NULA
            objParcelaOV.dtDesconto3Ate = DATA_NULA
            
            'Inclui na coleção
            objOrcamentoVenda.colParcela.Add objParcelaOV

        Next

    Else
        
        'Gera Parcela Única
        Set objParcelaOV = New ClassParcelaOV
    
        objParcelaOV.iNumParcela = 1
        objParcelaOV.dtDataVencimento = gdtDataAtual
        objParcelaOV.dValor = objOrcamentoVenda.dValorTotal
        objParcelaOV.dtDesconto1Ate = DATA_NULA
        objParcelaOV.dtDesconto2Ate = DATA_NULA
        objParcelaOV.dtDesconto3Ate = DATA_NULA
        
        'Inclui na coleção
        objOrcamentoVenda.colParcela.Add objParcelaOV
    
    End If
        
    Move_OrcamentoVenda_Memoria = SUCESSO

    Exit Function

Erro_Move_OrcamentoVenda_Memoria:

    Move_OrcamentoVenda_Memoria = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159886)

    End Select

    Exit Function

End Function

Public Function Exporta_PedidoVenda(colPedidoVenda As Collection) As Long

Dim lErro As Long
Dim lNumIntDocProj As Long
Dim objProjetoItens As New ClassProjetoItens
Dim colPedidoDeVenda2 As Collection
Dim objPedidoDeVenda As ClassPedidoDeVenda

On Error GoTo Erro_Exporta_PedidoVenda

    'Inicializar com o primeiro projeto
    lNumIntDocProj = colPedidoVenda.Item(1).lNumIntDocProj
    Set colPedidoDeVenda2 = New Collection
    
    'para cada item de projeto da coleção passada
    For Each objProjetoItens In colPedidoVenda
    
        'verificar se é do mesmo projeto
        If lNumIntDocProj = objProjetoItens.lNumIntDocProj Then
        
            'sendo do mesmo projeto inclui na segunda coleção
            colPedidoDeVenda2.Add objProjetoItens
            
        Else
            
            'sendo de outro projeto... gerar o anterior
            Set objPedidoDeVenda = New ClassPedidoDeVenda
            
            'Gera o PedidoDeVenda a partir dos Itens do Projeto
            lErro = Move_PedidoDeVenda_Memoria(objPedidoDeVenda, colPedidoDeVenda2)
            If lErro <> SUCESSO Then gError 134200
            
            'Grava PedidoDeVenda Gerado no BD
            lErro = CF("Projeto_Grava_PedidoDeVenda", objPedidoDeVenda, colPedidoDeVenda2)
            If lErro <> SUCESSO Then gError 134201
            
            'reinicializa com o novo projeto
            lNumIntDocProj = objProjetoItens.lNumIntDocProj
            Set colPedidoDeVenda2 = New Collection
            
            'inclui na nova coleção
            colPedidoDeVenda2.Add objProjetoItens
            
        End If
    
    Next
    
    'sendo o último projeto ou projeto único... gerar este
    Set objPedidoDeVenda = New ClassPedidoDeVenda
    
    'Gera o PedidoDeVenda a partir dos Itens do Projeto
    lErro = Move_PedidoDeVenda_Memoria(objPedidoDeVenda, colPedidoDeVenda2)
    If lErro <> SUCESSO Then gError 134202
    
    'Grava PedidoDeVenda Gerado no BD
    lErro = CF("Projeto_Grava_PedidoDeVenda", objPedidoDeVenda, colPedidoDeVenda2)
    If lErro <> SUCESSO Then gError 134203
        
    Exporta_PedidoVenda = SUCESSO

    Exit Function

Erro_Exporta_PedidoVenda:

    Exporta_PedidoVenda = gErr

    Select Case gErr
    
        Case 134200 To 134203
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159887)

    End Select
    
    Exit Function
    
End Function

Public Function Exporta_OrdemDeProducao(colOrdemDeProducao As Collection) As Long

Dim lErro As Long
Dim lNumIntDocProj As Long
Dim objProjetoItens As New ClassProjetoItens
Dim colOrdemDeProducao2 As Collection
Dim objOrdemDeProducao As ClassOrdemDeProducao

On Error GoTo Erro_Exporta_OrdemDeProducao

    'Inicializar com o primeiro projeto
    lNumIntDocProj = colOrdemDeProducao.Item(1).lNumIntDocProj
    Set colOrdemDeProducao2 = New Collection
    
    'para cada item de projeto da coleção passada
    For Each objProjetoItens In colOrdemDeProducao
    
        'verificar se é do mesmo projeto
        If lNumIntDocProj = objProjetoItens.lNumIntDocProj Then
        
            'sendo do mesmo projeto inclui na segunda coleção
            colOrdemDeProducao2.Add objProjetoItens
            
        Else
            
            'sendo de outro projeto... gerar o anterior
            Set objOrdemDeProducao = New ClassOrdemDeProducao
            
            'Gera a OrdemDeProducao a partir dos Itens do Projeto
            lErro = Move_OrdemDeProducao_Memoria(objOrdemDeProducao, colOrdemDeProducao2)
            If lErro <> SUCESSO Then gError 134200
            
            'Grava OrdemDeProducao Gerado no BD
'            lErro = CF("Projeto_Grava_OrdemDeProducao", objOrdemDeProducao, colOrdemDeProducao2)
            If lErro <> SUCESSO Then gError 134201
            
            'reinicializa com o novo projeto
            lNumIntDocProj = objProjetoItens.lNumIntDocProj
            Set colOrdemDeProducao2 = New Collection
            
            'inclui na nova coleção
            colOrdemDeProducao2.Add objProjetoItens
            
        End If
    
    Next
    
    'sendo o último projeto ou projeto único... gerar este
    Set objOrdemDeProducao = New ClassOrdemDeProducao
    
    'Gera a OrdemDeProducao a partir dos Itens do Projeto
    lErro = Move_OrdemDeProducao_Memoria(objOrdemDeProducao, colOrdemDeProducao2)
    If lErro <> SUCESSO Then gError 134202
    
    'Grava OrdemDeProducao Gerado no BD
'    lErro = CF("Projeto_Grava_OrdemDeProducao", objOrdemDeProducao, colOrdemDeProducao2)
    If lErro <> SUCESSO Then gError 134203
        
    Exporta_OrdemDeProducao = SUCESSO

    Exit Function

Erro_Exporta_OrdemDeProducao:

    Exporta_OrdemDeProducao = gErr

    Select Case gErr
    
        Case 134200 To 134203
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159888)

    End Select
    
    Exit Function
    
End Function

Public Function Exporta_NFSimples(colNFSimples As Collection) As Long

Dim lErro As Long
Dim lNumIntDocProj As Long
Dim objProjetoItens As New ClassProjetoItens
Dim colNFSimples2 As Collection
Dim objNFSimples As ClassNFiscal

On Error GoTo Erro_Exporta_NFSimples

    'Inicializar com o primeiro projeto
    lNumIntDocProj = colNFSimples.Item(1).lNumIntDocProj
    Set colNFSimples2 = New Collection
    
    'para cada item de projeto da coleção passada
    For Each objProjetoItens In colNFSimples
    
        'verificar se é do mesmo projeto
        If lNumIntDocProj = objProjetoItens.lNumIntDocProj Then
        
            'sendo do mesmo projeto inclui na segunda coleção
            colNFSimples2.Add objProjetoItens
            
        Else
            
            'sendo de outro projeto... gerar o anterior
            Set objNFSimples = New ClassNFiscal
            
            'Gera a NFSimples a partir dos Itens do Projeto
'            lErro = Move_NFSimples_Memoria(objNFSimples, colNFSimples2)
            If lErro <> SUCESSO Then gError 134200
            
            'Grava NFSimples Gerado no BD
'            lErro = CF("Projeto_Grava_NFSimples", objNFSimples, colNFSimples2)
            If lErro <> SUCESSO Then gError 134201
            
            'reinicializa com o novo projeto
            lNumIntDocProj = objProjetoItens.lNumIntDocProj
            Set colNFSimples2 = New Collection
            
            'inclui na nova coleção
            colNFSimples2.Add objProjetoItens
            
        End If
    
    Next
    
    'sendo o último projeto ou projeto único... gerar este
    Set objNFSimples = New ClassNFiscal
    
    'Gera a NFSimples a partir dos Itens do Projeto
'    lErro = Move_NFSimples_Memoria(objNFSimples, colNFSimples2)
    If lErro <> SUCESSO Then gError 134202
    
    'Grava NFSimples Gerado no BD
'    lErro = CF("Projeto_Grava_NFSimples", objNFSimples, colNFSimples2)
    If lErro <> SUCESSO Then gError 134203
        
    Exporta_NFSimples = SUCESSO

    Exit Function

Erro_Exporta_NFSimples:

    Exporta_NFSimples = gErr

    Select Case gErr
    
        Case 134200 To 134203
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159889)

    End Select
    
    Exit Function
    
End Function

Public Function Exporta_NFFatura(colNFFatura As Collection) As Long

Dim lErro As Long
Dim lNumIntDocProj As Long
Dim objProjetoItens As New ClassProjetoItens
Dim colNFFatura2 As Collection
Dim objNFFatura As ClassNFiscal

On Error GoTo Erro_Exporta_NFFatura

    'Inicializar com o primeiro projeto
    lNumIntDocProj = colNFFatura.Item(1).lNumIntDocProj
    Set colNFFatura2 = New Collection
    
    'para cada item de projeto da coleção passada
    For Each objProjetoItens In colNFFatura
    
        'verificar se é do mesmo projeto
        If lNumIntDocProj = objProjetoItens.lNumIntDocProj Then
        
            'sendo do mesmo projeto inclui na segunda coleção
            colNFFatura2.Add objProjetoItens
            
        Else
            
            'sendo de outro projeto... gerar o anterior
            Set objNFFatura = New ClassNFiscal
            
            'Gera a NFFatura a partir dos Itens do Projeto
'            lErro = Move_NFFatura_Memoria(objNFFatura, colNFFatura2)
            If lErro <> SUCESSO Then gError 134200
            
            'Grava NFFatura Gerado no BD
'            lErro = CF("Projeto_Grava_NFFatura", objNFFatura, colNFFatura2)
            If lErro <> SUCESSO Then gError 134201
            
            'reinicializa com o novo projeto
            lNumIntDocProj = objProjetoItens.lNumIntDocProj
            Set colNFFatura2 = New Collection
            
            'inclui na nova coleção
            colNFFatura2.Add objProjetoItens
            
        End If
    
    Next
    
    'sendo o último projeto ou projeto único... gerar este
    Set objNFFatura = New ClassNFiscal
    
    'Gera a NFFatura a partir dos Itens do Projeto
'    lErro = Move_NFFatura_Memoria(objNFFatura, colNFFatura2)
    If lErro <> SUCESSO Then gError 134202
    
    'Grava NFFatura Gerado no BD
'    lErro = CF("Projeto_Grava_NFFatura", objNFFatura, colNFFatura2)
    If lErro <> SUCESSO Then gError 134203
        
    Exporta_NFFatura = SUCESSO

    Exit Function

Erro_Exporta_NFFatura:

    Exporta_NFFatura = gErr

    Select Case gErr
    
        Case 134200 To 134203
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159890)

    End Select
    
    Exit Function
    
End Function

Private Function Move_PedidoDeVenda_Memoria(objPedidoDeVenda As ClassPedidoDeVenda, colPedidoDeVenda As Collection) As Long

Dim lErro As Long
Dim lNumPedido As Long
Dim objProjeto As ClassProjeto
Dim objCliente As ClassCliente
Dim objTipoCliente As New ClassTipoCliente
Dim dDesconto As Double
Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim objProjetoItens As New ClassProjetoItens
Dim objItemPedido As ClassItemPedido
Dim objProdutos As ClassProduto
Dim objCondicaoPagto As ClassCondicaoPagto
Dim iIndice As Integer
Dim objParcelaPV As ClassParcelaPedidoVenda

Dim objTributacaoItemPV As ClassTributacaoItemPV
Dim objItemPedidoNovo As ClassItemPedido

On Error GoTo Erro_PedidoDeVenda_Memoria

    Set objProjeto = New ClassProjeto
    
    objProjeto.lNumIntDoc = colPedidoDeVenda.Item(1).lNumIntDocProj
    
    'Le o Projeto
    lErro = CF("Projeto_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
    
    Set objCliente = New ClassCliente

    objCliente.lCodigo = objProjeto.lCodCliente

    'Busca o Cliente no BD
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 84037
                        
    'se encontrou...
    If lErro = SUCESSO Then
    
        objPedidoDeVenda.lCliente = objCliente.lCodigo
            
        'Se o Tipo estiver preenchido
        If objCliente.iTipo > 0 Then
            objTipoCliente.iCodigo = objCliente.iTipo
            'Lê o Tipo de Cliente
            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 126956
        End If

        'Guarda o Desconto do cliente
        If objCliente.dDesconto > 0 Then
            dDesconto = objCliente.dDesconto
        ElseIf objTipoCliente.dDesconto > 0 Then
            dDesconto = objTipoCliente.dDesconto
        Else
            dDesconto = 0
        End If

        'Preenche Tabela de Preço
        If objCliente.iTabelaPreco > 0 Then
            objPedidoDeVenda.iTabelaPreco = objCliente.iTabelaPreco
        ElseIf objTipoCliente.iTabelaPreco > 0 Then
            objPedidoDeVenda.iTabelaPreco = objTipoCliente.iTabelaPreco
        End If
        
        'Preenche Condição de Pagamento
        If objCliente.iCondicaoPagto > 0 Then
            objPedidoDeVenda.iCondicaoPagto = objCliente.iCondicaoPagto
        ElseIf objTipoCliente.iCondicaoPagto > 0 Then
            objPedidoDeVenda.iCondicaoPagto = objTipoCliente.iCondicaoPagto
        End If
  
    End If
                       
    objPedidoDeVenda.iFilial = objProjeto.iCodFilial
    objPedidoDeVenda.iFilialEntrega = objProjeto.iCodFilial
        
    'Preenche outros dados do PV
    objPedidoDeVenda.dtDataEmissao = gdtDataAtual
    objPedidoDeVenda.dtDataReferencia = gdtDataAtual
    objPedidoDeVenda.dtDataEntrega = DATA_NULA
    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa
    objPedidoDeVenda.iFilialEmpresaFaturamento = giFilialEmpresa
    objPedidoDeVenda.sNaturezaOp = "5101"
    
    dValorProdutos = 0
    
    'Para cada PV a ser gerado
    For Each objProjetoItens In colPedidoDeVenda
    
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objProjetoItens.sProduto
        
        'Le o Produto para pegar sua Descrição
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134768
    
        Set objItemPedido = New ClassItemPedido
        
        'Coloca os dados no Item do Pedido
        objItemPedido.sProduto = objProjetoItens.sProduto
        objItemPedido.sDescricao = objProdutos.sDescricao
        objItemPedido.sUnidadeMed = objProjetoItens.sUMedida
        objItemPedido.dQuantidade = objProjetoItens.dQuantidade
        objItemPedido.dPrecoUnitario = objProjetoItens.dPrecoTotalItem
        objItemPedido.dPrecoTotal = objProjetoItens.dQuantidade * objProjetoItens.dPrecoTotalItem
        objItemPedido.iFilialEmpresa = giFilialEmpresa
        objItemPedido.dtDataEntrega = objPedidoDeVenda.dtDataEntrega
                
        dValorProdutos = dValorProdutos + objProjetoItens.dQuantidade * objProjetoItens.dPrecoTotalItem
        
        Set objTributacaoItemPV = New ClassTributacaoItemPV
                
        'Adiciona o item na colecao de itens do pedido de venda
        With objItemPedido
            Set objItemPedidoNovo = objPedidoDeVenda.IncluirItem(.sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReserva, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, 0, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, 0, objTributacaoItemPV)
        End With
        
    Next
    
    objPedidoDeVenda.dValorProdutos = dValorProdutos
    
    'Se o cliente possui desconto
    If dDesconto > 0 Then

        'Calcula o valor do desconto para o cliente
        dValorDesconto = dDesconto * dValorProdutos

        'Para tributação
        objPedidoDeVenda.dValorDesconto = dValorDesconto

    End If
    
    objPedidoDeVenda.dValorTotal = dValorProdutos - dValorDesconto
    
    'Trata a Condição de Pagamento X Parcelas
    Set objCondicaoPagto = New ClassCondicaoPagto
    
    objCondicaoPagto.iCodigo = objPedidoDeVenda.iCondicaoPagto
    
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 84046
    
    'Se encontrou a Condição de Pagamento...
    If lErro <> SUCESSO Then
                    
        objCondicaoPagto.dValorTotal = objPedidoDeVenda.dValorTotal
        objCondicaoPagto.dtDataRef = objPedidoDeVenda.dtDataReferencia

        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True)
        If lErro <> SUCESSO Then gError 84076

        'Coloca os valores das Parcelas na coleção de Parcelas
        For iIndice = 1 To objCondicaoPagto.colParcelas.Count
            
            Set objParcelaPV = New ClassParcelaPedidoVenda

            objParcelaPV.iNumParcela = iIndice
            objParcelaPV.dValor = objCondicaoPagto.colParcelas(iIndice).dValor
            objParcelaPV.dtDataVencimento = objCondicaoPagto.colParcelas(iIndice).dtVencimento
            objParcelaPV.dtDesconto1Ate = DATA_NULA
            objParcelaPV.dtDesconto2Ate = DATA_NULA
            objParcelaPV.dtDesconto3Ate = DATA_NULA
            
            'Inclui na coleção
            With objParcelaPV
                objPedidoDeVenda.colParcelas.Add .dValor, .dtDataVencimento, .iNumParcela, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .dtDesconto3Ate, .dDesconto3Valor, .iDesconto3Codigo
            End With

        Next

    Else
        
        'Gera Parcela Única
        Set objParcelaPV = New ClassParcelaPedidoVenda
    
        objParcelaPV.iNumParcela = 1
        objParcelaPV.dtDataVencimento = gdtDataAtual
        objParcelaPV.dValor = objPedidoDeVenda.dValorTotal
        objParcelaPV.dtDesconto1Ate = DATA_NULA
        objParcelaPV.dtDesconto2Ate = DATA_NULA
        objParcelaPV.dtDesconto3Ate = DATA_NULA
        
        'Inclui na coleção
        With objParcelaPV
            objPedidoDeVenda.colParcelas.Add .dValor, .dtDataVencimento, .iNumParcela, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .dtDesconto3Ate, .dDesconto3Valor, .iDesconto3Codigo
        End With
    
    End If
    
    'Obtem o Código do Novo Pedido de Venda
    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_CODIGO_PEDVENDA", "PedVenTodos", "Codigo", lNumPedido)
    If lErro <> SUCESSO Then Error 26797
    
    objPedidoDeVenda.lCodigo = lNumPedido
    
    Move_PedidoDeVenda_Memoria = SUCESSO

    Exit Function

Erro_PedidoDeVenda_Memoria:

    Move_PedidoDeVenda_Memoria = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159891)

    End Select

    Exit Function

End Function

Private Function Move_OrdemDeProducao_Memoria(objOrdemDeProducao As ClassOrdemDeProducao, colOrdemDeProducao As Collection) As Long

'Dim lErro As Long
'Dim sCodigoOP As String
'Dim objProjeto As ClassProjeto
'Dim objProjetoItens As New ClassProjetoItens
'Dim objItemOP As ClassItemOP
'Dim objProdutos As ClassProduto
'
'On Error GoTo Erro_Move_OrdemDeProducao_Memoria
'
'    Set objProjeto = New ClassProjeto
'
'    objProjeto.lNumIntDoc = colOrdemDeProducao.Item(1).lNumIntDocProj
'
'    'Le o Projeto
'    lErro = CF("Projeto_Le", objProjeto)
'    If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
'
'    'Obtem o número da proxima ordem de producao
'    lErro = CF("OrdemProducao_Automatico", sCodigoOP, giFilialEmpresa)
'    If lErro <> SUCESSO Then gError 131860
'
'    'Preenche dados da OP
'    objOrdemDeProducao.sCodigo = sCodigoOP
'    objOrdemDeProducao.dtDataEmissao = gdtDataAtual
'    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
'
'    'Para cada Item da OP a ser gerado
'    For Each objProjetoItens In colOrdemDeProducao
'
'        Set objProdutos = New ClassProduto
'
'        objProdutos.sCodigo = objProjetoItens.sProduto
'
'        'Le o Produto para pegar sua Descrição
'        lErro = CF("Produto_Le", objProdutos)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 134768
'
'        Set objItemOrcamento = New ClassItemOV
'
'        'Coloca os dados no Item do Orçamento
'        objItemOP.sProduto = objProjetoItens.sProduto
'        objItemOP.sDescricao = objProdutos.sDescricao
'        objItemOP.sUnidadeMed = objProjetoItens.sUMedida
'        objItemOP.dQuantidade = objProjetoItens.dQuantidade
'        objItemOP.dPrecoUnitario = objProjetoItens.dPrecoTotalItem
'        objItemOP.dPrecoTotal = objProjetoItens.dQuantidade * objProjetoItens.dPrecoTotalItem
'        objItemOP.iFilialEmpresa = giFilialEmpresa
'        objItemOP.dtDataEntrega = objOrdemDeProducao.dtDataEmissao
'
'        'Adiciona o item na colecao de itens da ordem de producao
'        objOrdemDeProducao.colItens.Add objItemOP
'
'        Set objItemOP = New ClassItemOP
'
'        objItemOP.sCodigo = objOrdemDeProducao.sCodigo
'        objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
'
'        sProduto = GridOP.TextMatrix(iIndice, GRIDOP_PRODUTO_COL)
'
'        'Critica o formato do Produto
'        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then Error 36620
'
'        objItemOP.sProduto = sProdutoFormatado
'
'        objItemOP.sSiglaUM = GridOP.TextMatrix(iIndice, GRIDOP_UM_COL)
'
'        If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_QUANT_COL))) > 0 Then
'            objItemOP.dQuantidade = CDbl(GridOP.TextMatrix(iIndice, GRIDOP_QUANT_COL))
'        Else
'            objItemOP.dQuantidade = 0
'        End If
'
'        objAlmoxarifado.sNomeReduzido = GridOP.TextMatrix(iIndice, GRIDOP_ALMOXARIFADO_COL)
'
'        If colCodigoNome.Count > 0 Then
'
'            For Each objCodigoNome In colCodigoNome
'                If objCodigoNome.sNome = objAlmoxarifado.sNomeReduzido Then
'                    objItemOP.iAlmoxarifado = objCodigoNome.iCodigo
'                    Exit For
'                End If
'            Next
'
'        End If
'
'        If objItemOP.iAlmoxarifado = 0 Then
'
'            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
'            If lErro <> SUCESSO And lErro <> 25056 Then Error 36621
'
'            'trata se o almoxarifado não existir
'            If lErro = 25056 Then Error 36622
'
'            objItemOP.iAlmoxarifado = objAlmoxarifado.iCodigo
'
'            colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
'
'        End If
'
'        sCcl = GridOP.TextMatrix(iIndice, GRIDOP_CCL_COL)
'
'        If Len(Trim(sCcl)) <> 0 Then
'
'            'Formata Ccl para BD
'            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
'            If lErro <> SUCESSO Then Error 36623
'
'        Else
'
'            sCclFormatada = ""
'
'        End If
'
'        objItemOP.sCcl = sCclFormatada
'
'        objItemOP.iItem = iIndice
'
'        'continuação de Move_Grid_Memoria
'        lErro = Move_Grid_Memoria1(objItemOP, iIndice)
'        If lErro <> SUCESSO Then Error 36626
'
'        objOrdemDeProducao.colItens.Add objItemOP
'
'        objOrdemDeProducao.iNumItens = objOrdemDeProducao.iNumItens + 1
'
'
'    Next
'
'    objOrdemDeProducao.dValorProdutos = dValorProdutos
'
'    Move_OrdemDeProducao_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_OrdemDeProducao_Memoria:
'
'    Move_OrdemDeProducao_Memoria = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159892)
'
'    End Select
'
'    Exit Function
'
End Function


