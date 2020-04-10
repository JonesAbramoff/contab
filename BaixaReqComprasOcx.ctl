VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaReqComprasOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8265
      Index           =   2
      Left            =   195
      TabIndex        =   33
      Top             =   735
      Visible         =   0   'False
      Width           =   16530
      Begin VB.CheckBox Baixa 
         DragMode        =   1  'Automatic
         Height          =   210
         Left            =   195
         TabIndex        =   38
         Top             =   720
         Width           =   870
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   840
         Left            =   2400
         Picture         =   "BaixaReqComprasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   7350
         Width           =   1440
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   840
         Left            =   615
         Picture         =   "BaixaReqComprasOcx.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   7350
         Width           =   1440
      End
      Begin VB.TextBox NomeRed 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   9810
         TabIndex        =   42
         Text            =   "Nome"
         Top             =   3825
         Width           =   3000
      End
      Begin VB.CommandButton BotaoRequisicao 
         Caption         =   "Editar Requisição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5880
         Picture         =   "BaixaReqComprasOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   7350
         Width           =   1440
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "BaixaReqComprasOcx.ctx":2E7A
         Left            =   1725
         List            =   "BaixaReqComprasOcx.ctx":2E8A
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   225
         Width           =   3480
      End
      Begin VB.TextBox DataLimite 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   6750
         TabIndex        =   44
         Text            =   "Data Limite"
         Top             =   1005
         Width           =   1185
      End
      Begin VB.TextBox Requisitante 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   9795
         TabIndex        =   41
         Text            =   "Requisitante"
         Top             =   3345
         Width           =   3000
      End
      Begin VB.TextBox Ccl 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2115
         TabIndex        =   40
         Text            =   "Ccl"
         Top             =   720
         Width           =   990
      End
      Begin VB.TextBox Requisicao 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1095
         TabIndex        =   39
         Text            =   "Requisição"
         Top             =   705
         Width           =   1100
      End
      Begin VB.CommandButton BotaoBaixa 
         Caption         =   "Baixar Requisições"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5865
         Picture         =   "BaixaReqComprasOcx.ctx":2EC6
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   60
         Width           =   1830
      End
      Begin VB.TextBox Data 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   6315
         TabIndex        =   43
         Text            =   "Data"
         Top             =   1290
         Width           =   1185
      End
      Begin VB.TextBox PercentMinRecebItens 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   5790
         TabIndex        =   45
         Text            =   "% Min Recebimento Itens"
         Top             =   810
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid GridReqCompras 
         Height          =   6510
         Left            =   270
         TabIndex        =   37
         Top             =   750
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   11483
         _Version        =   393216
         Rows            =   11
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label4 
         Caption         =   "Ordenados por:"
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
         Left            =   375
         TabIndex        =   34
         Top             =   270
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   15045
      ScaleHeight     =   495
      ScaleWidth      =   1770
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   45
      Width           =   1830
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   90
         Picture         =   "BaixaReqComprasOcx.ctx":302C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   570
         Picture         =   "BaixaReqComprasOcx.ctx":355E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8295
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   720
      Width           =   16515
      Begin VB.Frame Frame7 
         Caption         =   "Exibe Requisições"
         Height          =   4830
         Left            =   585
         TabIndex        =   2
         Top             =   315
         Width           =   6960
         Begin VB.Frame Frame11 
            Caption         =   "Data"
            Height          =   735
            Left            =   555
            TabIndex        =   19
            Top             =   2910
            Width           =   5970
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   1665
               TabIndex        =   21
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
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   2805
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   4200
               TabIndex        =   24
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
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   5340
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   270
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label13 
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
               Left            =   3735
               TabIndex        =   23
               Top             =   330
               Width           =   360
            End
            Begin VB.Label Label12 
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
               Left            =   1215
               TabIndex        =   20
               Top             =   300
               Width           =   315
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Requisições"
            Height          =   1065
            Left            =   585
            TabIndex        =   3
            Top             =   285
            Width           =   5970
            Begin VB.CheckBox SoResiduais 
               Caption         =   "Somente residuais"
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
               Left            =   1245
               TabIndex        =   4
               Top             =   270
               Width           =   1905
            End
            Begin MSMask.MaskEdBox RequisicaoDe 
               Height          =   300
               Left            =   1650
               TabIndex        =   6
               Top             =   645
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox RequisicaoAte 
               Height          =   300
               Left            =   4200
               TabIndex        =   8
               Top             =   660
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label RequisicaoDeLabel 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1230
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   5
               Top             =   705
               Width           =   315
            End
            Begin VB.Label RequisicaoAteLabel 
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
               Left            =   3735
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   7
               Top             =   720
               Width           =   360
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Centros de Custo/Lucro"
            Height          =   735
            Left            =   570
            TabIndex        =   9
            Top             =   1410
            Width           =   5970
            Begin MSMask.MaskEdBox CclDe 
               Height          =   300
               Left            =   1680
               TabIndex        =   11
               Top             =   300
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CclAte 
               Height          =   300
               Left            =   4200
               TabIndex        =   13
               Top             =   300
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               PromptChar      =   " "
            End
            Begin VB.Label CclAteLabel 
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
               Left            =   3735
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   12
               Top             =   315
               Width           =   360
            End
            Begin VB.Label CclDeLabel 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1200
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   10
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data Limite"
            Height          =   735
            Left            =   555
            TabIndex        =   26
            Top             =   3675
            Width           =   5970
            Begin MSMask.MaskEdBox DataLimiteDe 
               Height          =   300
               Left            =   1620
               TabIndex        =   28
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
            Begin MSComCtl2.UpDown UpDownLimiteDe 
               Height          =   300
               Left            =   2775
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   270
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataLimiteAte 
               Height          =   300
               Left            =   4230
               TabIndex        =   31
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
            Begin MSComCtl2.UpDown UpDownLimiteAte 
               Height          =   300
               Left            =   5370
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   270
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label5 
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
               Left            =   1200
               TabIndex        =   27
               Top             =   330
               Width           =   315
            End
            Begin VB.Label Label2 
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
               Left            =   3765
               TabIndex        =   30
               Top             =   330
               Width           =   360
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Requisitantes"
            Height          =   735
            Left            =   585
            TabIndex        =   14
            Top             =   2160
            Width           =   5970
            Begin MSMask.MaskEdBox RequisitanteDe 
               Height          =   300
               Left            =   1680
               TabIndex        =   16
               Top             =   300
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox RequisitanteAte 
               Height          =   300
               Left            =   4200
               TabIndex        =   18
               Top             =   300
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label RequisitanteDeLabel 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1200
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   15
               Top             =   360
               Width           =   315
            End
            Begin VB.Label RequisitanteAteLabel 
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
               Left            =   3735
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   17
               Top             =   315
               Width           =   360
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8775
      Left            =   105
      TabIndex        =   0
      Top             =   345
      Width           =   16770
      _ExtentX        =   29580
      _ExtentY        =   15478
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisições"
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
Attribute VB_Name = "BaixaReqComprasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Grid de Requisicoes de Compras
Dim objGridReqCompras As AdmGrid
Dim iGrid_Baixa_Col As Integer
Dim iGrid_Requisicao_Col As Integer
Dim iGrid_CCL_Col As Integer
Dim iGrid_Requisitante_Col As Integer
Dim iGrid_Nome_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_MinRecebItens_Col As Integer

Dim asOrdenacao(3) As String
Dim asOrdenacaoString(3) As String
Public iAlterado As Integer
Dim iFramePrincipalAlterado As Integer
Dim gobjBaixaReqCompras As ClassBaixaReqCompras
Dim iFrameAtual As Integer

Const TAB_Selecao = 1

Private WithEvents objEventoRequisicaoDe As AdmEvento
Attribute objEventoRequisicaoDe.VB_VarHelpID = -1
Private WithEvents objEventoRequisicaoAte As AdmEvento
Attribute objEventoRequisicaoAte.VB_VarHelpID = -1
Private WithEvents objEventoCclDe As AdmEvento
Attribute objEventoCclDe.VB_VarHelpID = -1
Private WithEvents objEventoCclAte As AdmEvento
Attribute objEventoCclAte.VB_VarHelpID = -1
Private WithEvents objEventoRequisitanteDe As AdmEvento
Attribute objEventoRequisitanteDe.VB_VarHelpID = -1
Private WithEvents objEventoRequisitanteAte As AdmEvento
Attribute objEventoRequisitanteAte.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    'Preenche a combo de ordenacao com as 4 opcoes possíveis de ordenação
    asOrdenacao(0) = "RequisicaoCompra.Codigo"
    asOrdenacao(1) = "RequisicaoCompra.Data, RequisicaoCompra.Codigo"
    asOrdenacao(2) = "RequisicaoCompra.DataLimite,RequisicaoCompra.Codigo"
    asOrdenacao(3) = "RequisicaoCompra.Requisitante,RequisicaoCompra.Codigo"

    asOrdenacaoString(0) = "Código da Requisição"
    asOrdenacaoString(1) = "Data da Requisição"
    asOrdenacaoString(2) = "Data Limite de Entrega"
    asOrdenacaoString(3) = "Requisitante"

    iFrameAtual = 1

    'Inicializa as variáveis globais
    Set objEventoRequisicaoDe = New AdmEvento
    Set objEventoRequisicaoAte = New AdmEvento
    Set objEventoRequisitanteDe = New AdmEvento
    Set objEventoRequisitanteAte = New AdmEvento
    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento
    Set objGridReqCompras = New AdmGrid
    Set gobjBaixaReqCompras = New ClassBaixaReqCompras

    'Executa inicializacao do GridPedidos
    lErro = Inicializa_Grid_ReqCompras(objGridReqCompras)
    If lErro <> SUCESSO Then Error 63341

    'Inicializa mascara dos Ccl's
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then Error 63342

     'Limpa a Combobox Ordenados
    Ordenados.Clear
    'Carrega a Combobox Ordenados
    For iIndice = 0 To 3
        Ordenados.AddItem asOrdenacaoString(iIndice)
    Next

    'Seleciona a primeira ordenação
    Ordenados.ListIndex = 0

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case Err

        Case 63341, 63342
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143447)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera as variáveis globais
    Set objEventoRequisicaoDe = Nothing
    Set objEventoRequisicaoAte = Nothing
    Set objEventoRequisitanteDe = Nothing
    Set objEventoRequisitanteAte = Nothing
    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing

    Set gobjBaixaReqCompras = Nothing
    Set objGridReqCompras = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)
End Sub

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a máscara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 63343

    'Mascara CclDe
    CclDe.Mask = sMascaraCcl
    CclAte.Mask = sMascaraCcl
    
    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = Err

    Select Case Err

        Case 63343
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143448)

    End Select

    Exit Function

End Function

Private Function Inicializa_MascaraCclAte() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_MascaraCclAte

    sMascaraCcl = String(STRING_CCL, 0)

    'lê a máscara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 63345

    'Mascara CclAte
    CclAte.Mask = sMascaraCcl

    Inicializa_MascaraCclAte = SUCESSO

    Exit Function

Erro_Inicializa_MascaraCclAte:

    Inicializa_MascaraCclAte = Err

    Select Case Err

        Case 63345
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143449)

    End Select

    Exit Function

End Function

Private Sub Baixa_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todas as Requisicoes do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridReqCompras.iLinhasExistentes
        'Desmarca na tela o pedido da linha do Grid em questão (OK que pedido em questão?)
        GridReqCompras.TextMatrix(iLinha, iGrid_Baixa_Col) = GRID_CHECKBOX_INATIVO
    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridReqCompras)

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

    SoResiduais.Value = vbUnchecked
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridReqCompras)
    
    TabStrip1.Tabs.Item(TAB_Selecao).Selected = True
    
    Exit Sub
    
End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todas as Requisicoes do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridReqCompras.iLinhasExistentes
        'Marca na tela o pedido em questão (Ok que pedido em questão?)
        GridReqCompras.TextMatrix(iLinha, iGrid_Baixa_Col) = GRID_CHECKBOX_ATIVO
    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridReqCompras)

    Exit Sub

End Sub

Private Sub BotaoRequisicao_Click()
'Chama a tela de Requisicao de Compras passando a requisicao
'de compras selecionada
Dim objRequisicaoCompras As New ClassRequisicaoCompras

    'Verifica se alguma linha do grid de Requisicoes está selecionada
    If GridReqCompras.Row > 0 Then

        'Guarda código e FilialEmpresa da Requisicao de Compras
        objRequisicaoCompras.lCodigo = StrParaLong(GridReqCompras.TextMatrix(GridReqCompras.Row, iGrid_Requisicao_Col))
        objRequisicaoCompras.iFilialEmpresa = giFilialEmpresa

    End If

    'Chama a tela ReqCompras
    Call Chama_Tela("ReqComprasEnv", objRequisicaoCompras)

    Exit Sub

End Sub

Private Sub CclAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclAte_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(CclAte, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub CclAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CclDe_Validate

    'Verifica se Cclate está preenchido
    If Len(Trim(CclAte.ClipText)) = 0 Then Exit Sub

    'Passa o Ccl para o formato do BD
    lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then Error 63326

    objCcl.sCcl = sCclFormata

    'Le o Ccl formatado
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 63327

    'Se não encontrou o Ccl
    If lErro = 5599 Then Error 63328

    Exit Sub

Erro_CclDe_Validate:

    Cancel = True

    Select Case Err

        Case 63326, 63327
            'Erros tratados nas rotinas chamadas

        Case 63328
            'Avisa de deseja criar novo Ccl
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", sCclFormata)
            If vbMsgRes = vbYes Then Call Chama_Tela("CclTela", objCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143450)

    End Select

    Exit Sub

End Sub

Private Sub CclAteLabel_Click()

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

On Error GoTo Erro_CclAteLabel_Click

    'Verifica se CclAte está preenchido
    If Len(Trim(CclAte.ClipText)) > 0 Then

        'Coloca CclAte no formato do Banco de Dados
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then Error 63320

        'Coloca Ccl Formatado em objCcl
        objCcl.sCcl = sCclFormata

    End If

    'Ok o objEvento que você está passando está errado.
    'Chama a tela CclLista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclAte)

    Exit Sub

Erro_CclAteLabel_Click:

    Select Case Err

        Case 63320
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143451)

    End Select

    Exit Sub

End Sub

Private Sub CclDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclDe_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(CclDe, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub CclDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CclDe_Validate

    'Verifica se CclDe está preenchido
    If Len(Trim(CclDe.ClipText)) = 0 Then Exit Sub

    'Passa o Ccl para o formato do BD
    lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then Error 63323

    objCcl.sCcl = sCclFormata

    'Le o Ccl formatado
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 63324

    'Se não encontrou o Ccl
    If lErro = 5599 Then Error 63325

    Exit Sub

Erro_CclDe_Validate:

    Cancel = True

    Select Case Err

        Case 63323, 63324
            'Erros tratados nas rotinas chamadas

        Case 63325
            'Avisa de deseja criar novo Ccl
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", sCclFormata)

            If vbMsgRes = vbYes Then
                'Chama a tela CclTela
                Call Chama_Tela("CclTela", objCcl)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143452)

    End Select

    Exit Sub

End Sub

Private Sub CclDeLabel_Click()

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

On Error GoTo Erro_CclDeLabel_Click

    'Verifica se CclDe está preenchido
    If Len(Trim(CclDe.ClipText)) > 0 Then

        'Coloca CclDe no formato do Banco de Dados
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then Error 63319

        'Coloca Ccl Formatado em objCcl
        objCcl.sCcl = sCclFormata

    End If

    'Chama a tela CclLista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclDe)

    Exit Sub

Erro_CclDeLabel_Click:

    Select Case Err

        Case 63319
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143453)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataAte está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataAte informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then Error 63330

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case Err

        Case 63330
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143454)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then Error 63329

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case Err

        Case 63329
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143455)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimiteAte_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataLimiteAte, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub DataLimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteAte_Validate

    'Verifica se a DataLimiteAte está preenchida
    If Len(Trim(DataLimiteAte.Text)) = 0 Then Exit Sub

    'Critica a DataLimiteAte informada
    lErro = Data_Critica(DataLimiteAte.Text)
    If lErro <> SUCESSO Then Error 63332

    Exit Sub

Erro_DataLimiteAte_Validate:

    Cancel = True

    Select Case Err

        Case 63332
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143456)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimiteDe_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataLimiteDe, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub DataLimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteDe_Validate

    'Verifica se a DataLimiteDe está preenchida
    If Len(Trim(DataLimiteDe.Text)) = 0 Then Exit Sub

    'Critica a DataLimiteDe informada
    lErro = Data_Critica(DataLimiteDe.Text)
    If lErro <> SUCESSO Then Error 63331

    Exit Sub

Erro_DataLimiteDe_Validate:

    Cancel = True

    Select Case Err

        Case 63331
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143457)

    End Select

    Exit Sub
End Sub

Private Sub objEventoCclAte_evSelecao(obj1 As Object)

Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclAte_evSelecao

    Set objCcl = obj1

    'Mascara o Ccl
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 63322

    'Coloca o Ccl mascarado na tela
    CclAte.PromptInclude = False
    CclAte.Text = sCclMascarado
    CclAte.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoCclAte_evSelecao:

    Select Case Err

        Case 63322
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143458)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCclDe_evSelecao(obj1 As Object)

Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclDe_evSelecao

    Set objCcl = obj1

    'Mascara o Ccl
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 63321

    'Coloca o Ccl mascarado na tela
    CclDe.PromptInclude = False
    CclDe.Text = sCclMascarado
    CclDe.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoCclDe_evSelecao:

    Select Case Err

        Case 63321
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143459)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRequisicaoAte_evSelecao(obj1 As Object)

Dim objRequisicao As ClassRequisicaoCompras

    Set objRequisicao = obj1

    'Coloca o código retornado em RequisicaoAte
    RequisicaoAte.Text = objRequisicao.lCodigo
    
    Me.Show

End Sub

Private Sub objEventoRequisicaoDe_evSelecao(obj1 As Object)

Dim objRequisicao As ClassRequisicaoCompras

    Set objRequisicao = obj1
    
    'Coloca o código retornado em RequisicaoDe
    RequisicaoDe.Text = objRequisicao.lCodigo
    
    Me.Show

End Sub

Private Sub objEventoRequisitanteAte_evSelecao(obj1 As Object)

Dim objRequisitante As ClassRequisitante

    Set objRequisitante = obj1

    'Coloca o código retornado em RequisicaoAte
    RequisitanteAte.Text = objRequisitante.lCodigo
    
    Me.Show

End Sub

Private Sub objEventoRequisitanteDe_evSelecao(obj1 As Object)

Dim objRequisitante As ClassRequisitante

    Set objRequisitante = obj1

    'Coloca o código retornado em RequisitanteDe
    RequisitanteDe.Text = objRequisitante.lCodigo
    
    Me.Show

End Sub

Private Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordenados_Click()

Dim lErro As Long
Dim colCampos As New Collection
Dim colReqMarcados As New Collection
Dim colSaida As New Collection
Dim iIndice As Integer
Dim iLinha As Integer

On Error GoTo Erro_Ordenados_Click

    'Se o grid não foi preenchido, sai da rotina
    If objGridReqCompras.iLinhasExistentes = 0 Then Exit Sub
    
    Select Case Ordenados.Text
    
        Case "Código da Requisição"
            colCampos.Add "lCodRequisicao"
            
        Case "Data da Requisição"
            colCampos.Add "dtData"
            colCampos.Add "lCodRequisicao"
            
        Case "Data Limite de Entrega"
            colCampos.Add "dtDataLimite"
            colCampos.Add "lCodRequisicao"
            
        Case "Requisitante"
            colCampos.Add "lRequisitante"
            colCampos.Add "lCodRequisicao"
            
    End Select

    'Ordena a coleção
    Call Ordena_Colecao(gobjBaixaReqCompras.colReqComprasInfo, colSaida, colCampos)
    Set gobjBaixaReqCompras.colReqComprasInfo = colSaida
    
    'Guarda as Requisicoes de Compra marcados
    For iIndice = 1 To objGridReqCompras.iLinhasExistentes
        If GridReqCompras.TextMatrix(iIndice, iGrid_Baixa_Col) = "1" Then
            colReqMarcados.Add CLng(GridReqCompras.TextMatrix(iIndice, iGrid_Requisicao_Col))
        End If
    Next
    
    Call Grid_Limpa(objGridReqCompras)
    
    'Preenche o GridPedido
    lErro = Grid_ReqCompras_Preenche(gobjBaixaReqCompras.colReqComprasInfo)
    If lErro <> SUCESSO Then Error 63346
    
    'Marca novamente as Requisicoes de Compra
    For iIndice = 1 To colReqMarcados.Count
        For iLinha = 1 To objGridReqCompras.iLinhasExistentes
            If CStr(colReqMarcados(iIndice)) = GridReqCompras.TextMatrix(iLinha, iGrid_Requisicao_Col) Then
                GridReqCompras.TextMatrix(iLinha, iGrid_Baixa_Col) = "1"
            End If
        Next
    Next
    
    Call Grid_Refresh_Checkbox(objGridReqCompras)
    
    Exit Sub
    
Erro_Ordenados_Click:

    Select Case Err

        Case 63346
            'Erro tratado na rotina chamada
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143460)

    End Select

    Exit Sub

End Sub

Private Sub RequisicaoAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RequisicaoAte_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(RequisicaoAte, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub RequisicaoAteLabel_Click()

Dim objRequisicao As New ClassRequisicaoCompras
Dim colSelecao As New Collection

    'Verifica se o código de RequisicaoAte foi preenchido
    If Len(Trim(RequisicaoAte.Text)) > 0 Then

        'Preenche objRequisicao com RequisicaoAte
        objRequisicao.lCodigo = StrParaInt(RequisicaoAte.Text)

    End If

    'Chama a tela RequisicaoComprasLista
    Call Chama_Tela("ReqComprasEnvLista", colSelecao, objRequisicao, objEventoRequisicaoAte)

    Exit Sub

End Sub

Private Sub RequisicaoDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RequisicaoDe_GotFocus()

Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(RequisicaoDe, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub RequisicaoDeLabel_Click()

Dim objRequisicao As New ClassRequisicaoCompras
Dim colSelecao As New Collection

    'Verifica se o código de RequisicaoDe foi preenchido
    If Len(Trim(RequisicaoDe.Text)) > 0 Then
        'Preenche objRequisicao com RequisicaoDe
        objRequisicao.lCodigo = StrParaInt(RequisicaoDe.Text)
    End If

    'Chama a tela RequisicaoComprasLista
    Call Chama_Tela("ReqComprasEnvLista", colSelecao, objRequisicao, objEventoRequisicaoDe)

    Exit Sub

End Sub

Private Sub RequisitanteAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RequisitanteAte_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(RequisitanteAte, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub RequisitanteAteLabel_Click()

Dim objRequisitante As New ClassRequisitante
Dim colSelecao As New Collection

    'Verifica se RequisitanteAte está preenchido
    If Len(Trim(RequisitanteAte.Text)) > 0 Then

        'Preenche o código do Requisitante em objRequisitante
        objRequisitante.lCodigo = StrParaInt(RequisitanteAte.Text)

    End If

    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoRequisitanteAte)

    Exit Sub

End Sub

Private Sub RequisitanteDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RequisitanteDe_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(RequisitanteDe, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Private Sub RequisitanteDeLabel_Click()

Dim objRequisitante As New ClassRequisitante
Dim colSelecao As New Collection

    'Verifica se RequisitanteDe está preenchido
    If Len(Trim(RequisitanteDe.Text)) > 0 Then
        'Preenche o código do Requisitante em objRequisitante
        objRequisitante.lCodigo = StrParaInt(RequisitanteDe.Text)
    End If

    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoRequisitanteDe)

    Exit Sub

End Sub

Private Function Inicializa_Grid_ReqCompras(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ReqCompras

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Requisição")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Requisitante")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("% Mín. Receb. Itens")

    'campos de edição do grid
    objGridInt.colCampo.Add (Baixa.Name)
    objGridInt.colCampo.Add (Requisicao.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Requisitante.Name)
    objGridInt.colCampo.Add (NomeRed.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (PercentMinRecebItens.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Baixa_Col = 1
    iGrid_Requisicao_Col = 2
    iGrid_CCL_Col = 3
    iGrid_Requisitante_Col = 4
    iGrid_Nome_Col = 5
    iGrid_Data_Col = 6
    iGrid_DataLimite_Col = 7
    iGrid_MinRecebItens_Col = 8

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridReqCompras

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REQUISICOES + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 24

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ReqCompras = SUCESSO

    Exit Function

End Function

Private Sub SoResiduais_Click()
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    'Se o frame selecionado foi o de Requisicoes e houve alteracao no tab de Selecao
    If TabStrip1.SelectedItem.Index = 2 And iFramePrincipalAlterado = REGISTRO_ALTERADO Then

        'Recolhe os dados do TabSelecao
        lErro = Move_TabSelecao_Memoria()
        If lErro <> SUCESSO Then Error 63347
        
        lErro = Traz_Requisicoes_Tela()
        If lErro <> SUCESSO Then Error 63348
        
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case 63347, 63348

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143461)

    End Select

    Exit Sub

End Sub

Function Traz_Requisicoes_Tela() As Long
'Traz para a tela as Requisicoes de Compra que tem as características
'definidas no TabSelecao

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Traz_Requisicoes_Tela

    'Limpa a colecao de Requisicoes
    Set gobjBaixaReqCompras.colReqComprasInfo = New Collection

    'Limpa o GridReqCompras
    Call Grid_Limpa(objGridReqCompras)

    'Lê as Requisicoes de Compras de acordo com as características definidas no TabSelecao
    lErro = CF("BaixaReqCompras_ObterRequisicoes", gobjBaixaReqCompras)
    If lErro <> SUCESSO Then Error 63354
       
    'Preenche o GridReqCompras
    Call Grid_ReqCompras_Preenche(gobjBaixaReqCompras.colReqComprasInfo)

    'Selecionar todas as requisicoes da tela
    For iLinha = 1 To objGridReqCompras.iLinhasExistentes
        'Marca a Requisicao na tela
        GridReqCompras.TextMatrix(iLinha, iGrid_Baixa_Col) = MARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridReqCompras)

    iFramePrincipalAlterado = 0

    Traz_Requisicoes_Tela = SUCESSO

    Exit Function

Erro_Traz_Requisicoes_Tela:

    Traz_Requisicoes_Tela = Err

    Select Case Err

        Case 63354
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143462)

    End Select

    Exit Function

End Function

Private Function Grid_ReqCompras_Preenche(colReqComprasInfo As Collection) As Long
'Preenche o Grid de Requisição de Compras com as Requisicoes de Compras
'que podem ser baixadas

Dim objReqComprasInfo As New ClassReqComprasInfo
Dim iLinha As Integer
Dim objRequisitante As New ClassRequisitante
Dim lErro As Long
Dim sCclMascarado As String

On Error GoTo Erro_Grid_ReqCompras_Preenche

    'Verifica se o número de pedidos encontrados é superior ao máximo permitido
    If colReqComprasInfo.Count + 1 > NUM_MAX_REQUISICOES Then Error 57487

    iLinha = 1

    For Each objReqComprasInfo In colReqComprasInfo

        'Preenche o GridReqCompras com os dados de objReqComprasInfo
        If Len(Trim(objReqComprasInfo.sCcl)) > 0 Then
            
            'Mascara o Ccl
            lErro = Mascara_MascararCcl(objReqComprasInfo.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 63824
            
            GridReqCompras.TextMatrix(iLinha, iGrid_CCL_Col) = sCclMascarado
        
        End If
        
        GridReqCompras.TextMatrix(iLinha, iGrid_Data_Col) = Format(objReqComprasInfo.dtData, "dd/mm/yyyy")
        If objReqComprasInfo.dtDataLimite <> DATA_NULA Then GridReqCompras.TextMatrix(iLinha, iGrid_DataLimite_Col) = Format(objReqComprasInfo.dtDataLimite, "dd/mm/yyyy")
        GridReqCompras.TextMatrix(iLinha, iGrid_MinRecebItens_Col) = Format(objReqComprasInfo.dMinPercRecItens, "Percent")
        GridReqCompras.TextMatrix(iLinha, iGrid_Nome_Col) = objReqComprasInfo.sNomeRequisitante
        GridReqCompras.TextMatrix(iLinha, iGrid_Requisicao_Col) = objReqComprasInfo.lCodRequisicao
        GridReqCompras.TextMatrix(iLinha, iGrid_Requisitante_Col) = objReqComprasInfo.lRequisitante
        iLinha = iLinha + 1
        
        'Atualiza o número de Linhas Existentes do GridReqCompras
        objGridReqCompras.iLinhasExistentes = objGridReqCompras.iLinhasExistentes + 1

    Next

    Call Grid_Refresh_Checkbox(objGridReqCompras)

    'Passa para o Obj o número de ReqCompra passados pela Coleção
    objGridReqCompras.iLinhasExistentes = colReqComprasInfo.Count

    Grid_ReqCompras_Preenche = SUCESSO

    Exit Function

Erro_Grid_ReqCompras_Preenche:

    Grid_ReqCompras_Preenche = Err

    Select Case Err

        Case 63824
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143463)

    End Select

    Exit Function

End Function
Public Function Trata_Parametros()

    Trata_Parametros = SUCESSO
    
    iAlterado = 0
    
    Exit Function
    
End Function

Function Move_TabSelecao_Memoria() As Long
'Recolhe os dados do TabSelecao da tela

Dim lErro As Long
Dim sCclFormataDe As String
Dim sCclFormataAte As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria
    
    Set gobjBaixaReqCompras = Nothing
    Set gobjBaixaReqCompras = New ClassBaixaReqCompras

    gobjBaixaReqCompras.iSoResiduais = SoResiduais.Value
    
    'Verifica se RequisicaoDe e RequisicaoAte estão preenchidas
    If Len(Trim(RequisicaoDe.Text)) > 0 And Len(Trim(RequisicaoAte.Text)) > 0 Then
        
        'Verifica se RequisicaoDe é maior que RequisicaoAte
        If StrParaLong(RequisicaoDe.Text) > StrParaLong(RequisicaoAte.Text) Then Error 63349
    
    End If
    
    gobjBaixaReqCompras.lRequisicaoDe = StrParaLong(RequisicaoDe.Text)
    gobjBaixaReqCompras.lRequisicaoAte = StrParaLong(RequisicaoAte.Text)
    
    'Verifica se CclDe e CclAte estão preenchidos
    If Len(Trim(CclDe.ClipText)) > 0 Then
    
        'Passa CclDe para o formato do BD
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormataDe, iCclPreenchida)
        If lErro <> SUCESSO Then Error 63825
    
        gobjBaixaReqCompras.sCclDe = sCclFormataDe
    End If
    
    If Len(Trim(CclAte.ClipText)) > 0 Then
        
        'Passa CclAte para o formato do BD
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormataAte, iCclPreenchida)
        If lErro <> SUCESSO Then Error 63826
        
        gobjBaixaReqCompras.sCclAte = sCclFormataAte
              
    End If
    
    'Verifica se CclDe e CclAte estão preenchidos
    If Len(Trim(CclDe.ClipText)) > 0 And Len(Trim(CclAte.ClipText)) > 0 Then
        'Verifica se CclDe é maior que CclAte
        If gobjBaixaReqCompras.sCclDe > gobjBaixaReqCompras.sCclAte Then Error 63350
    End If
        
    'Verifica se RequisitanteDe e RequisitanteAte estão preenchidos
    If Len(Trim(RequisitanteDe.Text)) > 0 And Len(Trim(RequisitanteAte.Text)) > 0 Then
        
        'Verifica se RequisitanteDe é maior que RequisitanteAte
        If StrParaLong(RequisitanteDe.Text) > StrParaLong(RequisitanteAte.Text) Then Error 63351
    
    End If
    
    gobjBaixaReqCompras.lRequisitanteDe = StrParaLong(RequisitanteDe.Text)
    gobjBaixaReqCompras.lRequisitanteAte = StrParaLong(RequisitanteAte.Text)
        
    'Verifica se DataDe e DataAte estão preenchidas
    If Len(Trim(DataDe.ClipText)) > 0 And Len(Trim(DataAte.ClipText)) > 0 Then
        
        'Verifica se DataDe é maior que DataAte
        If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then Error 63352
        
    End If
    
    gobjBaixaReqCompras.dtDataDe = StrParaDate(DataDe.Text)
    gobjBaixaReqCompras.dtDataAte = StrParaDate(DataAte.Text)
    
    'Verifica se DataLimiteDe e DataLimiteAte estão preenchidas
    If Len(Trim(DataLimiteDe.ClipText)) > 0 And Len(Trim(DataLimiteAte.ClipText)) > 0 Then
        
        'Verifica se DataLimiteDe é maior que DataLimiteAte
        If StrParaDate(DataLimiteDe.Text) > StrParaDate(DataLimiteAte.Text) Then Error 63353
        
    End If
    
    gobjBaixaReqCompras.dtDataLimiteDe = StrParaDate(DataLimiteDe.Text)
    gobjBaixaReqCompras.dtDataLimiteAte = StrParaDate(DataLimiteAte.Text)
    
    gobjBaixaReqCompras.sOrdenacao = asOrdenacao(Ordenados.ListIndex)

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = Err

    Select Case Err

        Case 63349
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_INICIAL_MAIOR", Err)

        Case 63350
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", Err)

        Case 63351
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INICIAL_MAIOR", Err)

        Case 63352
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", Err)

        Case 63353
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITEDE_MAIOR", Err)

        Case 63825, 63826
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143464)

    End Select

    Exit Function

End Function

Private Sub GridReqCompras_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridReqCompras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridReqCompras, iAlterado)
    End If
    
    Exit Sub

End Sub

Private Sub GridReqCompras_GotFocus()
    Call Grid_Recebe_Foco(objGridReqCompras)
End Sub

Private Sub GridReqCompras_EnterCell()
    Call Grid_Entrada_Celula(objGridReqCompras, iAlterado)
End Sub
Private Sub Baixa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridReqCompras)
End Sub

Private Sub Baixa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridReqCompras)
End Sub

Private Sub Baixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridReqCompras.objControle = Baixa
    lErro = Grid_Campo_Libera_Foco(objGridReqCompras)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub


Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 63402

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 63402
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143465)

    End Select

    Exit Function

End Function

Private Sub GridReqCompras_LeaveCell()
    Call Saida_Celula(objGridReqCompras)
End Sub

Private Sub GridReqCompras_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridReqCompras)
End Sub

Private Sub GridReqCompras_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridReqCompras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridReqCompras, iAlterado)
    End If

End Sub

Private Sub GridReqCompras_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridReqCompras)
End Sub

Private Sub GridReqCompras_RowColChange()
    Call Grid_RowColChange(objGridReqCompras)
End Sub

Private Sub GridReqCompras_Scroll()
    Call Grid_Scroll(objGridReqCompras)
End Sub

Private Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Private Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    Set Form_Load_Ocx = Me
    Caption = "Baixa de Requisições de Compra"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "BaixaReqCompras"
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is RequisicaoDe Then
            Call RequisicaoDeLabel_Click
        ElseIf Me.ActiveControl Is RequisicaoAte Then
            Call RequisicaoAteLabel_Click
        ElseIf Me.ActiveControl Is CclDe Then
            Call CclDeLabel_Click
        ElseIf Me.ActiveControl Is CclAte Then
            Call CclAteLabel_Click
        ElseIf Me.ActiveControl Is RequisitanteDe Then
            Call RequisitanteDeLabel_Click
        ElseIf Me.ActiveControl Is RequisitanteAte Then
            Call RequisitanteAteLabel_Click
        End If
    End If

End Sub

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

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub RequisicaoDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RequisicaoDeLabel, Source, X, Y)
End Sub

Private Sub RequisicaoDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RequisicaoDeLabel, Button, Shift, X, Y)
End Sub

Private Sub RequisicaoAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RequisicaoAteLabel, Source, X, Y)
End Sub

Private Sub RequisicaoAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RequisicaoAteLabel, Button, Shift, X, Y)
End Sub

Private Sub CclAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclAteLabel, Source, X, Y)
End Sub

Private Sub CclAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclAteLabel, Button, Shift, X, Y)
End Sub

Private Sub CclDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclDeLabel, Source, X, Y)
End Sub

Private Sub CclDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclDeLabel, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub RequisitanteDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RequisitanteDeLabel, Source, X, Y)
End Sub

Private Sub RequisitanteDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RequisitanteDeLabel, Button, Shift, X, Y)
End Sub

Private Sub RequisitanteAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RequisitanteAteLabel, Source, X, Y)
End Sub

Private Sub RequisitanteAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RequisitanteAteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub


Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataAte
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 63334

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case Err

        Case 63334
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143466)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta um dia em DataAte
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 63337

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case Err

        Case 63337
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143467)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 63333

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case Err

        Case 63333
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143468)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 63338

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case Err

        Case 63338
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143469)

    End Select

    Exit Sub

End Sub

Private Sub UpDownLimiteAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteAte_DownClick

    'Diminui um dia em DataLimiteAte
    lErro = Data_Up_Down_Click(DataLimiteAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 63335

    Exit Sub

Erro_UpDownDataLimiteAte_DownClick:

    Select Case Err

        Case 63335
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143470)

    End Select

    Exit Sub

End Sub

Private Sub UpDownLimiteAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteAte_UpClick

    'Aumenta um dia em DataLimiteAte
    lErro = Data_Up_Down_Click(DataLimiteAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 63339

    Exit Sub

Erro_UpDownDataLimiteAte_UpClick:

    Select Case Err

        Case 63339
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143471)

    End Select

    Exit Sub

End Sub

Private Sub UpDownLimiteDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteDe_DownClick

    'Diminui um dia em DataLimiteDe
    lErro = Data_Up_Down_Click(DataLimiteDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 63336

    Exit Sub

Erro_UpDownDataLimiteDe_DownClick:

    Select Case Err

        Case 63336
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143472)

    End Select

    Exit Sub

End Sub

Private Sub UpDownLimiteDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteDe_UpClick

    'Aumenta um dia em DataLimiteDe
    lErro = Data_Up_Down_Click(DataLimiteDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 63340

    Exit Sub

Erro_UpDownDataLimiteDe_UpClick:

    Select Case Err

        Case 63340
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143473)

    End Select

    Exit Sub

End Sub

Private Sub BotaoBaixa_Click()
'Faz a baixa das Requisicoes de Compras selecionadas no Grid

Dim lErro As Long
Dim iLinha As Integer
Dim sNomeArqParam As String
Dim colReqComprasInfo As Collection

On Error GoTo Erro_BotaoBaixa_Click

    Set colReqComprasInfo = New Collection
    
    GL_objMDIForm.MousePointer = vbHourglass
        
    For iLinha = 1 To objGridReqCompras.iLinhasExistentes

        'Verifica se existe alguma Requisicao Selecionada no grid
        If GridReqCompras.TextMatrix(iLinha, iGrid_Baixa_Col) = MARCADO Then

            'adiciona na colecao as requisicoes de compra marcadas
            colReqComprasInfo.Add gobjBaixaReqCompras.colReqComprasInfo(iLinha)
        
        End If
        
    Next
    
    'Se nenhuma Requisicao do Grid estiver selecionada==>erro
    If colReqComprasInfo.Count = 0 Then Error 63356

    'Prepara para chamar rotina batch
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 63405

    'Chama rotina batch que calcula custo médio de produção
    'e valoriza movimentos de materiais produzidos
    lErro = CF("Rotina_ReqComprasBaixar_Batch", sNomeArqParam, colReqComprasInfo)
    If lErro <> SUCESSO Then Error 63355

    TabStrip1.Tabs(1).Selected = True
    iFramePrincipalAlterado = REGISTRO_ALTERADO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixa_Click:

    Select Case Err

        Case 63355, 63357, 63405
            'Erros tratados nas rotinas chamadas

        Case 63356
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REQUISICOES_BAIXAR", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143474)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub
