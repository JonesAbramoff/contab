VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ContratoSrvOcx 
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
      Caption         =   "Frame1"
      Height          =   5010
      Index           =   1
      Left            =   15
      TabIndex        =   21
      Top             =   780
      Width           =   9375
      Begin VB.Frame Frame3 
         Caption         =   "Outras Informações do Contrato"
         Height          =   1110
         Left            =   90
         TabIndex        =   46
         Top             =   3855
         Width           =   9270
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Data prevista para próxima cobrança:"
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
            Left            =   4590
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   52
            Top             =   300
            Width           =   3195
         End
         Begin VB.Label DataProxCobranca 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7830
            TabIndex        =   51
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label Label6 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   675
            Width           =   1080
         End
         Begin VB.Label Observacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1215
            TabIndex        =   49
            Top             =   675
            Width           =   7920
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Vlr Unitário:"
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
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label ValorUnitario 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1215
            TabIndex        =   47
            Top             =   255
            Width           =   1965
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   2850
         Left            =   75
         TabIndex        =   25
         Top             =   135
         Width           =   6300
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2280
            Picture         =   "ContratoSrvOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   405
            Width           =   300
         End
         Begin VB.ComboBox Item 
            Height          =   315
            Left            =   5475
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   885
            Width           =   750
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   330
            Left            =   1215
            TabIndex        =   4
            Top             =   2325
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contrato 
            Height          =   315
            Left            =   1215
            TabIndex        =   2
            Top             =   885
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1215
            TabIndex        =   0
            Top             =   390
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label FilCliContrato 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4410
            TabIndex        =   39
            Top             =   1380
            Width           =   1800
         End
         Begin VB.Label CliContrato 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1215
            TabIndex        =   38
            Top             =   1365
            Width           =   2655
         End
         Begin VB.Label ProdutoLabel 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   465
            TabIndex        =   37
            Top             =   1860
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
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
            Height          =   210
            Left            =   165
            TabIndex        =   36
            Top             =   2400
            Width           =   1050
         End
         Begin VB.Label DescricaoProduto1 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3045
            TabIndex        =   35
            Top             =   1845
            Width           =   3165
         End
         Begin VB.Label ContratoLabel 
            Caption         =   "Contrato:"
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
            Height          =   210
            Left            =   405
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   34
            Top             =   915
            Width           =   795
         End
         Begin VB.Label DescContrato 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2505
            TabIndex        =   33
            Top             =   885
            Width           =   2325
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   32
            Top             =   1410
            Width           =   660
         End
         Begin VB.Label Label1 
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
            Index           =   2
            Left            =   3900
            TabIndex        =   31
            Top             =   1440
            Width           =   465
         End
         Begin VB.Label CodigoLabel 
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
            Left            =   540
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   435
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Item:"
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
            Index           =   0
            Left            =   5010
            TabIndex        =   29
            Top             =   930
            Width           =   435
         End
         Begin VB.Label Produto1 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1215
            TabIndex        =   28
            Top             =   1845
            Width           =   1845
         End
         Begin VB.Label Label30 
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
            Left            =   2895
            TabIndex        =   27
            Top             =   2400
            Width           =   480
         End
         Begin VB.Label UM 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3465
            TabIndex        =   26
            Top             =   2325
            Width           =   780
         End
      End
      Begin VB.Frame FrameSerie 
         Caption         =   "Rastreamento por Números de Série"
         Height          =   2820
         Left            =   6420
         TabIndex        =   44
         Top             =   165
         Width           =   2940
         Begin VB.TextBox NumSerie 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   645
            MaxLength       =   20
            TabIndex        =   45
            Top             =   780
            Width           =   2010
         End
         Begin MSFlexGridLib.MSFlexGrid GridNumSerie 
            Height          =   2250
            Left            =   90
            TabIndex        =   7
            Top             =   345
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3969
            _Version        =   393216
         End
      End
      Begin VB.Frame FrameLote 
         Caption         =   "Rastreamento por Lote ou OP"
         Height          =   795
         Left            =   75
         TabIndex        =   40
         Top             =   3030
         Width           =   9285
         Begin VB.Frame FrameOP 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   2550
            TabIndex        =   41
            Top             =   195
            Width           =   4320
            Begin VB.ComboBox FilialOP 
               Height          =   315
               Left            =   915
               TabIndex        =   6
               Top             =   90
               Width           =   2805
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Filial OP:"
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
               Left            =   90
               TabIndex        =   42
               Top             =   150
               Width           =   780
            End
         End
         Begin MSMask.MaskEdBox Lote 
            Height          =   300
            Left            =   1215
            TabIndex        =   5
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin VB.Label LoteLabel 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   750
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   43
            Top             =   330
            Width           =   450
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4950
      Index           =   2
      Left            =   75
      TabIndex        =   12
      Top             =   810
      Visible         =   0   'False
      Width           =   9300
      Begin VB.Frame Frame7 
         Caption         =   "Serviços/Peças"
         Height          =   4275
         Left            =   495
         TabIndex        =   13
         Top             =   570
         Width           =   8235
         Begin VB.CommandButton BotaoServicos 
            Caption         =   "Serviços/Peças"
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
            Left            =   360
            TabIndex        =   24
            Top             =   3885
            Width           =   1740
         End
         Begin VB.TextBox DescricaoServico 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2790
            MaxLength       =   250
            TabIndex        =   15
            Top             =   2340
            Width           =   4560
         End
         Begin VB.CheckBox GarantiaTotal 
            Caption         =   "Todos c/exceção dos listados abaixo"
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
            Left            =   420
            TabIndex        =   14
            Top             =   255
            Width           =   5100
         End
         Begin MSMask.MaskEdBox Servico 
            Height          =   225
            Left            =   630
            TabIndex        =   16
            Top             =   2340
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridServicos 
            Height          =   3255
            Left            =   375
            TabIndex        =   17
            Top             =   660
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   5741
            _Version        =   393216
            Rows            =   6
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
         End
      End
      Begin MSMask.MaskEdBox TipoGarantia 
         Height          =   315
         Left            =   915
         TabIndex        =   18
         Top             =   150
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LblTipoGarantia 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   210
         Width           =   450
      End
      Begin VB.Label DescTipoGarantia 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1545
         TabIndex        =   19
         Top             =   150
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7320
      ScaleHeight     =   450
      ScaleWidth      =   2115
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   15
      Width           =   2175
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1605
         Picture         =   "ContratoSrvOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1095
         Picture         =   "ContratoSrvOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   585
         Picture         =   "ContratoSrvOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   90
         Picture         =   "ContratoSrvOcx.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5430
      Left            =   -15
      TabIndex        =   23
      Top             =   435
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   9578
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Serviços Contratados"
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
Attribute VB_Name = "ContratoSrvOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iTipoAlterado As Integer
Dim iAlterado As Integer
Dim iContratoAlterado As Integer


'Eventos de browser
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoContrato As AdmEvento
Attribute objEventoContrato.VB_VarHelpID = -1
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1

Dim objGridServico As AdmGrid
Dim objGridNumSerie As AdmGrid

Dim iGrid_Servico_Col As Integer
Dim iGrid_ServicoDesc_Col As Integer

Dim iGrid_NumSerie_Col As Integer

Dim giFrameAtual As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giFrameAtual = 1

    'Inicializa eventos de browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoContrato = New AdmEvento
    Set objEventoLote = New AdmEvento
    Set objEventoServico = New AdmEvento
    Set objEventoTipo = New AdmEvento

    Set objGridServico = New AdmGrid
    Set objGridNumSerie = New AdmGrid

    Call Inicializa_Grid_Servico(objGridServico)
    Call Inicializa_Grid_NumSerie(objGridNumSerie)

    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 195526

    'Inicializa a Máscara de Servico
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Servico)
    If lErro <> SUCESSO Then gError 195527

    lErro_Chama_Tela = SUCESSO

    Exit Function

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 195526, 195527

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195528)

    End Select

End Function

Private Function Inicializa_Grid_Servico(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Desc. Serviço")

    objGridInt.colCampo.Add (Servico.Name)
    objGridInt.colCampo.Add (DescricaoServico.Name)

    'Controles que participam do Grid
    iGrid_Servico_Col = 1
    iGrid_ServicoDesc_Col = 2

    objGridInt.objGrid = GridServicos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_GARANTIA_SERVICOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridServicos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Servico = SUCESSO

End Function

Private Function Inicializa_Grid_NumSerie(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Número de Série")

    objGridInt.colCampo.Add (NumSerie.Name)

    'Controles que participam do Grid
    iGrid_NumSerie_Col = 1

    objGridInt.objGrid = GridNumSerie

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_NUM_SERIE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridNumSerie.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_NumSerie = SUCESSO

End Function

Public Function Trata_Parametros(Optional ByVal objItensDeContratoSrv As ClassItensDeContratoSrv) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se recebeu um objeto com dados de um relacionamento
    If Not (objItensDeContratoSrv Is Nothing) Then

        'Lê e traz os dados do relacionamento para a tela
        lErro = Traz_ItensDeContratoSrv_Tela(objItensDeContratoSrv)
        If lErro <> SUCESSO Then gError 195529

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 195529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195530)

    End Select

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    Set objEventoContrato = Nothing
    Set objEventoLote = Nothing
    Set objEventoServico = Nothing
    Set objEventoTipo = Nothing

    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 195531

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 195531

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195532)

    End Select

    Exit Sub

End Sub

Private Sub GarantiaTotal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LblTipoGarantia_Click()

Dim objTipoGarantia As New ClassTipoGarantia
Dim colSelecao As New Collection

    objTipoGarantia.lCodigo = StrParaLong(Codigo.Text)

    Call Chama_Tela("TipoGarantiaLista", colSelecao, objTipoGarantia, objEventoTipo)

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim objTipoGarantia As ClassTipoGarantia
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoGarantia = obj1

    'Lê o tipo
    lErro = CF("TipoGarantia_Le", objTipoGarantia)
    If lErro <> SUCESSO And lErro <> 183849 Then gError 195533

    'Se não encontrar --> Erro
    If lErro <> SUCESSO Then gError 195534

    TipoGarantia.Text = objTipoGarantia.lCodigo

    lErro = Exibe_Dados_TipoGarantia(objTipoGarantia)
    If lErro <> SUCESSO Then gError 195535

    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case gErr

        Case 195533, 195535

        Case 195534
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOGARANTIA_NAO_CADASTRADA", gErr, objTipoGarantia.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195536)

    End Select

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Veifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then Exit Sub

    'Critica a Quantidade
    lErro = Valor_Positivo_Critica(Quantidade.Text)
    If lErro <> SUCESSO Then gError 195537

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 195537

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195538)

    End Select

    Exit Sub

End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Lote_GotFocus()

    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Private Sub FilialOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialOP_Validate

    'Se não estiver preenchida ou alterada pula a crítica
    If Len(Trim(FilialOP.Text)) = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialOP, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195539

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 195540

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 195541

    Exit Sub

Erro_FilialOP_Validate:

    Cancel = True

    Select Case gErr

        Case 195539

        Case 195540
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 195541
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195542)

    End Select

    Exit Sub

End Sub

Public Sub FilialOP_Click()

Dim lErro As Long

On Error GoTo Erro_FilialOP_Click

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_FilialOP_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195543)

    End Select

    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim objItensDeContratoSrv As New ClassItensDeContratoSrv
Dim colSelecao As New Collection

    objItensDeContratoSrv.lCodigo = StrParaLong(Codigo.Text)

    Call Chama_Tela("ItensDeContratoSRVLista", colSelecao, objItensDeContratoSrv, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objItensDeContratoSrv As ClassItensDeContratoSrv
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objItensDeContratoSrv = obj1

    'Traz para a tela o relacionamento com código passado pelo browser
    lErro = Traz_ItensDeContratoSrv_Tela(objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195543

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 195543

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195544)

    End Select

End Sub

Private Sub ContratoLabel_Click()

Dim objContrato As New ClassContrato
Dim colSelecao As New Collection

    If Len(Trim(Contrato.Text)) > 0 Then
        objContrato.sCodigo = Codigo.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    End If

    Call Chama_Tela("ContratosLista", colSelecao, objContrato, objEventoContrato)

End Sub

Private Sub objEventoContrato_evSelecao(obj1 As Object)

Dim objContrato As ClassContrato
Dim bCancel As Boolean

    Set objContrato = obj1

    Contrato.Text = objContrato.sCodigo

    Call Contrato_Validate(bCancel)

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iContratoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Contrato_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Contrato_Validate

    If Len(Trim(Contrato.Text)) = 0 Then
        DescContrato.Caption = ""
        CliContrato.Caption = ""
        FilCliContrato.Caption = ""
        Item.Clear
        Produto1.Caption = ""
        DescricaoProduto1.Caption = ""
        
        Quantidade.Text = ""
        UM.Caption = ""
        DataProxCobranca.Caption = ""
        Observacao.Caption = ""
        ValorUnitario.Caption = ""
        Exit Sub
    End If

    objContrato.sCodigo = Contrato.Text
    objContrato.iFilialEmpresa = giFilialEmpresa

    If iContratoAlterado = REGISTRO_ALTERADO Then

        lErro = Traz_Contrato_Tela(objContrato)
        If lErro <> SUCESSO Then gError 195545

    End If

    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Contrato_Validate:

    Cancel = True

    Select Case gErr

        Case 195545

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195546)

    End Select

End Sub

Private Function Traz_Contrato_Tela(objContrato As ClassContrato) As Long

Dim lErro As Long
Dim sCclMascarado As String
Dim objcliente As New ClassCliente
Dim objItensDeContrato As ClassItensDeContrato
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Traz_Contrato_Tela

    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 195547

    If lErro <> SUCESSO Then gError 195548

    If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 195549

    DescContrato.Caption = objContrato.sDescricao
    
    objcliente.lCodigo = objContrato.lCliente

    lErro = CF("Cliente_Le", objcliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 195550

    CliContrato.Caption = objcliente.sNomeReduzido

    objFilialCliente.lCodCliente = objContrato.lCliente
    objFilialCliente.iCodFilial = objContrato.iFilCli

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 195551

    FilCliContrato.Caption = objFilialCliente.sNome

    Item.Clear
    Produto1.Caption = ""
    DescricaoProduto1.Caption = ""
    
    Quantidade.Text = ""
    UM.Caption = ""
    DataProxCobranca.Caption = ""
    Observacao.Caption = ""
    ValorUnitario.Caption = ""
        
    For Each objItensDeContrato In objContrato.colItens
        Item.AddItem objItensDeContrato.iSeq
        Item.ItemData(Item.NewIndex) = objItensDeContrato.lNumIntDoc
    Next
    
    If objContrato.colItens.Count = 1 Then
        Item.ListIndex = 0
        Call Item_Click
    End If

    Traz_Contrato_Tela = SUCESSO

    Exit Function

Erro_Traz_Contrato_Tela:

    Traz_Contrato_Tela = gErr

    Select Case gErr

        Case 195547, 195550, 195551

        Case 195548
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, objContrato.sCodigo)

        Case 195549
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195552)

    End Select

    Exit Function

End Function

Private Function Traz_ItensDeContratoSrv_Tela(objItensDeContratoSrv As ClassItensDeContratoSrv) As Long

Dim lErro As Long
Dim objContrato As New ClassContrato
Dim objTipoGarantia As New ClassTipoGarantia
Dim iIndice As Integer

On Error GoTo Erro_Traz_ItensDeContratoSrv_Tela

    lErro = CF("ItensDeContratoSRV_Le", objItensDeContratoSrv)
    If lErro <> SUCESSO And lErro <> 195584 Then gError 195610

    If lErro <> SUCESSO Then gError 195611

    'Limpa a tela
    Call Limpa_ItensDeContratoSRV

    Codigo.PromptInclude = False
    Codigo.Text = objItensDeContratoSrv.lCodigo
    Codigo.PromptInclude = True

    Contrato.Text = objItensDeContratoSrv.sCodigoContrato

    objContrato.sCodigo = objItensDeContratoSrv.sCodigoContrato
    objContrato.iFilialEmpresa = objItensDeContratoSrv.iFilialEmpresa

    lErro = Traz_Contrato_Tela(objContrato)
    If lErro <> SUCESSO Then gError 195612

    For iIndice = 0 To Item.ListCount - 1
    
        If Item.ItemData(iIndice) = objItensDeContratoSrv.lNumIntItemContrato Then
        
            Item.ListIndex = iIndice
            Exit For
        End If
    
    Next
    Call Item_Click

    Quantidade.Text = Formata_Estoque(objItensDeContratoSrv.dQuantidade)

    Lote.Text = objItensDeContratoSrv.sLote

    'Se o Rastreamento possui FilialOP (Rastro Por Ordem de Produção)
    If objItensDeContratoSrv.iFilialOP <> 0 Then

        For iIndice = 0 To FilialOP.ListCount - 1
            If FilialOP.ItemData(iIndice) = objItensDeContratoSrv.iFilialOP Then
                FilialOP.ListIndex = iIndice
                Exit For
            End If
        Next

    End If

    TipoGarantia.Text = objItensDeContratoSrv.lTipoGarantia

    objTipoGarantia.lCodigo = objItensDeContratoSrv.lTipoGarantia

    lErro = CF("TipoGarantia_Le", objTipoGarantia)
    If lErro <> SUCESSO And lErro <> 183849 Then gError 195613

    If lErro = SUCESSO Then
        DescTipoGarantia.Caption = objTipoGarantia.sDescricao
    End If

    GarantiaTotal.Value = objItensDeContratoSrv.iGarantiaTotal

    lErro = Carrega_Grid_Servicos(objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195614

    lErro = Carrega_Grid_NumSerie(objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195615

    iAlterado = 0

    Traz_ItensDeContratoSrv_Tela = SUCESSO

    Exit Function

Erro_Traz_ItensDeContratoSrv_Tela:

    Traz_ItensDeContratoSrv_Tela = gErr

    Select Case gErr

        Case 195610, 195612, 195613, 195614, 195615

        Case 195611
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCONTRATOSRV_NAO_CADASTRADO", gErr, objItensDeContratoSrv.iFilialEmpresa, objItensDeContratoSrv.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195616)

    End Select

    Exit Function

End Function

Private Sub Item_Click()

Dim lErro As Long
Dim objContrato As New ClassContrato
Dim objItensDeContrato As ClassItensDeContrato
Dim sProduto As String
Dim objProduto As ClassProduto

On Error GoTo Erro_Item_Click

    If Item.ListIndex <> -1 Then

        objContrato.sCodigo = Contrato.Text
        objContrato.iFilialEmpresa = giFilialEmpresa

        lErro = CF("Contrato_Le", objContrato)
        If lErro <> SUCESSO And lErro <> 129332 Then gError 195556

        For Each objItensDeContrato In objContrato.colItens

            If objItensDeContrato.iSeq = Item.Text Then

                lErro = Mascara_RetornaProdutoTela(objItensDeContrato.sProduto, sProduto)
                If lErro <> SUCESSO Then gError 195557

                Produto1.Caption = sProduto

                DescricaoProduto1.Caption = objItensDeContrato.sDescProd

                Quantidade.Text = Formata_Estoque(objItensDeContrato.dQuantidade)

                UM.Caption = objItensDeContrato.sUM
                
                DataProxCobranca.Caption = Format(objItensDeContrato.dtDataProxCobranca, "dd/mm/yyyy")
                Observacao.Caption = objItensDeContrato.sObservacao
                ValorUnitario.Caption = Format(objItensDeContrato.dValor, "STANDARD")
                
                Set objProduto = New ClassProduto
                
                objProduto.sCodigo = objItensDeContrato.sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 195557
                
                If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                    FrameLote.Enabled = True
                    FrameSerie.Enabled = False
                    FrameOP.Enabled = False
                    Call Grid_Limpa(objGridNumSerie)
                    FilialOP.ListIndex = -1
                ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                    FrameLote.Enabled = True
                    FrameSerie.Enabled = False
                    FrameOP.Enabled = True
                    Call Grid_Limpa(objGridNumSerie)
                ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    FrameLote.Enabled = False
                    FrameSerie.Enabled = True
                    FrameOP.Enabled = False
                    Lote.Text = ""
                    FilialOP.ListIndex = -1
                Else
                    FrameLote.Enabled = False
                    FrameSerie.Enabled = False
                    FrameOP.Enabled = False
                    Call Grid_Limpa(objGridNumSerie)
                    Lote.Text = ""
                    FilialOP.ListIndex = -1
                End If

                Exit For

            End If

        Next

    End If

    Exit Sub

Erro_Item_Click:

    Select Case gErr

        Case 195556, 195557

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195558)

    End Select

    Exit Sub

End Sub

Private Sub LoteLabel_Click()

Dim colSelecao As New Collection
Dim objRastroLote As New ClassRastreamentoLote
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sSelecao As String
Dim lErro As Long

On Error GoTo Erro_LoteLabel_Click

    objRastroLote.sCodigo = Lote.Text

    sProduto = Produto1.Caption

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195553

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195559

    'Selecao
    colSelecao.Add sProdutoFormatado

    sSelecao = "Produto = ?"

    'Chama tela de Browse de RastreamentoLote
    Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoLote, sSelecao)

    Exit Sub

Erro_LoteLabel_Click:

    Select Case gErr

        Case 195553

        Case 195559
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCONTRATO_NAO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195554)

    End Select

    Exit Sub

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As ClassRastreamentoLote
Dim iCodigo As Integer

On Error GoTo Erro_objEventoLote_evSelecao

    Set objRastroLote = obj1

    Lote.Text = objRastroLote.sCodigo

    If objRastroLote.iFilialOP <> 0 Then
    
        'Tenta selecionar na combo
        lErro = Combo_Seleciona(FilialOP, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195720

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr

        Case 195720

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195555)

    End Select

    Exit Sub

End Sub

Public Sub TipoGarantia_Change()

    iAlterado = REGISTRO_ALTERADO
    iTipoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoGarantia_GotFocus()

Dim iTipoGarantiaAux As Integer

    iTipoGarantiaAux = iTipoAlterado
    Call MaskEdBox_TrataGotFocus(TipoGarantia, iAlterado)
    iTipoAlterado = iTipoGarantiaAux

End Sub

Public Sub TipoGarantia_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoGarantia As New ClassTipoGarantia
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoGarantia_Validate

    'Verifica se o Tipo foi alterado
    If iTipoAlterado = 0 Then Exit Sub

    'Verifica se o Tipo está preenchido
    If Len(Trim(TipoGarantia.Text)) = 0 Then
        DescTipoGarantia.Caption = ""
        iTipoAlterado = 0
        Exit Sub
    End If

    'Critica o valor
    lErro = Long_Critica(TipoGarantia.Text)
    If lErro <> SUCESSO Then gError 195560

    objTipoGarantia.lCodigo = StrParaLong(TipoGarantia.Text)

    'Lê o tipo
    lErro = CF("TipoGarantia_Le", objTipoGarantia)
    If lErro <> SUCESSO And lErro <> 183849 Then gError 195561

    'Se não encontrar --> Erro
    If lErro <> SUCESSO Then gError 195562

    lErro = Exibe_Dados_TipoGarantia(objTipoGarantia)
    If lErro <> SUCESSO Then gError 195563

    iTipoAlterado = 0

    Exit Sub

Erro_TipoGarantia_Validate:

    Cancel = True

    Select Case gErr

        Case 195560, 195561, 195563

        Case 195562
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOGARANTIA", objTipoGarantia.lCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("TipoGarantia", objTipoGarantia)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195564)

    End Select

    Exit Sub

End Sub

Function Exibe_Dados_TipoGarantia(objTipoGarantia As ClassTipoGarantia) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String

On Error GoTo Erro_Exibe_Dados_TipoGarantia

    'Coloca a Descrição na Tela
    DescTipoGarantia.Caption = objTipoGarantia.sDescricao

    GarantiaTotal.Value = objTipoGarantia.iGarantiaTotal

    Call Grid_Limpa(objGridServico)

    'Exibe os dados da coleção na tela
    For iIndice = 1 To objTipoGarantia.colTipoGarantiaProduto.Count

        objProduto.sCodigo = objTipoGarantia.colTipoGarantiaProduto.Item(iIndice).sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195565
        
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Servico.PromptInclude = False
        Servico.Text = sProduto
        Servico.PromptInclude = True

        'Insere no Grid Categoria
        GridServicos.TextMatrix(iIndice, iGrid_Servico_Col) = Servico.Text
        GridServicos.TextMatrix(iIndice, iGrid_ServicoDesc_Col) = objProduto.sDescricao

    Next

    objGridServico.iLinhasExistentes = objTipoGarantia.colTipoGarantiaProduto.Count

    Exibe_Dados_TipoGarantia = SUCESSO

    Exit Function

Erro_Exibe_Dados_TipoGarantia:

    Exibe_Dados_TipoGarantia = gErr

    Select Case gErr

        Case 195565
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195566)

    End Select

    Exit Function

End Function

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoServicos_Click

    If Me.ActiveControl Is Servico Then

        sProduto1 = Servico.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 195567

        sProduto1 = GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 195568

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoServico)

    Exit Sub

Erro_BotaoServicos_Click:

    Select Case gErr

        Case 195567
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 195568

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195569)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridServicos.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 195570

    Servico.PromptInclude = False
    Servico.Text = sProduto
    Servico.PromptInclude = True

    If Not (Me.ActiveControl Is Servico) Then

        GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = Servico.Text

        'Faz o Tratamento do produto
        lErro = Traz_Servico_Tela()
        If lErro <> SUCESSO Then gError 195571

    End If

    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 195570
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 195571
            GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = ""

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195572)

    End Select

    Exit Sub

End Sub

Private Function Traz_Servico_Tela() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_Servico_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", Servico.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 195573

    If lErro = 51381 Then gError 195574

    'Descricao Servico
    GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao

    'Acrescenta uma linha no Grid se for o caso
    If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then

        objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1

    End If

    Traz_Servico_Tela = SUCESSO

    Exit Function

Erro_Traz_Servico_Tela:

    Traz_Servico_Tela = gErr

    Select Case gErr

        Case 195573

        Case 195574
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Servico.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195575)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 195576

    'Limpa a Tela
    Call Limpa_ItensDeContratoSRV

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 195576

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195577)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objItensDeContratoSrv As New ClassItensDeContratoSrv
Dim lErro As Long
Dim sAviso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 195580

    'Guarda no obj, código do contrato e filial empresa
    'Essas informações são necessárias para excluir a garantia
    objItensDeContratoSrv.lCodigo = StrParaLong(Codigo.Text)
    objItensDeContratoSrv.iFilialEmpresa = giFilialEmpresa

    'Lê a garantia
    lErro = CF("ItensDeContratoSRV_Le", objItensDeContratoSrv)
    If lErro <> SUCESSO And lErro <> 195584 Then gError 185586

    'Se não encontrou => erro
    If lErro <> SUCESSO Then gError 195587

    'Pede a confirmação da exclusão da garantia
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ITENSDECONTRATOSRV", objItensDeContratoSrv.lCodigo, objItensDeContratoSrv.iFilialEmpresa)

    If vbMsgRes = vbYes Then

        'Faz a exclusão da Solicitacao
        lErro = CF("ItensDeContratoSRV_Exclui", objItensDeContratoSrv)
        If lErro <> SUCESSO Then gError 195588

        'Limpa a Tela de Orcamento de Venda
        Call Limpa_ItensDeContratoSRV

        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 195580
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 195586, 195588
        
        Case 195587
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCONTRATOSRV_NAO_CADASTRADO", gErr, objItensDeContratoSrv.iFilialEmpresa, objItensDeContratoSrv.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195589)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 195590

    'Limpa a Tela
    Call Limpa_ItensDeContratoSRV

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 195590

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195591)

    End Select

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código de relacionamento para giFilialEmpresa
    lErro = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_ITENSDECONTRATOSRV", "ItensDeContratoSRV", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 195592

    'Exibe o código obtido
    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 195592

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195593)

    End Select

End Sub

Private Sub TabStrip1_Click()

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195607)

    End Select

    Exit Sub

End Sub


'*** TRATAMENTO DO EVENTO KEYDOWN  - INÍCIO ***
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is Contrato Then
            Call ContratoLabel_Click
        ElseIf Me.ActiveControl Is Lote Then
            Call LoteLabel_Click
        ElseIf Me.ActiveControl Is Servico Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is TipoGarantia Then
            Call LblTipoGarantia_Click
        End If

    End If

End Sub


'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Contrato de Manutenção"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "ContratoSrv"
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
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

Private Function Carrega_Grid_Servicos(objItensDeContratoSrv As ClassItensDeContratoSrv) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sServicoEnxuto As String
Dim objItensDeContratoSrvProd As ClassItensDeContratoSrvProd
Dim objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_Servicos

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridServico)

    For iIndice = 1 To objItensDeContratoSrv.colProduto.Count

        Set objItensDeContratoSrvProd = objItensDeContratoSrv.colProduto(iIndice)

        lErro = Mascara_RetornaProdutoEnxuto(objItensDeContratoSrvProd.sProduto, sServicoEnxuto)
        If lErro <> SUCESSO Then gError 195617

        'Mascara o produto enxuto
        Servico.PromptInclude = False
        Servico.Text = sServicoEnxuto
        Servico.PromptInclude = True

        GridServicos.TextMatrix(iIndice, iGrid_Servico_Col) = Servico.Text

        objProduto.sCodigo = objItensDeContratoSrvProd.sProduto

        'Lê o Servico
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195618

        If lErro = SUCESSO Then
            GridServicos.TextMatrix(iIndice, iGrid_ServicoDesc_Col) = objProduto.sDescricao
        End If

    Next

    'Atualiza o número de linhas existentes
    objGridServico.iLinhasExistentes = objItensDeContratoSrv.colProduto.Count

    Carrega_Grid_Servicos = SUCESSO

    Exit Function

Erro_Carrega_Grid_Servicos:

    Carrega_Grid_Servicos = gErr

    Select Case gErr

        Case 195617, 195618

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195619)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_NumSerie(objItensDeContratoSrv As ClassItensDeContratoSrv) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sServicoEnxuto As String
Dim objItensContSrvNumSerie As ClassItensContSrvNumSerie

On Error GoTo Erro_Carrega_Grid_NumSerie

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridNumSerie)

    For iIndice = 1 To objItensDeContratoSrv.colNumSerie.Count

        Set objItensContSrvNumSerie = objItensDeContratoSrv.colNumSerie(iIndice)

        GridNumSerie.TextMatrix(iIndice, iGrid_NumSerie_Col) = objItensContSrvNumSerie.sNumSerie

    Next

    'Atualiza o número de linhas existentes
    objGridNumSerie.iLinhasExistentes = objItensDeContratoSrv.colNumSerie.Count

    Carrega_Grid_NumSerie = SUCESSO

    Exit Function

Erro_Carrega_Grid_NumSerie:

    Carrega_Grid_NumSerie = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195620)

    End Select

    Exit Function

End Function

Private Sub GridServicos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)

    End If

End Sub

Private Sub GridServicos_EnterCell()

    Call Grid_Entrada_Celula(objGridServico, iAlterado)

End Sub

Private Sub GridServicos_GotFocus()

    Call Grid_Recebe_Foco(objGridServico)

End Sub

Private Sub GridServicos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)
    End If


End Sub

Private Sub GridServicos_LeaveCell()

    Call Saida_Celula(objGridServico)

End Sub

Private Sub GridServicos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridServico)

End Sub

Private Sub GridServicos_Scroll()

    Call Grid_Scroll(objGridServico)

End Sub

Private Sub GridServicos_RowColChange()

    Call Grid_RowColChange(objGridServico)

End Sub

Private Sub GridServicos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridServico)

End Sub

Public Sub Servico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Servico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Servico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Servico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Servico
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If objGridInt.objGrid.Name = GridServicos.Name Then

            'Verifica qual a coluna atual do Grid
            Select Case objGridInt.objGrid.Col

                'Se for a de Servico
                Case iGrid_Servico_Col
                    lErro = Saida_Celula_Servico(objGridInt)
                    If lErro <> SUCESSO Then gError 195621

            End Select


        Else

            Select Case objGridInt.objGrid.Col

                'Se for a de Servico
                Case iGrid_NumSerie_Col
                    'lErro = Saida_Celula_NumSerie(objGridInt)
                    lErro = Saida_Celula_Padrao(objGridInt, NumSerie, True)
                    If lErro <> SUCESSO Then gError 195622


            End Select

        End If


        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 195623

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 195621 To 195623

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195624)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Servico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Servico

    Set objGridInt.objControle = Servico

    If Len(Trim(Servico.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica", Servico.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 195625

        'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
        If lErro <> SUCESSO Then gError 195626

    Else

        GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195627

    If Len(Trim(Servico.ClipText)) <> 0 Then

        GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao

        If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then

            objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1

        End If

    End If

    Saida_Celula_Servico = SUCESSO

    Exit Function

Erro_Saida_Celula_Servico:

    Saida_Celula_Servico = gErr

    Select Case gErr

        Case 195625, 195627
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195626
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Servico.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Servico.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195628)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NumSerie(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NumSerie

    Set objGridInt.objControle = NumSerie

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195629

    Saida_Celula_NumSerie = SUCESSO

    Exit Function

Erro_Saida_Celula_NumSerie:

    Saida_Celula_NumSerie = gErr

    Select Case gErr

        Case 195629
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195630)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 195631

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOP.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOP.ItemData(FilialOP.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 195631

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195632)

    End Select

    Exit Function

End Function

Private Sub Limpa_ItensDeContratoSRV()

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela(Me)

    FilialOP.ListIndex = -1
    Item.Clear

    DescContrato.Caption = ""
    CliContrato.Caption = ""
    FilCliContrato.Caption = ""
    Produto1.Caption = ""
    DescricaoProduto1.Caption = ""
    UM.Caption = ""
    DataProxCobranca.Caption = ""
    Observacao.Caption = ""
    ValorUnitario.Caption = ""

    Call Grid_Limpa(objGridServico)
    Call Grid_Limpa(objGridNumSerie)

    iAlterado = 0

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objItensDeContratoSrv As New ClassItensDeContratoSrv

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se todos os campos obrigatórios estão preenchidos
    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 195633

    'Move os dados da tela para o objItensDeContratoSrv
    lErro = Move_ItensDeContratoSrv_Memoria(objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195634

    'Verifica se essa solicitação já existe no BD
    'e, em caso positivo, alerta ao usuário que está sendo feita uma alteração
    lErro = Trata_Alteracao(objItensDeContratoSrv, objItensDeContratoSrv.iFilialEmpresa, objItensDeContratoSrv.lCodigo)
    If lErro <> SUCESSO Then gError 195635

    'Grava no BD
    lErro = CF("ItensDeContratoSrv_Grava", objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195636

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 195633 To 195636

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195637)

    End Select

    Exit Function

End Function

Private Function Valida_Gravacao() As Long
'Verifica se os dados da tela são válidos para a gravação do registro

Dim lErro As Long
Dim iIndice As Integer
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Gravacao

    'Se o código não estiver preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 195638

    'Se o contrato não estiver preenchido => erro
    If Len(Trim(Contrato.ClipText)) = 0 Then gError 195639

    'Se o item nao estiver selecionado => erro
    If Len(Trim(Item.Text)) = 0 Then gError 195640

    'Se a quantidade não estiver preenchido => erro
    If Len(Trim(Quantidade.Text)) = 0 Then gError 195641

    If StrParaDbl(Quantidade.Text) = 0 Then gError 195642

    For iIndice = 1 To objGridServico.iLinhasExistentes

        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Servico_Col))) = 0 Then gError 195643

    Next

    Valida_Gravacao = SUCESSO

    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr

    Select Case gErr

        Case 195638
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 195639
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_PREENCHIDO", gErr)

        Case 195640
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCONTRATO_NAO_PREENCHIDO", gErr)

        Case 195641
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr)

        Case 195642
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ZERADA", gErr)

        Case 195643
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195644)

    End Select

End Function

Private Function Move_ItensDeContratoSrv_Memoria(objItensDeContratoSrv As ClassItensDeContratoSrv) As Long
'Move os dados da tela para objItensDeContratoSrv

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim iPreenchido As Integer
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_ItensDeContratoSrv_Memoria

    objItensDeContratoSrv.iFilialEmpresa = giFilialEmpresa

    objItensDeContratoSrv.lCodigo = StrParaLong(Codigo.Text)

    objItensDeContratoSrv.lNumIntItemContrato = Item.ItemData(Item.ListIndex)

    objItensDeContratoSrv.dQuantidade = StrParaDbl(Quantidade.Text)

    objItensDeContratoSrv.sLote = Lote.Text
    
    sProduto = Produto1.Caption
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195665
    
    objItensDeContratoSrv.sProduto = sProdutoFormatado

    objItensDeContratoSrv.iFilialOP = Codigo_Extrai(FilialOP.Text)

    objItensDeContratoSrv.lTipoGarantia = StrParaLong(TipoGarantia.Text)

    objItensDeContratoSrv.iGarantiaTotal = GarantiaTotal.Value

    Set objItensDeContratoSrv.objTela = Me

    'Move Grid Itens para memória
    lErro = Move_GridServico_Memoria(objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195645

    lErro = Move_GridNumSerie_Memoria(objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195646

    Move_ItensDeContratoSrv_Memoria = SUCESSO

    Exit Function

Erro_Move_ItensDeContratoSrv_Memoria:

    Move_ItensDeContratoSrv_Memoria = gErr

    Select Case gErr

        Case 195645, 195646, 195665

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195647)

    End Select

    Exit Function

End Function

Private Function Move_GridServico_Memoria(objItensDeContratoSrv As ClassItensDeContratoSrv) As Long
'Recolhe do Grid os dados

Dim lErro As Long
Dim sProduto As String
Dim sServico As String
Dim iPreenchido As Integer
Dim dQuantidade As Double
Dim objItensDeContratoSrvProd As ClassItensDeContratoSrvProd
Dim iIndice As Integer

On Error GoTo Erro_Move_GridServico_Memoria

    For iIndice = 1 To objGridServico.iLinhasExistentes

        Set objItensDeContratoSrvProd = New ClassItensDeContratoSrvProd

        'Formata o produto
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Servico_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 195648

        If iPreenchido = PRODUTO_VAZIO Then gError 195649

        objItensDeContratoSrvProd.sProduto = sProduto

        objItensDeContratoSrv.colProduto.Add objItensDeContratoSrvProd

    Next

    Move_GridServico_Memoria = SUCESSO

    Exit Function

Erro_Move_GridServico_Memoria:

    Move_GridServico_Memoria = gErr

    Select Case gErr

        Case 195648

        Case 195649
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195650)

    End Select

    Exit Function

End Function

Private Function Move_GridNumSerie_Memoria(objItensDeContratoSrv As ClassItensDeContratoSrv) As Long
'Recolhe do Grid os dados

Dim lErro As Long
Dim sProduto As String
Dim sServico As String
Dim iPreenchido As Integer
Dim dQuantidade As Double
Dim objItensContSrvNumSerie As ClassItensContSrvNumSerie
Dim iIndice As Integer

On Error GoTo Erro_Move_GridNumSerie_Memoria

    For iIndice = 1 To objGridNumSerie.iLinhasExistentes

        Set objItensContSrvNumSerie = New ClassItensContSrvNumSerie

        objItensContSrvNumSerie.sNumSerie = GridNumSerie.TextMatrix(iIndice, iGrid_NumSerie_Col)

        objItensDeContratoSrv.colNumSerie.Add objItensContSrvNumSerie

    Next

    Move_GridNumSerie_Memoria = SUCESSO

    Exit Function

Erro_Move_GridNumSerie_Memoria:

    Move_GridNumSerie_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195651)

    End Select

    Exit Function

End Function

'**** TRATAMENTO DO SISTEMA DE SETAS - INÍCIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objItensDeContratoSrv As New ClassItensDeContratoSrv
Dim objCampoValor As AdmCampoValor
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ItensDeContratoSrv"

    'Guarda no obj os dados que serão usados para identifica o registro a ser exibido
    objItensDeContratoSrv.lCodigo = StrParaLong(Trim(Codigo.Text))
    objItensDeContratoSrv.iFilialEmpresa = giFilialEmpresa

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objItensDeContratoSrv.lCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", objItensDeContratoSrv.iFilialEmpresa, 0, "FilialEmpresa"

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195652)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objItensDeContratoSrv As New ClassItensDeContratoSrv

On Error GoTo Erro_Tela_Preenche

    'Guarda o código do campo em questão no obj
    objItensDeContratoSrv.lCodigo = colCampoValor.Item("Codigo").vValor
    objItensDeContratoSrv.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor

    lErro = Traz_ItensDeContratoSrv_Tela(objItensDeContratoSrv)
    If lErro <> SUCESSO Then gError 195653

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 195653

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195654)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub
'**** FIM DO TRATAMENTO DO SISTEMA DE SETAS ****


Private Sub GridNumSerie_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNumSerie, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNumSerie, iAlterado)

    End If

End Sub

Private Sub GridNumSerie_EnterCell()

    Call Grid_Entrada_Celula(objGridNumSerie, iAlterado)

End Sub

Private Sub GridNumSerie_GotFocus()

    Call Grid_Recebe_Foco(objGridNumSerie)

End Sub

Private Sub GridNumSerie_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNumSerie, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNumSerie, iAlterado)
    End If


End Sub

Private Sub GridNumSerie_LeaveCell()

    Call Saida_Celula(objGridNumSerie)

End Sub

Private Sub GridNumSerie_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridNumSerie)

End Sub

Private Sub GridNumSerie_Scroll()

    Call Grid_Scroll(objGridNumSerie)

End Sub

Private Sub GridNumSerie_RowColChange()

    Call Grid_RowColChange(objGridNumSerie)

End Sub

Private Sub GridNumSerie_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridNumSerie)

End Sub

Public Sub NumSerie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NumSerie_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNumSerie)

End Sub

Public Sub NumSerie_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNumSerie)

End Sub

Public Sub NumSerie_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNumSerie.objControle = NumSerie
    lErro = Grid_Campo_Libera_Foco(objGridNumSerie)
    If lErro <> SUCESSO Then Cancel = True

End Sub
