VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TipoClienteOcx 
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   8895
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4008
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   1020
      Visible         =   0   'False
      Width           =   8490
      Begin VB.ComboBox CondicaoPagto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   13
         Top             =   2730
         Width           =   2280
      End
      Begin VB.ComboBox TabelaPreco 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         Top             =   2078
         Width           =   2280
      End
      Begin VB.ComboBox Mensagem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   14
         Top             =   3353
         Width           =   4005
      End
      Begin VB.Frame SSFrame6 
         Height          =   510
         Left            =   135
         TabIndex        =   30
         Top             =   60
         Width           =   8265
         Begin VB.Label Tipo 
            Height          =   210
            Index           =   0
            Left            =   1710
            TabIndex        =   48
            Top             =   195
            Width           =   6300
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo de Cliente:"
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
            Left            =   195
            TabIndex        =   49
            Top             =   195
            Width           =   1365
         End
      End
      Begin MSMask.MaskEdBox LimiteCredito 
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   1448
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desconto 
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   818
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Caption         =   "Limite de Crédito:"
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
         Left            =   1275
         TabIndex        =   50
         Top             =   1500
         Width           =   1530
      End
      Begin VB.Label CondicaoPagtoLabel 
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
         Height          =   210
         Left            =   645
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   51
         Top             =   2760
         Width           =   2160
      End
      Begin VB.Label Label8 
         Caption         =   "Desconto:"
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
         Left            =   1920
         TabIndex        =   52
         Top             =   870
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "Tabela de Preços:"
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
         Left            =   1215
         TabIndex        =   53
         Top             =   2130
         Width           =   1590
      End
      Begin VB.Label Label9 
         Caption         =   "Mensagem para Nota Fiscal:"
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
         Left            =   345
         TabIndex        =   54
         Top             =   3405
         Width           =   2460
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4056
      Index           =   3
      Left            =   165
      TabIndex        =   15
      Top             =   975
      Visible         =   0   'False
      Width           =   8505
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   6315
         TabIndex        =   19
         Top             =   1620
         Width           =   1965
      End
      Begin VB.ComboBox Transportadora 
         Height          =   315
         Left            =   6315
         TabIndex        =   23
         Top             =   2940
         Width           =   1965
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   2085
         TabIndex        =   22
         Top             =   2940
         Width           =   1965
      End
      Begin VB.ComboBox PadraoCobranca 
         Height          =   315
         Left            =   2085
         TabIndex        =   20
         Top             =   2295
         Width           =   1965
      End
      Begin VB.Frame SSFrame1 
         Height          =   510
         Left            =   120
         TabIndex        =   29
         Top             =   60
         Width           =   8295
         Begin VB.Label Label3 
            Caption         =   "Tipo de Cliente:"
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
            Left            =   195
            TabIndex        =   37
            Top             =   195
            Width           =   1365
         End
         Begin VB.Label Tipo 
            Height          =   210
            Index           =   1
            Left            =   1710
            TabIndex        =   38
            Top             =   195
            Width           =   6300
         End
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   6315
         TabIndex        =   17
         Top             =   990
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FreqVisitas 
         Height          =   315
         Left            =   6315
         TabIndex        =   21
         Top             =   2295
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ComissaoVendas 
         Height          =   315
         Left            =   2085
         TabIndex        =   18
         Top             =   1620
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   315
         Left            =   2085
         TabIndex        =   16
         Top             =   990
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label ContaContabilLabel 
         AutoSize        =   -1  'True
         Caption         =   "Conta Contábil:"
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
         Left            =   4935
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label48 
         Caption         =   "dias"
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
         Left            =   6855
         TabIndex        =   40
         Top             =   2347
         Width           =   345
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Frequência de Visitas:"
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
         Left            =   4365
         TabIndex        =   41
         Top             =   2355
         Width           =   1905
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Comissão:"
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
         Left            =   1140
         TabIndex        =   42
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label VendedorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   43
         Top             =   1050
         Width           =   885
      End
      Begin VB.Label PadraoCobrancaLabel 
         AutoSize        =   -1  'True
         Caption         =   "Padrão de Cobrança:"
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
         TabIndex        =   44
         Top             =   2355
         Width           =   1815
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Cobrador:"
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
         Left            =   5415
         TabIndex        =   45
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label TransportadoraLabel 
         AutoSize        =   -1  'True
         Caption         =   "Transportadora:"
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
         Left            =   4920
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   46
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Região:"
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
         Left            =   1335
         TabIndex        =   47
         Top             =   3000
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4020
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   1005
      Width           =   8475
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2130
         Picture         =   "TipoClienteOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   315
         Width           =   300
      End
      Begin VB.ListBox Tipos 
         Height          =   3375
         ItemData        =   "TipoClienteOcx.ctx":00EA
         Left            =   5700
         List            =   "TipoClienteOcx.ctx":00EC
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   450
         Width           =   2625
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   1590
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1245
         Width           =   3675
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   2025
         Left            =   330
         TabIndex        =   32
         Top             =   1815
         Width           =   4995
         Begin VB.ComboBox ComboCategoriaClienteItem 
            Height          =   315
            Left            =   2460
            TabIndex        =   6
            Top             =   480
            Width           =   1632
         End
         Begin VB.ComboBox ComboCategoriaCliente 
            Height          =   315
            Left            =   945
            TabIndex        =   5
            Top             =   480
            Width           =   1548
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   1545
            Left            =   600
            TabIndex        =   7
            Top             =   390
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   2725
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
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Top             =   300
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Top             =   765
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label Label13 
         Caption         =   "Tipos de Cliente"
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
         Left            =   5685
         TabIndex        =   33
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   375
         TabIndex        =   34
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   35
         Top             =   825
         Width           =   930
      End
      Begin VB.Label Label1 
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
         Left            =   810
         TabIndex        =   36
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6555
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoClienteOcx.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoClienteOcx.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoClienteOcx.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoClienteOcx.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4470
      Left            =   120
      TabIndex        =   31
      Top             =   630
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   7885
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Financeiros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vendas"
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
Attribute VB_Name = "TipoClienteOcx"
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
Dim objGridCategoria As AdmGrid

Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoCondicaoPagto As AdmEvento
Attribute objEventoCondicaoPagto.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoPadraoCobranca As AdmEvento
Attribute objEventoPadraoCobranca.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1

Const GRID_CATEGORIA_COL = 1
Const GRID_VALOR_COL = 2

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_DadosFinanceiros = 2
Private Const TAB_Vendas = 3

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("TipoCliente_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57556

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57556
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174723)
    
    End Select

    Exit Sub

End Sub

Private Sub Cobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCobrador As New ClassCobrador
Dim iCodigo As Integer

On Error GoTo Erro_Cobrador_Validate

    'Verifica se foi preenchida a ComboBox Cobrador
    If Len(Trim(Cobrador.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Cobrador
    If Cobrador.Text = Cobrador.List(Cobrador.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Cobrador, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19275

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCobrador.iCodigo = iCodigo

        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 19276

        If lErro <> SUCESSO Then Error 19277 'Não encontrou Cobrador no BD

        'Encontrou Cobrador no BD, coloca no Text da Combo
        Cobrador.Text = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19278

    Exit Sub

Erro_Cobrador_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19275, 19276

    Case 19277  'Não encontrou Cobrador no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COBRADOR")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Cobradores", objCobrador)


        End If

    Case 19278

        lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", Err, Cobrador.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174724)

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

 Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se foi preenchido o campo Código do Cliente
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 19300

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 19300

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174725)

    End Select

    Exit Sub

End Sub

Private Sub ComboCategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCategoriaCliente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCategoriaCliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Private Sub ComboCategoriaCliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub ComboCategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaCliente
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ComboCategoriaClienteItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCategoriaClienteItem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCategoriaClienteItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaClienteItem_GotFocus

    Call Trata_ComboCategoriaClienteItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaClienteItem_GotFocus:

    Select Case Err

        Case 28933

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174726)

    End Select

    Exit Sub

End Sub

Private Sub ComboCategoriaClienteItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub ComboCategoriaClienteItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaClienteItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub CondicaoPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_CondicaoPagto_Validate

    'Verifica se foi preenchida a ComboBox CondicaoPagto
    If Len(Trim(CondicaoPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox CondicaoPagto
    If CondicaoPagto.Text = CondicaoPagto.List(CondicaoPagto.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CondicaoPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 33543

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 33544
        
        If lErro <> SUCESSO Then Error 33545 'Não encontrou CondicaoPagto no BD

        'Encontrou CondicaoPagto no BD e não é de Recebimento
        If objCondicaoPagto.iEmRecebimento = 0 Then Error 33546

        'Coloca no Text da Combo
        CondicaoPagto.Text = CondPagto_Traz(objCondicaoPagto)

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 33547

    Exit Sub

Erro_CondicaoPagto_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 33543, 33544
    
        Case 33545  'Não encontrou CondicaoPagto no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")
    
            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
    
            End If
    
        Case 33546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, iCodigo)
    
        Case 33547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagto.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174727)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

    Tipo(0).Caption = Trim(Descricao.Text)
    Tipo(1).Caption = Trim(Descricao.Text)

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load

    Set objEventoCondicaoPagto = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoPadraoCobranca = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    
    iFrameAtual = 1
    
    'preenche as combos da tela
    lErro = Preenche_Combos()
    If lErro <> SUCESSO Then Error 49753
    
    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        Set objEventoContaContabil = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 19054
    
        ContaContabil.Mask = sMascaraConta

    Else
       
        'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 19054
    
        ContaContabil.Mask = sMascaraConta
        
        ContaContabil.Enabled = False
        ContaContabilLabel.Enabled = False
        
    End If

    'Carrega a ComboBox CategoriaCliente com os códigos
    lErro = Carrega_ComboCategoriaCliente()
    If lErro <> SUCESSO Then Error 28916

    'Inicializa o Grid
    Set objGridCategoria = New AdmGrid
    
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then Error 28917

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 19054, 28916, 28917, 49753

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174728)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Preenche_Combos() As Long
'Preenche as combos da tela

Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCobrador As ClassCobrador
Dim ColCobrador As New Collection
Dim colPadroesCobranca As New Collection
Dim objPadraoCobranca As New ClassPadraoCobranca

On Error GoTo Erro_Preenche_Combos

    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê o Código e a Descrição de cada Tipo de Cliente
    lErro = CF("Cod_Nomes_Le", "TiposdeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 19050

    'preenche a ListBox Tipos com os objetos da colecao
    For Each objCodigoDescricao In colCodigoDescricao
        Tipos.AddItem objCodigoDescricao.sNome
        Tipos.ItemData(Tipos.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'Le cada Codigo e Descricao da tabela TabelasDePreco
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 19051

    'preenche a ComboBox TabelaPreco com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        TabelaPreco.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodigoDescricao.iCodigo

    Next
    
    lErro = CF("Carrega_CondicaoPagamento", CondicaoPagto, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 19052

'    Set colCodigoDescricao = New AdmColCodigoNome
'
'    'Lê cada código e descrição reduzida da tabela CondicoesPagto
'    lErro = CF("CondicoesPagto_Le_Recebimento", colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 19052
'
'    'Preenche a ComboBox CondicaoPagto com os objetos da coleção colCodigoDescricao
'    For Each objCodigoDescricao In colCodigoDescricao
'
'        CondicaoPagto.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
'        CondicaoPagto.ItemData(CondicaoPagto.NewIndex) = objCodigoDescricao.iCodigo
'
'    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'Le cada codigo e descricao da tabela Mensagem
    lErro = CF("Cod_Nomes_Le", "Mensagens", "Codigo", "Descricao", STRING_NFISCAL_MENSAGEM, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 19053

    'preenche a ComboBox Mensagem com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        Mensagem.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Mensagem.ItemData(Mensagem.NewIndex) = objCodigoDescricao.iCodigo

    Next
        
    Set colCodigoDescricao = New AdmColCodigoNome
    
    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 19055

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        Regiao.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'Le cada codigo e nome da tabela Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then Error 19056

    'preenche a ComboBox Cobrador com os objetos da colecao colCodigoDescricao
    For Each objCobrador In ColCobrador
        
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iInativo <> Inativo Then
            
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        
        End If

    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê todos os Padroes de Cobranca da tabela PadroesCobranca
    lErro = CF("PadroesCobranca_Le_Todos", colPadroesCobranca)
    If lErro <> SUCESSO Then Error 57339
    
    For Each objPadraoCobranca In colPadroesCobranca
        
        'Verifica se Padrao de Cobranca está ativo
        If objPadraoCobranca.iInativo <> Inativo Then
            
            PadraoCobranca.AddItem CStr(objPadraoCobranca.iCodigo) & SEPARADOR & objPadraoCobranca.sDescricao
            PadraoCobranca.ItemData(PadraoCobranca.NewIndex) = objPadraoCobranca.iCodigo

        End If

    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê códigos e nomes reduzidos da tabela Transportadoras
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 19058

    'Preenche a ComboBox Transportadora com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        Transportadora.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoDescricao.iCodigo

    Next
    
    Preenche_Combos = SUCESSO
    
    Exit Function
    
Erro_Preenche_Combos:
    
    Preenche_Combos = Err
    
    Select Case Err
        
        Case 19050, 19051, 19052, 19053, 19055, 19056, 17057, 17058, 57339
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174729)
    
    End Select
    
    Exit Function
    
End Function

Function Trata_Parametros(Optional objTipoCliente As ClassTipoCliente) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se tiver um TipoCliente selecionado, exibir seus dados
    If Not (objTipoCliente Is Nothing) Then

        lErro = CF("TipoDeCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 28943 Then Error 19073

        If lErro = SUCESSO Then
        
            lErro = Exibe_Dados_TipoCliente(objTipoCliente)
            If lErro <> SUCESSO Then Error 19074

        Else

            Codigo.Text = CStr(objTipoCliente.iCodigo)

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 19073, 19074

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174730)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub FreqVisitas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FreqVisitas_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(FreqVisitas, iAlterado)

End Sub

Private Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_EnterCell()

    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)

End Sub

Private Sub GridCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer

    iLinhasExistentesAnterior = objGridCategoria.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)
    
    If iLinhasExistentesAnterior <> objGridCategoria.iLinhasExistentes Then
    
        iAlterado = REGISTRO_ALTERADO
        
    End If
    
End Sub

Private Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_LeaveCell()

    Call Saida_Celula(objGridCategoria)

End Sub

Private Sub GridCategoria_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Private Sub GridCategoria_Scroll()

    Call Grid_Scroll(objGridCategoria)

End Sub

Private Sub Mensagem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMensagem As New ClassMensagem
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Mensagem_Validate

    'Verifica se foi preenchida a ComboBox Mensagem
    If Len(Trim(Mensagem.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Mensagem
    If Mensagem.Text = Mensagem.List(Mensagem.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Mensagem, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19267

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objMensagem.iCodigo = iCodigo

        'Tenta ler Mensagem com esse código no BD
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then Error 19268

        If lErro <> SUCESSO Then Error 19269 'Não encontrou Mensagem no BD

        'Encontrou Mensagem no BD, coloca no Text da Combo
        Mensagem.Text = CStr(objMensagem.iCodigo) & SEPARADOR & objMensagem.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19270

    Exit Sub

Erro_Mensagem_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19267, 19268

    Case 19269  'Não encontrou Mensagem no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_MENSAGEM")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Mensagens", objMensagem)

        End If

    Case 19270

        lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_ENCONTRADA", Err, Mensagem.Text)
        Mensagem.SetFocus

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174731)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PadraoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objPadraoCobranca As New ClassPadraoCobranca
Dim iCodigo As Integer

On Error GoTo Erro_PadraoCobranca_Validate

    'verifica se foi preenchido o campo PadraoCobranca
    If Len(Trim(PadraoCobranca.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox PadraoCobranca
    If PadraoCobranca.Text = PadraoCobranca.List(PadraoCobranca.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(PadraoCobranca, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19279

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objPadraoCobranca.iCodigo = iCodigo

        'Tenta ler Padrao Cobranca com esse código no BD
        lErro = CF("PadraoCobranca_Le", objPadraoCobranca)
        If lErro <> SUCESSO And lErro <> 19298 Then Error 19280

        If lErro = 19298 Then Error 19281 'Não encontrou Padrao Cobranca no BD

        'Encontrou Padrao Cobranca no BD, coloca no Text da Combo
        PadraoCobranca.Text = CStr(objPadraoCobranca.iCodigo) & SEPARADOR & objPadraoCobranca.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19282

    Exit Sub

Erro_PadraoCobranca_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19279, 19280

    Case 19281  'Não encontrou Padrao Cobranca no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PADRAO_COBRANCA")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("PadroesCobranca", objPadraoCobranca)

        End If

    Case 19282

        lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADO", Err, PadraoCobranca.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174732)

    End Select

    Exit Sub

End Sub

Private Sub Regiao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iCodigo As Integer

On Error GoTo Erro_Regiao_Validate

    'verifica se foi preenchido o campo Regiao
    If Len(Trim(Regiao.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox Regiao
    If Regiao.Text = Regiao.List(Regiao.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Regiao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19271

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tenta ler Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 19272
        If lErro <> SUCESSO Then Error 19273 'Não encontrou Regiao Venda BD

        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19274

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19271, 19272

    Case 19273  'Não encontrou RegiaoVenda no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

        If vbMsgRes = vbYes Then
            'Chama a tela de RegiaoVenda
            Call Chama_Tela("RegiaoVenda", objRegiaoVenda)

        End If

    Case 19274

        lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", Err, Regiao.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174733)

    End Select

    Exit Sub

End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTabelaPreco As New ClassTabelaPreco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19263

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then Error 19264

        If lErro = 28004 Then Error 19265 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19266

    Exit Sub

Erro_TabelaPreco_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 19263, 19264
    
        Case 19265  'Não encontrou Tabela de Preço no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")
    
            If vbMsgRes = vbYes Then
                'Chama a tela de Tabelas de Preço
                Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            End If
            
        Case 19266
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", Err, TabelaPreco.Text)
    
        Case Else
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174734)
    
        End Select

    Exit Sub

End Sub

Private Sub Tipos_DblClick()

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Tipos_DblClick

    'Guarda o valor do codigo do TipoCliente selecionado na ListBox Tipos
    objTipoCliente.iCodigo = Tipos.ItemData(Tipos.ListIndex)

    lErro = CF("TipoDeCliente_Le", objTipoCliente)
    If lErro <> SUCESSO And lErro <> 19062 Then Error 19083

    'verifica se TipoCliente nao esta cadastrado
    If lErro <> SUCESSO Then Error 19084

    lErro = Exibe_Dados_TipoCliente(objTipoCliente)
    If lErro <> SUCESSO Then Error 19085

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Tipos_DblClick:

    Tipos.SetFocus

    Select Case Err

    Case 19083, 19084, 19085

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174735)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 28964

    objTipoCliente.iCodigo = CInt(Codigo.Text)

    'Lê os dados do Tipo de Cliente
    lErro = CF("TipoDeCliente_Le", objTipoCliente)
    If lErro <> SUCESSO And lErro <> 28943 Then Error 28965

    'Se o Tipo de Cliente não está cadastrado ==> erro
    If lErro = 28943 Then Error 28966

    'Envia aviso perguntando se realmente deseja excluir o Tipo de Cliente
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_TIPOCLIENTE")

    If vbMsgRes = vbYes Then

        lErro = CF("TiposDeCliente_Exclui", objTipoCliente)
        If lErro <> SUCESSO Then Error 28967

        Call Tipos_Exclui(objTipoCliente.iCodigo)

        Call Limpa_Tela_TipoCliente

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 28964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_COD_NAO_PREENCHIDO", Err)

        Case 28965, 28967, 28968

        Case 28966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174736)

    End Select

    Exit Sub

End Sub

Private Sub LimiteCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LimiteCredito_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sLimiteCredito As String
Dim dLimiteCredito As Double

On Error GoTo Erro_LimiteCredito_Validate

    sLimiteCredito = LimiteCredito.Text

    'verifica se foi preenchida a ComboBox LimiteCredito
    If Len(Trim(sLimiteCredito)) = 0 Then Exit Sub

    lErro = Valor_NaoNegativo_Critica(sLimiteCredito)
    If lErro <> SUCESSO Then Error 19087

    LimiteCredito.Text = Format(sLimiteCredito, "Fixed")

    Exit Sub

Erro_LimiteCredito_Validate:

    Cancel = True



    Select Case Err

    Case 19087

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174737)

    End Select

    Exit Sub

End Sub

Private Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDesconto As String

On Error GoTo Erro_Desconto_Validate

    sDesconto = Desconto.Text

    'verifica se foi preenchido o Desconto
    If Len(Trim(Desconto.Text)) = 0 Then Exit Sub

    lErro = Porcentagem_Critica(Desconto.Text)
    If lErro <> SUCESSO Then Error 19088

    Desconto.Text = Format(sDesconto, "Fixed")

    Exit Sub

Erro_Desconto_Validate:

    Cancel = True


    Select Case Err

    Case 19088

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174738)

    End Select

    Exit Sub

End Sub

Private Sub TabelaPreco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabelaPreco_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Mensagem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Mensagem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaMascarada As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabil_Validate
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39809
        
    If lErro = SUCESSO Then
        
        sContaFormatada = objPlanoConta.sConta
            
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
            
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39810
            
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 19097
    
        If lErro = 5700 Then Error 19098
    
    End If
    
    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case Err

    Case 19097, 39809

    Case 19098
        lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)
    
    Case 39810
        lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
        
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174739)

    End Select

    Exit Sub

End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19283

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 19284

        If lErro <> SUCESSO Then Error 19285 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19286

    Exit Sub

Erro_Transportadora_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19283, 19284

    Case 19285  'Não encontrou Transportadora no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Transportadora", objTransportadora)

        End If

    Case 19286

        lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174740)

    End Select

    Exit Sub

End Sub

Private Sub TransportadoraLabel_Click()

'BROWSE TRANSPORTADORA :

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection
    
    'Se a Transportadora estiver preenchida passa o código para o objTransportadora
    If Len(Trim(Transportadora.Text)) > 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)
    
    'Chama a tela que lista as transportadoras
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora
Dim bCancel As Boolean

    Set objTransportadora = obj1

    'Preenche campo Transportadora
    Transportadora.Text = objTransportadora.iCodigo

    'Chama a rotina de validate
    Call Transportadora_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate
   
    'verifica se foi preenchido o campo Vendedor
    If Len(Trim(Vendedor.Text)) = 0 Then Exit Sub
    
    'Tenta ler o Vendedor (NomeReduzido ou Código)
    lErro = TP_Vendedor_Le2(Vendedor, objVendedor)
    If lErro <> SUCESSO And lErro <> 25032 Then Error 19099
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 19099
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174741)

    End Select

    Exit Sub

End Sub

Private Sub ComissaoVendas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComissaoVendas_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sComissaoVendas As String

On Error GoTo Erro_ComissaoVendas_Validate

    sComissaoVendas = ComissaoVendas.Text

    'verifica se foi preenchido a Comissao de Venda
    If Len(Trim(ComissaoVendas.Text)) = 0 Then Exit Sub

    lErro = Porcentagem_Critica(ComissaoVendas.Text)
    If lErro <> SUCESSO Then Error 19102

    ComissaoVendas.Text = Format(sComissaoVendas, "Fixed")

    Exit Sub

Erro_ComissaoVendas_Validate:

    Cancel = True


    Select Case Err

    Case 19102

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174742)

    End Select

    Exit Sub

End Sub

Private Sub Regiao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Regiao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cobrador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PadraoCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Transportadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Transportadora_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 19115

    Call Limpa_Tela_TipoCliente

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 19115

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174743)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim lCodigo As Long
Dim iAchou As Integer
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 19117

    'verifica se foi preenchida a Razao Social
    If Len(Trim(Descricao.Text)) = 0 Then Error 19118

    'verificar p/cada linha de "categoria"
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'se apenas a categoria ou apenas o valor estão preenchidos
        If (Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0) Or _
            (Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) = 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) <> 0) Then Error 56705

    Next
    
    'Le os dados da tela respectivos ao TipoCliente
    lErro = Move_Tela_Memoria(objTipoCliente)
    If lErro <> SUCESSO Then Error 19119

    lErro = Trata_Alteracao(objTipoCliente, objTipoCliente.iCodigo)
    If lErro <> SUCESSO Then Error 32284

    'Grava o Tipo de Cliente
    lErro = CF("TiposDeCliente_Grava", objTipoCliente)
    If lErro <> SUCESSO Then Error 19120

    'Remove o item da lista de Tipos
    Call Tipos_Exclui(objTipoCliente.iCodigo)

    'Insere o item na lista de Tipos
    Call Tipos_Adiciona(objTipoCliente)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 19117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_COD_NAO_PREENCHIDO", Err)

        Case 19118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_DESCR_NAO_PREENCHIDA", Err)
            
        Case 19119, 19120, 32284

        Case 56705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_CATEGCLI_INCOMPLETA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174744)

    End Select

    Exit Function

End Function

Private Sub Tipos_Adiciona(objTipoCliente As ClassTipoCliente)

    Tipos.AddItem objTipoCliente.sDescricao
    Tipos.ItemData(Tipos.NewIndex) = objTipoCliente.iCodigo

End Sub

Private Sub Tipos_Exclui(iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To Tipos.ListCount - 1

        If Tipos.ItemData(iIndice) = iCodigo Then

            Tipos.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 19188

    Call Limpa_Tela_TipoCliente

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 19188

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174745)

    End Select

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Function Move_Tela_Memoria(objTipoCliente As ClassTipoCliente) As Long
'Lê os dados que estão na tela de TipoCliente e coloca em objTipoCliente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objTipoClienteCategoria As New ClassTipoClienteCategoria
Dim objTabelaPreco As New ClassTabelaPreco
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objMensagem As New ClassMensagem
Dim objVendedor As New ClassVendedor
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim objCobrador As New ClassCobrador
Dim objPadraoCobranca As New ClassPadraoCobranca
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Move_Tela_Memoria

    'IDENTIFICACAO :

    If Len(Trim(Codigo.Text)) > 0 Then objTipoCliente.iCodigo = CLng(Codigo.Text)

    objTipoCliente.sDescricao = Descricao.Text

    objTipoCliente.sObservacao = Observacao.Text

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 19121

    If iContaPreenchida = CONTA_VAZIA Then
        objTipoCliente.sContaContabil = ""
    Else
        objTipoCliente.sContaContabil = sConta
    End If

    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) <> 0 Then

            Set objTipoClienteCategoria = New ClassTipoClienteCategoria

            If Len(Trim(Codigo.Text)) > 0 Then objTipoClienteCategoria.iTipoDeCliente = CInt(Codigo.Text)
            objTipoClienteCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objTipoClienteCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            objTipoCliente.colCategoriaItem.Add objTipoClienteCategoria

        End If

    Next

    'DADOS FINANCEIROS :

    If Len(Trim(LimiteCredito.Text)) > 0 Then objTipoCliente.dLimiteCredito = CDbl(LimiteCredito.Text)

    If Len(Trim(Desconto.Text)) > 0 Then objTipoCliente.dDesconto = CDbl(Desconto.Text) / 100

    If Len(Trim(TabelaPreco.Text)) > 0 Then objTipoCliente.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)

    If Len(Trim(CondicaoPagto.Text)) > 0 Then objTipoCliente.iCondicaoPagto = CondPagto_Extrai(CondicaoPagto)

    If Len(Trim(Mensagem.Text)) > 0 Then objTipoCliente.iCodMensagem = Codigo_Extrai(Mensagem.Text)

    'VENDAS :

    If Len(Trim(Vendedor.Text)) > 0 Then objTipoCliente.iVendedor = Codigo_Extrai(Vendedor.Text)

    If Len(Trim(ComissaoVendas.Text)) > 0 Then objTipoCliente.dComissaoVendas = CDbl(ComissaoVendas.Text) / 100

    If Len(Trim(Regiao.Text)) > 0 Then objTipoCliente.iRegiao = Codigo_Extrai(Regiao.Text)

    If Len(Trim(FreqVisitas.Text)) > 0 Then objTipoCliente.iFreqVisitas = CInt(FreqVisitas.Text)

    If Len(Trim(Cobrador.Text)) > 0 Then objTipoCliente.iCodCobrador = Codigo_Extrai(Cobrador.Text)

    If Len(Trim(PadraoCobranca)) > 0 Then objTipoCliente.iPadraoCobranca = Codigo_Extrai(PadraoCobranca.Text)

    If Len(Trim(Transportadora.Text)) > 0 Then objTipoCliente.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 19121, 33050, 33052, 33054, 33056, 33058, 33060, 33062, 33064

        Case 33051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", Err, objTabelaPreco.iCodigo)

        Case 33053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case 33055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", Err, objMensagem.iCodigo)

        Case 33057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)

        Case 33059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", Err, objRegiaoVenda.iCodigo)

        Case 33061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrador.iCodigo)

        Case 33063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADA", Err, objPadraoCobranca.iCodigo)

        Case 33065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)

        Case 28944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_REPETIDA_NO_GRID", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174746)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_TipoCliente()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TipoCliente

    Call Limpa_Tela(Me)

    'Limpa os campos que não são limpos pea função Limpa_Tela
    Call Grid_Limpa(objGridCategoria)
    Codigo.Text = ""
    TabelaPreco.Text = ""
    CondicaoPagto.Text = ""
    Mensagem.Text = ""
    Regiao.Text = ""
    Cobrador.Text = ""
    PadraoCobranca.Text = ""
    Transportadora.Text = ""

    Tipo(0).Caption = ""
    Tipo(1).Caption = ""

    Tipos.ListIndex = -1

    iAlterado = 0

    Exit Sub
    
Erro_Limpa_Tela_TipoCliente:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174747)

    End Select

    Exit Sub

End Sub

Function Exibe_Dados_TipoCliente(objTipoCliente As ClassTipoCliente) As Long
'Traz os dados para tela

Dim objEndereco As ClassEndereco
Dim iIndice As Integer
Dim lErro As Long
Dim sContaEnxuta As String
Dim bCancel As Boolean

On Error GoTo Erro_Exibe_Dados_TipoCliente

    'IDENTIFICACAO :

    Codigo.Text = objTipoCliente.iCodigo

    Descricao.Text = objTipoCliente.sDescricao

    Observacao.Text = objTipoCliente.sObservacao

    ContaContabil.PromptInclude = False

    If Len(Trim(objTipoCliente.sContaContabil)) = 0 Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objTipoCliente.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then Error 19076

        ContaContabil.Text = sContaEnxuta
    End If

    ContaContabil.PromptInclude = True

    'Preenche os campos TipoCliente existentes nos frames
    For iIndice = 0 To 1
        Tipo(iIndice).Caption = objTipoCliente.sDescricao
    Next

    'Lê como o tipo de cliente esta associado às categorias
    lErro = CF("TipoDeClienteCategorias_Le", objTipoCliente, objTipoCliente.colCategoriaItem)
    If lErro <> SUCESSO Then Error 28936

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGridCategoria)

    'Exibe os dados da coleção na tela
    For iIndice = 1 To objTipoCliente.colCategoriaItem.Count

        'Insere no Grid Categoria
        GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) = objTipoCliente.colCategoriaItem.Item(iIndice).sCategoria
        GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL) = objTipoCliente.colCategoriaItem.Item(iIndice).sItem

    Next

    objGridCategoria.iLinhasExistentes = objTipoCliente.colCategoriaItem.Count

    'DADOS FINANCEIROS :

    LimiteCredito.Text = CStr(objTipoCliente.dLimiteCredito)

    If objTipoCliente.iTabelaPreco = 0 Then
        TabelaPreco.Text = ""
    Else
        TabelaPreco.Text = CStr(objTipoCliente.iTabelaPreco)
        Call TabelaPreco_Validate(bCancel)
    End If

    If objTipoCliente.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objTipoCliente.iCondicaoPagto)
        CondicaoPagto_Validate (bCancel)
    End If

    If objTipoCliente.iCodMensagem = 0 Then
        Mensagem.Text = ""
    Else
        Mensagem.Text = CStr(objTipoCliente.iCodMensagem)
        Mensagem_Validate (bCancel)
    End If

    Desconto.Text = CStr(100 * objTipoCliente.dDesconto)

    'VENDAS :

    If objTipoCliente.iVendedor = 0 Then
        Vendedor.Text = ""
    Else
        Vendedor.Text = CStr(objTipoCliente.iVendedor)
        Vendedor_Validate (bCancel)
    End If

    If objTipoCliente.dComissaoVendas = 0# Then
        ComissaoVendas.Text = ""
    Else
        ComissaoVendas.Text = CStr(100 * objTipoCliente.dComissaoVendas)
    End If

    If objTipoCliente.iRegiao = 0 Then
        Regiao.Text = ""
    Else
        Regiao.Text = CStr(objTipoCliente.iRegiao)
        Regiao_Validate (bCancel)
    End If

    If objTipoCliente.iFreqVisitas = 0 Then
        FreqVisitas.Text = ""
    Else
        FreqVisitas.Text = CStr(objTipoCliente.iFreqVisitas)
    End If

    If objTipoCliente.iCodCobrador = 0 Then
        Cobrador.Text = ""
    Else
        Cobrador.Text = CStr(objTipoCliente.iCodCobrador)
        Cobrador_Validate (bCancel)
    End If

    If objTipoCliente.iPadraoCobranca = 0 Then
        PadraoCobranca.Text = ""
    Else
        PadraoCobranca.Text = CStr(objTipoCliente.iPadraoCobranca)
        PadraoCobranca_Validate (bCancel)
    End If

    If objTipoCliente.iCodTransportadora = 0 Then
        Transportadora.Text = ""
    Else
        Transportadora.Text = CStr(objTipoCliente.iCodTransportadora)
        Transportadora_Validate (bCancel)
    End If

    iAlterado = 0

    Exibe_Dados_TipoCliente = SUCESSO

Exit Function

Erro_Exibe_Dados_TipoCliente:

    Exibe_Dados_TipoCliente = Err

    Select Case Err

        Case 19076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", objTipoCliente.sContaContabil)

        Case 28936

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174748)

    End Select

    Exit Function

End Function

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_TIPOS_CLIENTE_ID
                
            Case TAB_DadosFinanceiros
                Parent.HelpContextID = IDH_TIPOS_CLIENTE_DADOS_FIN
                        
            Case TAB_Vendas
                Parent.HelpContextID = IDH_TIPOS_CLIENTE_VENDAS
            
        End Select
    
    End If

End Sub

Private Sub ContaContabilLabel_Click()
'chama browse de plano de contas

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 19091

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

    Case 19091

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174749)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagtoLabel_Click()
'chama browse de condicoes de pagto

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagto)

    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondicaoPagto)

End Sub

Private Sub VendedorLabel_Click()
'chama browse de vendedores

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection

    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)

    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub PadraoCobrancaLabel_Click()
'chama browse de tipos de padrao de cobranca

Dim objPadraoCobranca As New ClassPadraoCobranca
Dim colSelecao As New Collection

    If PadraoCobranca.ListIndex <> -1 Then objPadraoCobranca.iCodigo = Codigo_Extrai(PadraoCobranca.Text)

    Call Chama_Tela("PadraoCobrancaLista", colSelecao, objPadraoCobranca, objEventoPadraoCobranca)

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)
'retorno do browse do plano de contas

Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaContabil.Text = ""

    Else

        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 19208

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 19208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174750)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCondicaoPagto_evSelecao(obj1 As Object)
'retorno do browse de condicoes de pagto

Dim objCondicaoPagto As ClassCondicaoPagto
Dim lErro As Long, Cancel As Boolean

On Error GoTo Erro_objEventoCondicaoPagto_evSelecao

    Set objCondicaoPagto = obj1

    CondicaoPagto.Text = CStr(objCondicaoPagto.iCodigo)
    Call CondicaoPagto_Validate(Cancel)

    Me.Show

    Exit Sub

Erro_objEventoCondicaoPagto_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174751)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)
'retorno do browse de vendedores

Dim objVendedor As ClassVendedor
Dim lErro As Long

    Set objVendedor = obj1

    If objVendedor.iCodigo = 0 Then
        Vendedor.Text = ""
    Else
        Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & objVendedor.sNomeReduzido
    End If

    iAlterado = 0

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoPadraoCobranca_evSelecao(obj1 As Object)
'retorno do browse de tipos de padrao de cobranca

Dim objPadraoCobranca As ClassPadraoCobranca
Dim lErro As Long
Dim bCancel As Boolean

    Set objPadraoCobranca = obj1

    If objPadraoCobranca.iCodigo = 0 Then
        PadraoCobranca.Text = ""
    Else
        PadraoCobranca.Text = CStr(objPadraoCobranca.iCodigo)
        Call PadraoCobranca_Validate(bCancel)
    End If
    
    iAlterado = 0

    Me.Show

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposDeCliente"

    'Le os dados da Tela TipoCliente
    lErro = Move_Tela_Memoria(objTipoCliente)
    If lErro <> SUCESSO Then Error 19217

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "Codigo", objTipoCliente.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objTipoCliente.sDescricao, STRING_TIPO_CLIENTE_DESCRICAO, "Descricao"
    colCampoValor.Add "LimiteCredito", objTipoCliente.dLimiteCredito, 0, "LimiteCredito"
    colCampoValor.Add "CondicaoPagto", objTipoCliente.iCondicaoPagto, 0, "Condicaopagto"
    colCampoValor.Add "Desconto", objTipoCliente.dDesconto, 0, "Desconto"
    colCampoValor.Add "CodMensagem", objTipoCliente.iCodMensagem, 0, "CodMensagem"
    colCampoValor.Add "TabelaPreco", objTipoCliente.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "Observacao", objTipoCliente.sObservacao, STRING_TIPO_CLIENTE_OBS, "Observacao"
    colCampoValor.Add "ContaContabil", objTipoCliente.sContaContabil, STRING_CONTA, "ContaContabil"
    colCampoValor.Add "Vendedor", objTipoCliente.iVendedor, 0, "Vendedor"
    colCampoValor.Add "ComissaoVendas", objTipoCliente.dComissaoVendas, 0, "ComissaoVendas"
    colCampoValor.Add "Regiao", objTipoCliente.iRegiao, 0, "Regiao"
    colCampoValor.Add "FreqVisitas", objTipoCliente.iFreqVisitas, 0, "FreqVisitas"
    colCampoValor.Add "CodTransportadora", objTipoCliente.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodCobrador", objTipoCliente.iCodCobrador, 0, "CodCobrador"
    colCampoValor.Add "PadraoCobranca", objTipoCliente.iPadraoCobranca, 0, "PadraoCobranca"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 19217

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174752)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Tela_Preenche

    objTipoCliente.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTipoCliente.iCodigo <> 0 Then

        objTipoCliente.sDescricao = colCampoValor.Item("Descricao").vValor
        objTipoCliente.dLimiteCredito = colCampoValor.Item("LimiteCredito").vValor
        objTipoCliente.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        objTipoCliente.dDesconto = colCampoValor.Item("Desconto").vValor
        objTipoCliente.iCodMensagem = colCampoValor.Item("CodMensagem").vValor
        objTipoCliente.iTabelaPreco = colCampoValor.Item("TabelaPreco").vValor
        objTipoCliente.sObservacao = colCampoValor.Item("Observacao").vValor
        objTipoCliente.sContaContabil = colCampoValor.Item("ContaContabil").vValor
        objTipoCliente.iVendedor = colCampoValor.Item("Vendedor").vValor
        objTipoCliente.dComissaoVendas = colCampoValor.Item("ComissaoVendas").vValor
        objTipoCliente.iRegiao = colCampoValor.Item("Regiao").vValor
        objTipoCliente.iFreqVisitas = colCampoValor.Item("FreqVisitas").vValor
        objTipoCliente.iCodTransportadora = colCampoValor.Item("CodTransportadora").vValor
        objTipoCliente.iCodCobrador = colCampoValor.Item("CodCobrador").vValor
        objTipoCliente.iPadraoCobranca = colCampoValor.Item("PadraoCobranca").vValor

        lErro = Exibe_Dados_TipoCliente(objTipoCliente)
        If lErro <> SUCESSO Then Error 19218

        Tipos.ListIndex = -1

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 19218

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174753)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

 Dim lErro As Long

    Set objGridCategoria = Nothing

    Set objEventoContaContabil = Nothing
    Set objEventoCondicaoPagto = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoPadraoCobranca = Nothing

    Set objEventoTransportadora = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Carrega_ComboCategoriaCliente() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaCliente_Le_Todos", colCategorias)
    If lErro <> SUCESSO Then Error 28917

    For Each objCategoriaCliente In colCategorias

        'Insere na combo CategoriaCliente
        ComboCategoriaCliente.AddItem objCategoriaCliente.sCategoria

    Next

    Carrega_ComboCategoriaCliente = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaCliente:

    Carrega_ComboCategoriaCliente = Err

    Select Case Err

        Case 28917

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174754)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboCategoriaClienteItem(objCategoriaCliente As ClassCategoriaCliente) As Long
'Carrega o Item da Categoria na Combobox

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaClienteItem As ClassCategoriaClienteItem

On Error GoTo Erro_Carrega_ComboCategoriaClienteItem

    'Lê a tabela CategoriaClienteItem a partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO Then Error 28918

    'Insere na combo CategoriaClienteItem
    For Each objCategoriaClienteItem In colItensCategoria

        'Insere na combo CategoriaCliente
        ComboCategoriaClienteItem.AddItem objCategoriaClienteItem.sItem

    Next

    Carrega_ComboCategoriaClienteItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaClienteItem:

    Carrega_ComboCategoriaClienteItem = Err

    Select Case Err

        Case 28918

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174755)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long

    'Tela em questão
    Set objGridInt.objForm = Me

    'Títulos do Grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

    'Campos de edição do Grid
    objGridInt.colCampo.Add (ComboCategoriaCliente.Name)
    objGridInt.colCampo.Add (ComboCategoriaClienteItem.Name)

    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 51

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    GridCategoria.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridCategoria.Col

            Case GRID_CATEGORIA_COL

                lErro = Saida_Celula_Categoria(objGridInt)
                If lErro <> SUCESSO Then Error 28919

            Case GRID_VALOR_COL

                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then Error 28920

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 28921

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 28919, 28920

        Case 28921
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174756)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Categoria do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaCliente

    If Len(Trim(ComboCategoriaCliente.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaCliente)
        If lErro <> SUCESSO Then
    
            'Preenche o objeto com a Categoria
            objCategoriaCliente.sCategoria = ComboCategoriaCliente.Text

            'Lê Categoria De Cliente no BD
            lErro = CF("CategoriaCliente_Le", objCategoriaCliente)
            If lErro <> SUCESSO And lErro <> 28847 Then Error 28922

            'Categoria não está cadastrada
            If lErro = 28847 Then Error 28923
    
        End If
    
        'Verifica se já existe a categoria no Grid
        For iIndice = 1 To objGridCategoria.iLinhasExistentes
            If iIndice <> GridCategoria.Row Then If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) > 0 And GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) = ComboCategoriaCliente.Text Then Error 28924
        Next

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 28925

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = Err

    Select Case Err

        Case 28922
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 28923 'Categoria não está cadastrada

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTE")

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a Tela "CategoriaCliente"
                Call Chama_Tela("CategoriaCliente", objCategoriaCliente)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 28924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 28925
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174757)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim colItens As New Collection
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaClienteItem

    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(ComboCategoriaClienteItem)
    If lErro <> SUCESSO Then

        If Len(Trim(ComboCategoriaClienteItem.Text)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaClienteItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaClienteItem.sItem = ComboCategoriaClienteItem.Text

            'Lê Item De Categoria De Cliente no BD
            lErro = CF("CategoriaClienteItem_Le", objCategoriaClienteItem)
            If lErro <> SUCESSO And lErro <> 28991 Then Error 28930

            'Item da Categoria não está cadastrado
            If lErro = 28991 Then Error 28931

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 28932

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 28930
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 28931 'Item da Categoria não está cadastrado

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTEITEM")

            If vbMsgRes = vbYes Then

                'Preenche o objeto com a Categoria
                objCategoriaCliente.sCategoria = ComboCategoriaCliente.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a Tela "CategoriaCliente"
                Call Chama_Tela("CategoriaCliente", objCategoriaCliente)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 28932
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174758)

    End Select

    Exit Function

End Function

Private Sub Trata_ComboCategoriaClienteItem()

Dim lErro As Long
Dim iIndice As Integer, sValor As String
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Trata_ComboCategoriaClienteItem

    sValor = ComboCategoriaClienteItem.Text

    ComboCategoriaClienteItem.Clear

    ComboCategoriaClienteItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(Trim(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL))) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaCliente.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)

        lErro = Carrega_ComboCategoriaClienteItem(objCategoriaCliente)
        If lErro <> SUCESSO Then Error 28934

    End If

    For iIndice = 0 To ComboCategoriaClienteItem.ListCount - 1
        If ComboCategoriaClienteItem.List(iIndice) = GridCategoria.Text Then
            ComboCategoriaClienteItem.ListIndex = iIndice
            Exit For
        End If
    Next

    Exit Sub

Erro_Trata_ComboCategoriaClienteItem:

    Select Case Err

        Case 28934

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174759)

    End Select

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TIPOS_CLIENTE_ID
    Set Form_Load_Ocx = Me
    Caption = "Tipos de Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoCliente"
    
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
        
        If Me.ActiveControl Is CondicaoPagto Then
            Call CondicaoPagtoLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        ElseIf Me.ActiveControl Is PadraoCobranca Then
            Call PadraoCobrancaLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        End If
    
    End If
    
End Sub


Private Sub Tipo_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Tipo(Index), Source, X, Y)
End Sub

Private Sub Tipo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Tipo(Index), Button, Shift, X, Y)
End Sub


Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
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

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub Label47_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label47, Source, X, Y)
End Sub

Private Sub Label47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label47, Button, Shift, X, Y)
End Sub

Private Sub Label44_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label44, Source, X, Y)
End Sub

Private Sub Label44_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label44, Button, Shift, X, Y)
End Sub

Private Sub VendedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(VendedorLabel, Source, X, Y)
End Sub

Private Sub VendedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VendedorLabel, Button, Shift, X, Y)
End Sub

Private Sub PadraoCobrancaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PadraoCobrancaLabel, Source, X, Y)
End Sub

Private Sub PadraoCobrancaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PadraoCobrancaLabel, Button, Shift, X, Y)
End Sub

Private Sub Label42_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label42, Source, X, Y)
End Sub

Private Sub Label42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label42, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagtoLabel, Source, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

